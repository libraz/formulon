//
// Spill reconciliation for `RecalcEngine`: compact-range trace queries,
// spill-footprint dependency replacement, dynamic-read reconciliation and
// the spill-release dirty wave.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <unordered_set>
#include <vector>

#include "cell.h"
#include "eval/cell_evaluator.h"
#include "eval/dep_extractor.h"
#include "eval/dep_graph.h"
#include "eval/dynamic_read_log.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/recalc_reentry.h"
#include "eval/spill_release.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/rect_iterator.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

bool spill_intersects_range(const SpillFootprint& footprint, const CellRangeDependency& range) noexcept {
  // Both records use inclusive endpoints / positive dimensions. Convert to
  // half-open uint64 intervals before comparing so a max-grid coordinate or
  // a malformed external record cannot wrap an end point.
  const std::uint64_t spill_row_end = static_cast<std::uint64_t>(footprint.anchor_row) + footprint.rows;
  const std::uint64_t spill_col_end = static_cast<std::uint64_t>(footprint.anchor_col) + footprint.cols;
  const std::uint64_t range_row_end = static_cast<std::uint64_t>(range.row_last) + 1U;
  const std::uint64_t range_col_end = static_cast<std::uint64_t>(range.col_last) + 1U;
  return footprint.rows != 0U && footprint.cols != 0U &&
         static_cast<std::uint64_t>(footprint.anchor_row) < range_row_end &&
         static_cast<std::uint64_t>(range.row_first) < spill_row_end &&
         static_cast<std::uint64_t>(footprint.anchor_col) < range_col_end &&
         static_cast<std::uint64_t>(range.col_first) < spill_col_end;
}

}  // namespace

namespace {

// Appends every populated coordinate of `range` to `out`. `extent` is the
// rectangle's populated bounding box as reported by `Sheet::populated_extent`.
// Walk the stored rows when that is the smaller set, matching Sheet's range
// scanners, so a sparse rectangle does not pay for every row in its extent.
void collect_populated_cells(const Sheet& sheet, const CellRangeDependency& range, const Sheet::PopulatedExtent& extent,
                             std::vector<CellNodeId>& out) {
  const auto scan_row = [&](std::uint32_t row, const RowCells& cells) {
    for (const auto& [col, cell] :
         cells.entries_in(extent.first_col, static_cast<std::uint64_t>(extent.last_col) + 1U)) {
      if (!cell.formula_text.empty() || !cell.cached_value.is_blank()) {
        out.push_back(CellNodeId{range.sheet_id, row, col});
      }
    }
  };
  const auto& rows = sheet.rows();
  const std::uint64_t extent_rows = static_cast<std::uint64_t>(extent.last_row) - extent.first_row + 1U;
  if (static_cast<std::uint64_t>(rows.size()) <= extent_rows) {
    for (const auto& [row, cells] : rows) {
      if (row >= extent.first_row && row <= extent.last_row) {
        scan_row(row, cells);
      }
    }
    return;
  }
  for (std::uint64_t row = extent.first_row; row <= extent.last_row; ++row) {
    const auto it = rows.find(static_cast<std::uint32_t>(row));
    if (it != rows.end()) {
      scan_row(static_cast<std::uint32_t>(row), it->second);
    }
  }
}

}  // namespace

std::vector<CellNodeId> RecalcEngine::compact_range_precedents_of(CellNodeId cell, const Workbook& workbook) const {
  std::lock_guard<std::mutex> guard(mutex_);
  std::vector<CellNodeId> precedents;
  range_dependencies_.for_each_range_of_owner(cell, [&](const CellRangeDependency& range) {
    if (range.sheet_id >= workbook.sheet_count()) {
      return;
    }
    const Sheet& sheet = workbook.sheet(range.sheet_id);
    const auto extent = sheet.populated_extent(range.row_first, range.col_first, range.row_last, range.col_last);
    if (!extent.has_value()) {
      return;
    }
    collect_populated_cells(sheet, range, *extent, precedents);
  });
  // A cell can be reached through two overlapping rectangles of the same
  // formula; the trace surface reports each precedent once.
  std::sort(precedents.begin(), precedents.end(), CellNodeIdOrder{});
  precedents.erase(std::unique(precedents.begin(), precedents.end()), precedents.end());
  return precedents;
}

std::vector<CellNodeId> RecalcEngine::compact_range_dependents_of(CellNodeId cell) const {
  std::lock_guard<std::mutex> guard(mutex_);
  std::vector<CellNodeId> owners;
  std::unordered_set<CellNodeId, CellNodeIdHash> seen;
  range_dependencies_.for_each_owner_covering(cell, [&owners, &seen](CellNodeId owner) {
    if (seen.insert(owner).second) {
      owners.push_back(owner);
    }
  });
  return owners;
}

DepGraph::DependencyDelta RecalcEngine::reconcile_spill_dependencies_locked(const Workbook& workbook) {
  // With no compact watchers and no old derived ownership there is nothing
  // to snapshot or replace. This is the common scalar-only fast path.
  if (range_dependencies_.empty() && referenced_cells_.empty() &&
      !graph_.has_source_edges(DepGraph::DependencySource::kSpillFootprint)) {
    return {};
  }

  // Copy every sheet's committed geometry before touching the graph. This
  // keeps the lock order engine -> sheet/spill and ensures graph replacement
  // never holds a Sheet spill mutex while another graph operation runs.
  std::vector<std::vector<SpillFootprint>> committed(workbook.sheet_count());
  std::vector<bool> referenced_sheets(workbook.sheet_count(), false);
  range_dependencies_.for_each_distinct_range(
      [&](std::uint32_t, const CellRangeDependency& range, const std::vector<CellNodeId>&) {
        if (range.sheet_id < workbook.sheet_count()) {
          referenced_sheets[range.sheet_id] = true;
        }
      });
  bool any_spill = false;
  for (std::uint16_t sheet_id = 0; sheet_id < workbook.sheet_count(); ++sheet_id) {
    if (!referenced_sheets[sheet_id] && !referenced_cells_.references_sheet(sheet_id)) {
      continue;
    }
    committed[sheet_id] = workbook.sheet(sheet_id).committed_spill_footprints();
    any_spill = any_spill || !committed[sheet_id].empty();
  }
  if (!any_spill && !graph_.has_source_edges(DepGraph::DependencySource::kSpillFootprint)) {
    return {};
  }

  std::vector<DepGraph::Edge> desired;
  range_dependencies_.for_each_distinct_range(
      [&](std::uint32_t, const CellRangeDependency& range, const std::vector<CellNodeId>& owners) {
        if (range.sheet_id >= committed.size()) {
          return;
        }
        for (const SpillFootprint& footprint : committed[range.sheet_id]) {
          if (!spill_intersects_range(footprint, range)) {
            continue;
          }
          const CellNodeId producer{range.sheet_id, footprint.anchor_row, footprint.anchor_col};
          for (const CellNodeId owner : owners) {
            // A self edge is retained: when a formula's compact range really
            // reads its own committed spill, that is a genuine circular
            // dependency and must reach the SCC/cycle handling path rather
            // than being silently treated as an ordering artefact.
            desired.emplace_back(owner, producer);
          }
        }
      });
  // A spilled cell a formula reads holds its anchor's value, so it depends
  // on the anchor: its readers are ordered behind the anchor, dirtied with
  // it, and a read back into the anchor's own inputs is a cycle.
  for (std::uint16_t sheet_id = 0; sheet_id < committed.size(); ++sheet_id) {
    for (const SpillFootprint& footprint : committed[sheet_id]) {
      if (footprint.rows == 0U || footprint.cols == 0U) {
        continue;
      }
      const CellNodeId anchor{sheet_id, footprint.anchor_row, footprint.anchor_col};
      referenced_cells_.for_each_in(
          sheet_id, footprint.anchor_row, footprint.anchor_col, footprint.anchor_row + footprint.rows - 1U,
          footprint.anchor_col + footprint.cols - 1U, [&](std::uint32_t row, std::uint32_t col) {
            if (row != anchor.row || col != anchor.col) {
              desired.emplace_back(CellNodeId{sheet_id, row, col}, anchor);
            }
          });
    }
  }
  return graph_.replace_dependencies(DepGraph::DependencySource::kSpillFootprint, desired);
}

std::vector<CellNodeId> RecalcEngine::reconcile_dynamic_reads_locked(
    const Workbook& workbook, const DynamicReadLog& log, const std::unordered_set<CellNodeId, CellNodeIdHash>* closure,
    std::unordered_set<CellNodeId, CellNodeIdHash>& refreshed) {
  for (const auto& [reader, ordinal] : log.readers()) {
    (void)ordinal;
    // Edges accumulate across the waves of one recalc call, so a reader that
    // is retried keeps ordering its targets in later waves.
    if (refreshed.insert(reader).second) {
      graph_.clear_dynamic_dependencies_of(reader);
    }
  }
  std::vector<CellNodeId> stale;
  std::unordered_set<CellNodeId, CellNodeIdHash> stale_seen;
  std::vector<std::vector<SpillFootprint>> footprints(workbook.sheet_count());
  std::vector<char> footprints_loaded(workbook.sheet_count(), 0);
  for (const DynamicRead& read : log.reads()) {
    const std::size_t sheet_id = read.sheet_id;
    if (sheet_id >= workbook.sheet_count()) {
      continue;
    }
    const Sheet& sheet = workbook.sheet(sheet_id);
    const auto target_sheet = static_cast<std::uint16_t>(sheet_id);
    std::vector<CellNodeId> targets;
    for (const CellAddress& a :
         sheet.formula_cells_in(read.rect.row_first, read.rect.col_first, read.rect.row_last, read.rect.col_last)) {
      targets.push_back(CellNodeId{target_sheet, a.row, a.col});
    }
    // A phantom cell of a committed spill reads its anchor's result.
    if (footprints_loaded[sheet_id] == 0) {
      footprints[sheet_id] = sheet.committed_spill_footprints();
      footprints_loaded[sheet_id] = 1;
    }
    const CellRangeDependency range{target_sheet, read.rect.row_first, read.rect.row_last, read.rect.col_first,
                                    read.rect.col_last};
    for (const SpillFootprint& footprint : footprints[sheet_id]) {
      if (spill_intersects_range(footprint, range)) {
        targets.push_back(CellNodeId{target_sheet, footprint.anchor_row, footprint.anchor_col});
      }
    }
    const std::optional<std::uint64_t> reader_component = log.iterative_component(read.reader);
    for (const CellNodeId target : targets) {
      graph_.add_dynamic_dependency(read.reader, target);
      if (stale_seen.count(read.reader) != 0U) {
        continue;
      }
      if (reader_component && log.iterative_component(target) == reader_component) {
        continue;
      }
      const std::optional<std::uint64_t> committed = log.commit_ordinal(target);
      const bool read_too_early = committed && *committed >= read.ordinal;
      const bool left_dirty = closure != nullptr && dirty_.contains(target) && !committed;
      if (read_too_early || left_dirty) {
        stale_seen.insert(read.reader);
        stale.push_back(read.reader);
      }
    }
  }
  std::sort(stale.begin(), stale.end(), CellNodeIdOrder{});
  return stale;
}

void RecalcEngine::mark_spill_release_wave_locked(const std::vector<CellNodeId>& anchors) {
  for (const CellNodeId anchor : anchors) {
    mark_dirty_locked(anchor);
    for (const CellNodeId dependent : graph_.dependents_of_ref(anchor)) {
      mark_dirty_locked(dependent);
    }
    mark_range_dependents_dirty_locked(anchor);
  }
}

}  // namespace formulon::eval
