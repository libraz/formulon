//
// Implementation of `RecalcEngine`. See `recalc_engine.h` for the public
// contract.

#include "eval/recalc_engine.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/builtin_names.h"
#include "eval/cell_evaluator.h"
#include "eval/dep_extractor.h"
#include "eval/dep_graph.h"
#include "eval/dirty_set.h"
#include "eval/dynamic_read_log.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_reentry.h"
#include "eval/spill_release.h"
#include "eval/volatile_tracker.h"
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

// The spill-release queue, the progress snapshot and the wave ceilings are
// shared with the parallel scheduler; see `spill_release.h`.
using detail::queue_spill_release;
using detail::SpillReleaseQueue;
using detail::SpillWaveBudget;

void mark_spill_release_wave(const RecalcEngine::LockedMutator& mutator, const std::vector<CellNodeId>& anchors,
                             const DepGraph& graph) {
  for (const CellNodeId anchor : anchors) {
    mutator.mark_dirty(anchor);
    for (const CellNodeId dependent : graph.dependents_of_ref(anchor)) {
      mutator.mark_dirty(dependent);
    }
    mutator.mark_range_dependents_dirty(anchor);
  }
}

// Drops the virtual node of every compact rectangle that lost its last watcher.
void remove_range_nodes(DepGraph& graph, const std::vector<std::uint32_t>& released_range_ids) {
  for (const std::uint32_t range_id : released_range_ids) {
    graph.remove_node(range_node(range_id));
  }
}

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

// Largest runtime read rectangle whose open cells are checked one by one before admitting a producer.
constexpr std::uint64_t kMaxCheckedSpillReadCells = 256U;

}  // namespace

// ----------------------------------------------------------------------------
// RecalcEngine
// ----------------------------------------------------------------------------

RecalcEngine::RecalcEngine() : arena_(std::make_unique<Arena>(/*initial_chunk_bytes=*/4096, kMaxEvalArenaBytes)) {}
RecalcEngine::~RecalcEngine() = default;

// ---------------------------------------------------------------------------
// LockedMutator — passkey facade routing `Workbook`'s compound mutators
// through the engine's `*_locked` API. The facade never acquires the
// mutex; the caller (a `Workbook` member) holds `mutex_` for the entire
// scope.
// ---------------------------------------------------------------------------

void RecalcEngine::LockedMutator::register_formula(CellNodeId cell, const parser::AstNode& ast,
                                                   const Workbook& workbook) const {
  engine_.register_formula_locked(cell, ast, workbook);
}

void RecalcEngine::LockedMutator::unregister_formula(CellNodeId cell) const {
  engine_.unregister_formula_locked(cell);
}

void RecalcEngine::LockedMutator::clear_cell_dependencies(CellNodeId cell) const {
  engine_.clear_cell_dependencies_locked(cell);
}

void RecalcEngine::LockedMutator::mark_dirty(CellNodeId cell) const {
  engine_.mark_dirty_locked(cell);
}

void RecalcEngine::LockedMutator::mark_range_dependents_dirty(CellNodeId cell) const {
  engine_.mark_range_dependents_dirty_locked(cell);
}

void RecalcEngine::LockedMutator::reset_graph() const {
  engine_.reset_graph_locked();
}

const DepGraph& RecalcEngine::LockedMutator::dep_graph() const noexcept {
  return engine_.graph_;
}

bool RecalcEngine::LockedMutator::has_ever_registered_formula() const noexcept {
  return engine_.has_ever_registered_formula_;
}

namespace {

// Appends every populated coordinate of `range` to `out`. `extent` is the
// rectangle's populated bounding box as reported by `Sheet::populated_extent`,
// so the walk visits only rows that can carry content.
void collect_populated_cells(const Sheet& sheet, const CellRangeDependency& range, const Sheet::PopulatedExtent& extent,
                             std::vector<CellNodeId>& out) {
  for (std::uint32_t row = extent.first_row; row <= extent.last_row; ++row) {
    const auto it = sheet.rows().find(row);
    if (it == sheet.rows().end()) {
      continue;
    }
    const RowCells& cells = it->second;
    const std::size_t begin = std::max<std::size_t>(extent.first_col, cells.first_col());
    const std::size_t end = std::min<std::size_t>(static_cast<std::size_t>(extent.last_col) + 1U, cells.size());
    for (std::size_t col = begin; col < end; ++col) {
      const Cell& cell = cells[col];
      if (!cell.formula_text.empty() || !cell.cached_value.is_blank()) {
        out.push_back(CellNodeId{range.sheet_id, row, static_cast<std::uint32_t>(col)});
      }
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
  range_dependencies_.for_each_owner_covering(cell, [&owners](CellNodeId owner) {
    if (std::find(owners.begin(), owners.end(), owner) == owners.end()) {
      owners.push_back(owner);
    }
  });
  return owners;
}

std::vector<CellNodeId> RecalcEngine::LockedMutator::three_d_span_owners_covering_sheet(
    std::uint16_t edited_sheet) const {
  std::vector<CellNodeId> owners;
  for (const RegisteredThreeDSpan& entry : engine_.three_d_span_dependencies_) {
    if (edited_sheet < entry.span.sheet_first || edited_sheet > entry.span.sheet_last) {
      continue;
    }
    if (std::find(owners.begin(), owners.end(), entry.owner) == owners.end()) {
      owners.push_back(entry.owner);
    }
  }
  std::sort(owners.begin(), owners.end(), CellNodeIdOrder{});
  return owners;
}

void RecalcEngine::ReferencedCellIndex::add(CellNodeId cell) {
  if (cell.sheet_id >= by_sheet_.size()) {
    by_sheet_.resize(static_cast<std::size_t>(cell.sheet_id) + 1U);
  }
  ++by_sheet_[cell.sheet_id][key(cell.row, cell.col)];
}

void RecalcEngine::ReferencedCellIndex::remove(CellNodeId cell) {
  if (cell.sheet_id >= by_sheet_.size()) {
    return;
  }
  std::map<std::uint64_t, std::uint32_t>& cells = by_sheet_[cell.sheet_id];
  const auto found = cells.find(key(cell.row, cell.col));
  if (found != cells.end() && --found->second == 0U) {
    cells.erase(found);
  }
}

bool RecalcEngine::ReferencedCellIndex::empty() const noexcept {
  return std::all_of(by_sheet_.begin(), by_sheet_.end(), [](const auto& cells) { return cells.empty(); });
}

void RecalcEngine::forget_referenced_cells_locked(CellNodeId cell) {
  for (const CellNodeId dep : graph_.dependencies_of_ref(cell)) {
    if (!is_range_node(dep) && graph_.has_dependency_source(cell, dep, DepGraph::DependencySource::kAuthored)) {
      referenced_cells_.remove(dep);
    }
  }
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
  std::unordered_set<std::uint16_t> referenced_sheets;
  referenced_sheets.reserve(range_dependencies_.distinct_range_count());
  range_dependencies_.for_each_distinct_range(
      [&](std::uint32_t, const CellRangeDependency& range, const std::vector<CellNodeId>&) {
        if (range.sheet_id < workbook.sheet_count()) {
          referenced_sheets.insert(range.sheet_id);
        }
      });
  for (std::uint16_t sheet_id = 0; sheet_id < workbook.sheet_count(); ++sheet_id) {
    if (referenced_cells_.references_sheet(sheet_id)) {
      referenced_sheets.insert(sheet_id);
    }
  }
  bool any_spill = false;
  for (std::uint16_t sheet_id : referenced_sheets) {
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

RecalcEngine::DynamicReadPass::DynamicReadPass(RecalcEngine& engine, const Workbook& workbook)
    : engine_(engine), workbook_(workbook), log_(workbook, engine.volatiles_, engine.graph_) {}

void RecalcEngine::DynamicReadPass::settle_sccs(std::vector<std::vector<CellNodeId>>& sccs,
                                                const std::unordered_set<CellNodeId, CellNodeIdHash>& nodes) {
  if (!first_wave_) {
    return;
  }
  // Before anything is evaluated: a cycle through a dynamic read is left at
  // these values whichever path, serial or pooled, evaluates its members.
  log_.keep_prior_values_fed_by_readers(nodes);
  if (engine_.drop_dynamic_only_cycles_locked(sccs)) {
    sccs = engine_.graph_.tarjan_scc_subset(nodes);
  }
}

const std::vector<CellNodeId>& RecalcEngine::DynamicReadPass::end_wave(
    const std::unordered_set<CellNodeId, CellNodeIdHash>* closure, const FunctionRegistry* registry) {
  stale_ = engine_.reconcile_dynamic_reads_locked(workbook_, log_, closure, refreshed_);
  if (closure != nullptr && registry != nullptr) {
    discover_pending_candidates(*registry, *closure);
  }
  log_.clear_wave();
  first_wave_ = false;
  std::sort(stale_.begin(), stale_.end(), CellNodeIdOrder{});
  stale_.erase(std::unique(stale_.begin(), stale_.end()), stale_.end());
  return stale_;
}

void RecalcEngine::DynamicReadPass::discover_pending_candidates(
    const FunctionRegistry& registry, const std::unordered_set<CellNodeId, CellNodeIdHash>& closure) {
  Arena potential_arena;
  for (const DynamicRead& read : log_.reads()) {
    if (read.sheet_id >= workbook_.sheet_count() || !engine_.potential_spill_producers_.has_sheet(read.sheet_id)) {
      continue;
    }

    // Discovery is only for a blank coordinate whose producer has not committed. A small rectangle is checked cell
    // by cell; a larger one keeps the geometric test alone.
    const Sheet& read_sheet = workbook_.sheet(static_cast<std::uint16_t>(read.sheet_id));
    const bool checks_targets =
        static_cast<std::uint64_t>(read.rect.rows()) * read.rect.cols() <= kMaxCheckedSpillReadCells;
    std::vector<SpillTarget> open_targets;
    if (checks_targets) {
      for (std::uint32_t row = read.rect.row_first; row <= read.rect.row_last; ++row) {
        for (std::uint32_t col = read.rect.col_first; col <= read.rect.col_last; ++col) {
          Sheet::CellRead target;
          read_sheet.read_formula_cell(row, col, target);
          if (!target.is_formula() && target.value().is_blank()) {
            open_targets.emplace_back(read_sheet, row, col);
          }
        }
      }
      if (open_targets.empty()) {
        continue;
      }
    }

    CandidateSet* pending = nullptr;
    CandidateSet* attempted = nullptr;
    const auto attempted_it = attempted_candidates_.find(read.reader);
    if (attempted_it != attempted_candidates_.end()) {
      attempted = &attempted_it->second;
    }
    for (const auto& [producer, static_potential] : engine_.potential_spill_producers_.on_sheet(read.sheet_id)) {
      // A spill extends only down and right, so an anchor past either last coordinate cannot intersect the read.
      if (producer.row > read.rect.row_last || producer.col > read.rect.col_last) {
        continue;
      }
      const CellNodeId candidate = producer;
      if (checks_targets && std::none_of(open_targets.begin(), open_targets.end(),
                                         [candidate](const SpillTarget& t) { return t.reachable_from(candidate); })) {
        continue;
      }
      if (log_.commit_ordinal(producer).has_value()) {
        // A producer committed in this wave is covered by spill-footprint reconciliation.
        continue;
      }
      if (closure.count(producer) != 0U) {
        // A producer already in the closure was expanded by the ordinary BFS.
        continue;
      }
      if (attempted != nullptr && attempted->count(producer) != 0U) {
        continue;
      }
      if (!SpillProducerIndex::admissible(workbook_, registry, producer, static_potential, potential_arena)) {
        continue;
      }
      if (pending == nullptr) {
        pending = &pending_candidates_[read.reader];
      }
      if (pending->insert(producer).second) {
        stale_.push_back(read.reader);
      }
    }
    // Drop an empty entry so the next wave does not visit it.
    if (pending != nullptr && pending->empty()) {
      pending_candidates_.erase(read.reader);
    }
  }
}

bool RecalcEngine::DynamicReadPass::seed_pending_candidates(const FunctionRegistry& registry,
                                                            std::unordered_set<CellNodeId, CellNodeIdHash>& closure,
                                                            std::vector<CellNodeId>& bfs_queue) {
  bool expanded = false;
  Arena potential_arena;
  for (auto pending_it = pending_candidates_.begin(); pending_it != pending_candidates_.end();) {
    const CellNodeId reader = pending_it->first;
    if (closure.count(reader) == 0U) {
      ++pending_it;
      continue;
    }
    CandidateSet& attempted = attempted_candidates_[reader];
    for (const CellNodeId producer : pending_it->second) {
      attempted.insert(producer);
      const SpillPotential* candidate_potential = engine_.potential_spill_producers_.find(producer);
      if (candidate_potential == nullptr) {
        continue;
      }
      if (!SpillProducerIndex::admissible(workbook_, registry, producer, *candidate_potential, potential_arena)) {
        continue;
      }
      if (closure.insert(producer).second) {
        bfs_queue.push_back(producer);
        expanded = true;
      }
    }
    pending_it = pending_candidates_.erase(pending_it);
  }

  // An attempted pair remains a temporary seed for every later wave of this
  // partial call. A producer can be clean when first admitted, then become
  // dirty after a compact-range watcher learns its spill footprint. Keeping
  // the seed lets the next closure expand through that producer's authored
  // dependencies without rediscovering the pair or adding a speculative
  // graph edge. Revalidation removes candidates that were rewritten away.
  for (auto attempted_it = attempted_candidates_.begin(); attempted_it != attempted_candidates_.end();) {
    if (closure.count(attempted_it->first) == 0U) {
      ++attempted_it;
      continue;
    }
    CandidateSet& attempted = attempted_it->second;
    for (auto producer_it = attempted.begin(); producer_it != attempted.end();) {
      const CellNodeId producer = *producer_it;
      const SpillPotential* candidate_potential = engine_.potential_spill_producers_.find(producer);
      if (candidate_potential == nullptr ||
          !SpillProducerIndex::admissible(workbook_, registry, producer, *candidate_potential, potential_arena)) {
        producer_it = attempted.erase(producer_it);
        continue;
      }
      if (closure.insert(producer).second) {
        bfs_queue.push_back(producer);
        expanded = true;
      }
      ++producer_it;
    }
    if (attempted.empty()) {
      attempted_it = attempted_candidates_.erase(attempted_it);
    } else {
      ++attempted_it;
    }
  }
  return expanded;
}

void RecalcEngine::DynamicReadPass::mark_stale_dirty() {
  for (const CellNodeId reader : stale_) {
    engine_.dirty_.mark(reader);
  }
}

void RecalcEngine::commit_unresolved_cycle_locked(const SerialEvalPass& pass,
                                                  const std::vector<CellNodeId>& component) {
  // Excel shows a warning and leaves the cells as they were; with no UI to
  // host that banner Formulon surfaces #REF!, keeping Excel's last value
  // only for a cycle through an OFFSET / INDIRECT read.
  const bool dynamic_cycle = closes_through_dynamic_edge(component, graph_);
  const std::uint64_t ordinal = pass.dynamic.next_ordinal();
  if (dynamic_cycle) {
    pass.dynamic.log().restore_prior_values(pass.workbook, component);
  }
  for (const CellNodeId c : component) {
    if (c.sheet_id >= pass.workbook.sheet_count()) {
      continue;  // A virtual range node.
    }
    if (!dynamic_cycle) {
      Sheet& sheet = pass.workbook.sheet(c.sheet_id);
      SpillCommitter committer(&sheet, c.row, c.col, pass.release_callback, pass.release_user_data);
      sheet.set_cell_cached_value(c.row, c.col, committer.commit(Value::error(ErrorCode::Ref)));
    }
    pass.dynamic.log().note_commit(c, ordinal);
    ++pass.stats.cycle_cells;
  }
}

Expected<void, Error> RecalcEngine::evaluate_formula_cell_locked(const SerialEvalPass& pass, CellNodeId cell) {
  if (cell.sheet_id >= pass.workbook.sheet_count()) {
    return Expected<void, Error>::Ok();
  }
  Sheet& sheet = pass.workbook.sheet(cell.sheet_id);
  Cell staged;
  if (!stage_formula_cell(sheet, cell.row, cell.col, staged)) {
    // The dep graph may carry pure-input cells (read-only literals that
    // someone reads via `add_dependency`). They have nothing to evaluate.
    return Expected<void, Error>::Ok();
  }
  // Reset the per-pass arena before each evaluate so the bump allocator
  // does not grow without bound across cells. Both result shapes
  // already survive this reset:
  //   * Array results are deep-copied into sheet-owned storage by
  //     `Sheet::commit_spill` (driven by
  //     `EvalContext::dispatch_array_result`), keeping the spill table
  //     independent of the arena.
  //   * Scalar Text results are deep-copied into the destination cell's
  //     `Cell::cached_text_owned` by `Sheet::set_cell_cached_value` on
  //     the write below, so the cached `string_view` does not dangle
  //     when the next cell's evaluation resets the arena.
  arena_->reset();
  EvaluateCellOptions opts;
  opts.spill_release_callback = pass.release_callback;
  opts.spill_release_user_data = pass.release_user_data;
  const std::uint64_t ordinal = pass.dynamic.next_ordinal();
  pass.dynamic.log().observe(cell, ordinal, opts, &sheet);
  Value result =
      evaluate_cell_for_recalc(pass.workbook, sheet, staged, cell.row, cell.col, pass.registry, *arena_, opts);
  pass.dynamic.log().end();
  if (arena_->exhausted()) {
    return make_error(FormulonErrorCode::kOutOfMemory, pass.oom_message);
  }
  sheet.set_cell_cached_value(cell.row, cell.col, result);
  pass.dynamic.log().note_commit(cell, ordinal);
  ++pass.stats.cells_evaluated;
  return Expected<void, Error>::Ok();
}

Expected<void, Error> RecalcEngine::evaluate_cyclic_component_locked(const SerialEvalPass& pass,
                                                                     const std::vector<CellNodeId>& component,
                                                                     const char* iterative_oom_message) {
  const std::size_t sheet_count = pass.workbook.sheet_count();
  // Evaluation glue for the read-ordered settle (iterative calc off)
  // and the solver (on). `evaluate_with` mirrors the singleton path's
  // evaluator glue: parse on the fly, dispatch through `evaluate()`,
  // fold dynamic-array spills back into a scalar anchor. The `commit`
  // lambda writes the new value into the cell store so the next
  // evaluation reads the freshest value back.
  // Hoisted out of `evaluate_one` so a staged formula's bytes stay live
  // exactly as long as the arena contents of the same call: a string
  // literal inside the formula surfaces in the returned Value as a view
  // into this buffer, and the solver consumes each result before asking
  // for the next one (which resets the arena).
  Cell staged;
  const std::uint64_t component_ordinal = pass.dynamic.next_ordinal();
  auto evaluate_with = [&](CellNodeId c, EvalState* observer) -> Value {
    if (c.sheet_id >= sheet_count) {
      return Value::error(ErrorCode::Ref);
    }
    Sheet& sheet = pass.workbook.sheet(c.sheet_id);
    if (!stage_formula_cell(sheet, c.row, c.col, staged)) {
      // No formula text: nothing to evaluate. Treat as Blank so the
      // solver still has a value to compare against — this happens
      // only on logic bugs but we degrade gracefully.
      return Value::blank();
    }
    // Reset the bump arena per-evaluation. The previous iteration's
    // committed Text scalars survive this reset because
    // `Sheet::set_cell_cached_value` deep-copies Text payloads into
    // each cell's own `cached_text_owned` storage; Array results are
    // similarly deep-copied into `SpillRegion::owned_strings` by
    // `Sheet::commit_spill`.
    arena_->reset();
    EvaluateCellOptions opts;
    opts.spill_release_callback = pass.release_callback;
    opts.spill_release_user_data = pass.release_user_data;
    opts.read_observer = observer;
    pass.dynamic.log().observe(c, component_ordinal, opts, nullptr);
    pass.dynamic.log().note_iterative_member(c, component_ordinal);
    Value result = evaluate_cell_for_recalc(pass.workbook, sheet, staged, c.row, c.col, pass.registry, *arena_, opts);
    pass.dynamic.log().end();
    return result;
  };
  auto evaluate_one = [&](CellNodeId c) { return evaluate_with(c, nullptr); };
  auto commit = [&](CellNodeId c, Value v) {
    if (c.sheet_id >= sheet_count) {
      return;
    }
    Sheet& sheet = pass.workbook.sheet(c.sheet_id);
    sheet.set_cell_cached_value(c.row, c.col, v);
    pass.dynamic.log().note_commit(c, component_ordinal);
  };

  const std::vector<CellNodeId> cells = cells_of_component(component);
  if (!iterative_.enabled) {
    // Excel judges circularity by what evaluation reads, so only a cycle
    // that reads back into itself is one. A cycle through an OFFSET /
    // INDIRECT read keeps its own treatment.
    const bool settled = !closes_through_dynamic_edge(component, graph_) &&
                         settle_component_by_reads(pass.workbook, cells, evaluate_with, commit);
    if (arena_->exhausted()) {
      return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during recalc");
    }
    if (settled) {
      pass.stats.cells_evaluated += static_cast<std::uint32_t>(cells.size());
      return Expected<void, Error>::Ok();
    }
    commit_unresolved_cycle_locked(pass, component);
    return Expected<void, Error>::Ok();
  }
  // A reader first evaluated as a singleton in this recalc has already
  // stepped once; the solver starts from the value it showed before.
  pass.dynamic.log().restore_prior_values(pass.workbook, cells);
  prepare_iterative_component_seeds(pass.workbook, cells);
  const IterativeOutcome outcome =
      run_iterative_solve(cells, iterative_, evaluate_one, commit, progress_cb_, progress_user_data_);
  if (arena_->exhausted()) {
    return make_error(FormulonErrorCode::kOutOfMemory, iterative_oom_message);
  }
  if (outcome.converged) {
    // Solver wrote the converged values into the cell store; count
    // each member as evaluated (singleton-style accounting) plus
    // tagged as iterative.
    pass.stats.cells_evaluated += static_cast<std::uint32_t>(cells.size());
    pass.stats.iterative_cells += static_cast<std::uint32_t>(cells.size());
  } else {
    // Iteration-limit exhaustion or callback-driven abort leaves the
    // last-iteration values in place. The user-visible failure mode is
    // still "the cycle did not resolve", so we
    // count the members in `cycle_cells` to mirror the
    // disabled-iterative-calc accounting.
    // A later pass can continue from the retained approximation.
    pass.stats.cycle_cells += static_cast<std::uint32_t>(cells.size());
  }
  return Expected<void, Error>::Ok();
}

bool RecalcEngine::drop_dynamic_only_cycles_locked(const std::vector<std::vector<CellNodeId>>& sccs) {
  if (!graph_.has_source_edges(DepGraph::DependencySource::kDynamicReference)) {
    return false;
  }
  bool dropped = false;
  for (const std::vector<CellNodeId>& component : sccs) {
    if (!is_cyclic_component(component, graph_) || !closes_through_dynamic_edge(component, graph_)) {
      continue;
    }
    for (const CellNodeId member : component) {
      graph_.clear_dynamic_dependencies_of(member);
    }
    dropped = true;
  }
  return dropped;
}

// ---------------------------------------------------------------------------
// Public mutating API: each entry acquires `mutex_` and delegates to the
// `_locked` body. Internal callers (notably the parallel scheduler) take
// `mutex_` themselves and call the `_locked` helpers directly to avoid
// re-locking.
// ---------------------------------------------------------------------------

void RecalcEngine::register_formula(CellNodeId cell, const parser::AstNode& ast, const Workbook& workbook) {
  std::lock_guard<std::mutex> guard(mutex_);
  register_formula_locked(cell, ast, workbook);
}

void RecalcEngine::register_formula_locked(CellNodeId cell, const parser::AstNode& ast, const Workbook& workbook) {
  // Keep this separate from graph node presence: a syntactically valid
  // formula that currently references a missing sheet can have no static
  // dependency edges at all, yet must be reconsidered when that sheet is
  // appended or renamed into existence.
  has_ever_registered_formula_ = true;

  // A rewrite can change a producer's result shape. Remove its previous
  // candidate-index membership before replacing graph/volatile metadata.
  potential_spill_producers_.update(cell, SpillPotential::kNever);

  // Drop the cell's previous outgoing edges so re-registration is a clean
  // rewrite (the new dependency set may differ from the old one).
  forget_referenced_cells_locked(cell);
  graph_.clear_dependencies_of(cell);
  remove_range_nodes(graph_, range_dependencies_.erase_owner(cell));
  three_d_span_dependencies_.erase(
      std::remove_if(three_d_span_dependencies_.begin(), three_d_span_dependencies_.end(),
                     [cell](const RegisteredThreeDSpan& entry) { return entry.owner == cell; }),
      three_d_span_dependencies_.end());
  // Same for the volatile flag — only re-register if the new AST is still
  // volatile.
  volatiles_.unregister_cell(cell);

  const ExtractedDeps deps = extract_deps(ast, cell.sheet_id, workbook);
  for (CellNodeId dep : deps.cell_deps) {
    graph_.add_dependency(cell, dep);
    referenced_cells_.add(dep);
  }
  for (const ThreeDSheetSpanDependency span : deps.three_d_spans) {
    three_d_span_dependencies_.push_back(RegisteredThreeDSpan{cell, span});
  }
  for (const CellRangeDependency& range : deps.range_deps) {
    // The watcher reads the rectangle's virtual node, which every watcher of
    // the same rectangle shares.
    const RangeDepIndex::AddResult added = range_dependencies_.add(cell, range);
    const CellNodeId node = range_node(added.range_id);
    graph_.add_dependency(cell, node);
    if (!added.is_new) {
      continue;
    }

    // Preserve evaluation order for formulas already inside the range. The
    // compact range table handles literal writes (including cells created in
    // the future); explicit graph edges are needed only for formula cells so
    // Tarjan evaluates their fresh cached values before this aggregate.
    //
    // The sheet's formula-cell index answers that directly, so the cost is
    // the formula cells inside the rectangle — never its area, and never the
    // literal content a whole-column reference spans. A watcher inside its
    // own rectangle reads itself, which is a cycle as in the flattened path.
    const Sheet& sheet = workbook.sheet(range.sheet_id);
    for (const CellAddress source :
         sheet.formula_cells_in(range.row_first, range.col_first, range.row_last, range.col_last)) {
      graph_.add_dependency(node, CellNodeId{range.sheet_id, source.row, source.col});
    }
  }

  // A formula added after an aggregate must also become an explicit graph
  // dependency of every existing rectangle that contains it. Otherwise the
  // watcher would be dirtied, but Tarjan would have no ordering edge to
  // ensure the new formula's cached value is refreshed first.
  range_dependencies_.for_each_range_covering(
      cell, [this, cell](std::uint32_t range_id) { graph_.add_dependency(range_node(range_id), cell); });
  if (deps.is_volatile) {
    volatiles_.register_cell(cell, deps.has_dynamic_reference ? VolatileKind::kDynamicReference : VolatileKind::kValue);
  }
  potential_spill_producers_.update(cell, spill_potential(ast));
}

void RecalcEngine::unregister_formula(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  unregister_formula_locked(cell);
}

void RecalcEngine::unregister_formula_locked(CellNodeId cell) {
  // Preserve the pre-unregistration reverse snapshot. Removing the node
  // would otherwise erase the only path that can wake compact range
  // watchers (including a watcher that was linked through a previous spill
  // footprint), leaving them with a stale cached aggregate.
  for (CellNodeId dependent : graph_.dependents_of_ref(cell)) {
    dirty_.mark(dependent);
  }
  mark_range_dependents_dirty_locked(cell);
  forget_referenced_cells_locked(cell);
  graph_.remove_node(cell);
  drop_cell_registrations_locked(cell);
}

void RecalcEngine::clear_cell_dependencies(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  clear_cell_dependencies_locked(cell);
}

void RecalcEngine::clear_cell_dependencies_locked(CellNodeId cell) {
  forget_referenced_cells_locked(cell);
  graph_.clear_dependencies_of(cell);
  drop_cell_registrations_locked(cell);
}

void RecalcEngine::drop_cell_registrations_locked(CellNodeId cell) {
  remove_range_nodes(graph_, range_dependencies_.erase_owner(cell));
  three_d_span_dependencies_.erase(
      std::remove_if(three_d_span_dependencies_.begin(), three_d_span_dependencies_.end(),
                     [cell](const RegisteredThreeDSpan& entry) { return entry.owner == cell; }),
      three_d_span_dependencies_.end());
  volatiles_.unregister_cell(cell);
  potential_spill_producers_.update(cell, SpillPotential::kNever);
}

void RecalcEngine::mark_dirty(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  mark_dirty_locked(cell);
}

void RecalcEngine::mark_dirty_locked(CellNodeId cell) {
  dirty_.mark(cell);
}

void RecalcEngine::mark_range_dependents_dirty_locked(CellNodeId cell) {
  range_dependencies_.for_each_new_owner_covering(cell, dirty_.generation(),
                                                  [this](CellNodeId owner) { dirty_.mark(owner); });
}

void RecalcEngine::reset_graph_locked() {
  graph_ = DepGraph{};
  range_dependencies_.clear();
  referenced_cells_.clear();
  three_d_span_dependencies_.clear();
  potential_spill_producers_.clear();
  volatiles_.clear();
  dirty_.clear();
}

Expected<RecalcStats, Error> RecalcEngine::recalc(Workbook& workbook, const FunctionRegistry& registry) {
  // Same-thread re-entry would deadlock on the non-recursive `mutex_`.
  // The shared `g_in_recalc` flag (also tracked by `recalc_parallel_impl`)
  // converts that deadlock into a structured `kGraphRecalcReentrant`
  // error, which is the friendlier surface for a UDF or progress callback
  // that accidentally calls back into the engine.
  if (detail::g_in_recalc) {
    return make_error(FormulonErrorCode::kGraphRecalcReentrant,
                      "RecalcEngine::recalc called recursively on the same thread",
                      "the engine does not support nested recalc; the inner call is rejected");
  }
  detail::RecalcReentryGuard reentry_guard;
  std::lock_guard<std::mutex> guard(mutex_);
  return recalc_locked(workbook, registry);
}

Expected<RecalcStats, Error> RecalcEngine::recalc_locked(Workbook& workbook, const FunctionRegistry& registry) {
  RecalcStats stats;
  SpillReleaseQueue release_queue{&workbook};
  const SpillReleaseCallback release_callback = &queue_spill_release;
  SpillWaveBudget wave_budget;
  DynamicReadPass dynamic(*this, workbook);
  const SerialEvalPass pass{
      workbook, registry, dynamic, release_callback, &release_queue, stats, "evaluation arena exhausted during recalc"};

  // Per-wave locals begin after this label and are destroyed on the backward
  // jump, while the counters, queue, and accumulated stats above persist.
recalc_next_wave:
  // Reconcile before dirtiness propagation and SCC construction. A formula
  // rewrite can clear a committed spill before this wave; removing the old
  // derived edge here prevents a stale phantom-only cycle, while additions
  // wake the watcher before topology is built.
  const DepGraph::DependencyDelta pre_dependency_delta = reconcile_spill_dependencies_locked(workbook);
  for (const DepGraph::Edge& edge : pre_dependency_delta.added) {
    dirty_.mark(edge.first);
  }
  for (const DepGraph::Edge& edge : pre_dependency_delta.removed) {
    dirty_.mark(edge.first);
  }

  // ---- Phase 1: seed the dirty set with every volatile cell. ----
  // Volatile formulas re-execute every pass even if their inputs are
  // unchanged. Counting them here also reports how many of the eventually-
  // evaluated cells were forced by volatility.
  volatiles_.for_each([&](CellNodeId cell) {
    if (!dirty_.contains(cell)) {
      dirty_.mark(cell);
    }
    ++stats.volatile_cells;
  });

  // ---- Phase 2: BFS-propagate dirtiness through reverse edges. ----
  // Snapshot the seed list because `dirty_.mark` mutations during BFS
  // would otherwise shift the iteration target.
  std::vector<CellNodeId> bfs_queue;
  bfs_queue.reserve(dirty_.size());
  dirty_.for_each([&](CellNodeId c) { bfs_queue.push_back(c); });
  std::size_t bfs_head = 0;
  while (bfs_head < bfs_queue.size()) {
    const CellNodeId current = bfs_queue[bfs_head++];
    for (CellNodeId dependent : graph_.dependents_of_ref(current)) {
      if (!dirty_.contains(dependent)) {
        dirty_.mark(dependent);
        bfs_queue.push_back(dependent);
      }
    }
  }

  // ---- Phase 3: Tarjan SCC over the dirty induced subgraph. ----
  // Tarjan emits SCCs in reverse-topological order (leaves first), so the
  // evaluation walk below sees a cell's dependencies before the cell
  // itself. Reverse-edge propagation above includes every member of a
  // reachable cycle, so excluding clean nodes preserves SCC boundaries.
  std::unordered_set<CellNodeId, CellNodeIdHash> dirty_nodes;
  dirty_nodes.reserve(dirty_.size());
  dirty_.for_each([&](CellNodeId c) { dirty_nodes.insert(c); });
  std::vector<std::vector<CellNodeId>> sccs = graph_.tarjan_scc_subset(dirty_nodes);
  dynamic.settle_sccs(sccs, dirty_nodes);

  // ---- Phase 4: evaluate every dirty SCC. ----
  // Also track which dirty cells have been visited via the SCC walk, so a
  // final sweep can pick up dirty cells with no graph edges (e.g. a
  // standalone `=NOW()` that reads nothing). Tarjan only emits nodes that
  // appear in the forward / reverse adjacency maps; isolated formula
  // cells are absent from both.
  std::unordered_set<CellNodeId, CellNodeIdHash> visited_in_sccs;
  for (const std::vector<CellNodeId>& component : sccs) {
    // Skip components whose intersection with the dirty set is empty —
    // their cells are already up to date.
    bool any_dirty = false;
    for (CellNodeId c : component) {
      if (dirty_.contains(c)) {
        any_dirty = true;
        break;
      }
    }
    if (!any_dirty) {
      continue;
    }

    if (is_cyclic_component(component, graph_)) {
      // Track which members the dispatcher visited regardless of the
      // resolution path so the standalone-dirty sweep does not re-touch
      // them.
      for (CellNodeId c : component) {
        visited_in_sccs.insert(c);
      }

      RETURN_IF_ERROR(
          evaluate_cyclic_component_locked(pass, component, "evaluation arena exhausted during iterative recalc"));
      continue;
    }

    // Plain singleton: evaluate the cell.
    const CellNodeId only = component.front();
    visited_in_sccs.insert(only);
    RETURN_IF_ERROR(evaluate_formula_cell_locked(pass, only));
  }

  // ---- Phase 4b: defensive pickup for dirty cells not visited by Tarjan. ----
  // `tarjan_scc_subset` emits isolated selected nodes, so this normally
  // remains empty. Keep the sweep as a guard for future graph mutations.
  std::vector<CellNodeId> standalone_dirty;
  dirty_.for_each([&](CellNodeId c) {
    if (visited_in_sccs.count(c) == 0U) {
      standalone_dirty.push_back(c);
    }
  });
  for (CellNodeId c : standalone_dirty) {
    RETURN_IF_ERROR(evaluate_formula_cell_locked(pass, c));
  }

  // Reconcile compact-range -> spill-producer edges only after every cell in
  // this wave has committed its result. Newly acquired source edges schedule
  // their watcher for one targeted follow-up wave; old edges remain active
  // throughout the current wave, so a producer shrink still dirties and
  // orders its watcher correctly.
  const DepGraph::DependencyDelta dependency_delta = reconcile_spill_dependencies_locked(workbook);
  // A reader that saw a target before its final value (or itself) runs
  // again behind the edge it just taught the graph.
  const bool dynamic_retry = !dynamic.end_wave(nullptr).empty();
  const bool dependency_retry = !dependency_delta.added.empty() || dynamic_retry;
  if (dependency_retry) {
    wave_budget.count_dependency_wave();
    for (const DepGraph::Edge& edge : dependency_delta.added) {
      dirty_.mark(edge.first);
    }
  }

  // A spill shape change can release cells that blocked a different pending
  // producer. Keep those anchors for a dependency-ordered next wave instead
  // of losing them when this pass clears its dirty set.
  const std::vector<CellNodeId> released = release_queue.take();
  if (!released.empty() || dependency_retry) {
    wave_budget.count_release_wave();
    wave_budget.observe_release(workbook, released);
    if (wave_budget.exceeded(!released.empty())) {
      // Do not recurse through an unbounded chain of release waves. Preserve
      // the existing dirty set (including unrelated work) and keep the
      // release targets dirty for a caller retry after an external mutation.
      const LockedMutator mutator = locked_mutator();
      mark_spill_release_wave(mutator, released, graph_);
      return make_error(FormulonErrorCode::kGraphScheduleFailed, "recalc waves made no progress",
                        "spill recovery or dynamic-reference retries exceeded the bounded wave budget");
    }
    dirty_.clear();
    const LockedMutator mutator = locked_mutator();
    mark_spill_release_wave(mutator, released, graph_);
    for (const DepGraph::Edge& edge : dependency_delta.added) {
      mutator.mark_dirty(edge.first);
    }
    dynamic.mark_stale_dirty();
    goto recalc_next_wave;
  }

  // ---- Phase 5: clear the dirty set. ----
  dirty_.clear();
  disabled_cycle_refs_pending_ = false;
  return stats;
}

Expected<RecalcStats, Error> RecalcEngine::partial_recalc(Workbook& workbook, const FunctionRegistry& registry,
                                                          const SheetCellRange& viewport) {
  // Mirror `recalc()` re-entry handling: a callback that calls back into a
  // recalc API would otherwise deadlock on the engine mutex.
  if (detail::g_in_recalc) {
    return make_error(FormulonErrorCode::kGraphRecalcReentrant,
                      "RecalcEngine::partial_recalc called recursively on the same thread",
                      "the engine does not support nested recalc; the inner call is rejected");
  }
  detail::RecalcReentryGuard reentry_guard;
  std::lock_guard<std::mutex> guard(mutex_);
  return partial_recalc_locked(workbook, registry, viewport);
}

Expected<RecalcStats, Error> RecalcEngine::partial_recalc_locked(Workbook& workbook, const FunctionRegistry& registry,
                                                                 const SheetCellRange& viewport) {
  RecalcStats stats;
  SpillReleaseQueue release_queue{&workbook};
  const SpillReleaseCallback release_callback = &queue_spill_release;
  SpillWaveBudget wave_budget;
  DynamicReadPass dynamic(*this, workbook);
  const SerialEvalPass pass{workbook,
                            registry,
                            dynamic,
                            release_callback,
                            &release_queue,
                            stats,
                            "evaluation arena exhausted during partial recalc"};

  // ---- Phase 0: validate the viewport. ----
  // Empty viewport — collapsed row / column range, or unknown sheet —
  // is a no-op. The dirty set stays untouched so a subsequent full
  // recalc still picks up everything.
  const std::size_t sheet_count = workbook.sheet_count();
  if (viewport.sheet_id >= sheet_count) {
    return stats;
  }
  if (viewport.first_row > viewport.last_row || viewport.first_col > viewport.last_col) {
    return stats;
  }

  // Oversized viewport: the seed loop below visits every coordinate in the
  // rectangle, so a full-grid viewport would enumerate
  // `Sheet::kMaxRows * Sheet::kMaxCols` (~17e9) coordinates before any
  // dependency work starts. A viewport is a UI redraw region and never
  // approaches `kMaxRecalcViewportCells`; treat a larger rectangle like the
  // other invalid-viewport shapes above (no-op, dirty set preserved) so
  // callers fall back to a full `recalc()`.
  const utils::RectRange viewport_rect(viewport.first_row, viewport.first_col, viewport.last_row, viewport.last_col);
  ResourceBudget seed_budget(kMaxRecalcViewportCells);
  if (seed_budget.would_exceed(viewport_rect.size())) {
    return stats;
  }

  // As in the full pass, this label keeps the release-wave loop iterative;
  // locals in each wave's body leave scope before the backward jump.
partial_recalc_next_wave:
  // Keep the graph's spill-derived topology current before forming the
  // viewport closure. Both source additions and removals dirty their range
  // watcher; this is essential when a producer rewrite has just invalidated
  // its committed footprint.
  const DepGraph::DependencyDelta pre_dependency_delta = reconcile_spill_dependencies_locked(workbook);
  for (const DepGraph::Edge& edge : pre_dependency_delta.added) {
    dirty_.mark(edge.first);
  }
  for (const DepGraph::Edge& edge : pre_dependency_delta.removed) {
    dirty_.mark(edge.first);
  }

  // ---- Phase 1: enumerate the viewport's seed cells. ----
  // Walk the requested rectangle and pull every populated cell into the
  // seed set. Cells outside the sheet's stored extent are silently
  // dropped (the dep graph will not have entries for them anyway).
  // Phantoms of dynamic-array spills are intentionally NOT seeded
  // separately — the spill anchor is the formula cell, and resolving the
  // anchor forces the spill to refresh.
  const Sheet& view_sheet = workbook.sheet(viewport.sheet_id);
  std::vector<CellNodeId> seeds;
  seeds.reserve(static_cast<std::size_t>(viewport_rect.size()));
  for (auto [row, col] : viewport_rect) {
    // We seed every coordinate inside the viewport regardless of
    // whether it currently holds a stored cell: a viewport coordinate
    // that is presently blank may still have inbound dep-graph
    // edges (e.g. a formula on it that has been cleared but whose
    // dependents have not been re-registered yet). The closure walk
    // below tolerates absent nodes.
    (void)view_sheet;
    seeds.push_back(CellNodeId{viewport.sheet_id, row, col});
  }

  // ---- Phase 2: compute the dependency closure backward from seeds. ----
  // The closure is "every cell whose value the viewport (transitively)
  // reads". We walk forward edges (`dependencies_of`) from each seed:
  // if A reads B, B's value contributes to A. The closure includes the
  // seed cells themselves so any dirty viewport cell still gets
  // evaluated.
  std::unordered_set<CellNodeId, CellNodeIdHash> closure;
  closure.reserve(seeds.size());
  std::vector<CellNodeId> bfs_queue = seeds;
  for (CellNodeId seed : seeds) {
    closure.insert(seed);
  }
  std::size_t bfs_head = 0;
  while (bfs_head < bfs_queue.size()) {
    const CellNodeId current = bfs_queue[bfs_head++];
    for (CellNodeId predecessor : graph_.dependencies_of_ref(current)) {
      if (closure.insert(predecessor).second) {
        bfs_queue.push_back(predecessor);
      }
    }
  }

  // Pull indexed potential spill producers into the closure: no derived edge exists before the first commit.
  bool expanded_potential_producers = true;
  std::unordered_set<std::uint32_t> processed_range_entries;
  processed_range_entries.reserve(range_dependencies_.distinct_range_count());
  Arena potential_arena;

  // The potential-producer index ignores the registry, so re-classify each candidate against it.
  const auto is_admissible_potential_producer = [&](CellNodeId producer, SpillPotential static_potential) {
    return SpillProducerIndex::admissible(workbook, registry, producer, static_potential, potential_arena);
  };

  const auto enqueue_potential_producer = [&](CellNodeId producer, SpillPotential static_potential) {
    if (!is_admissible_potential_producer(producer, static_potential)) {
      return;
    }
    if (closure.insert(producer).second) {
      expanded_potential_producers = true;
      bfs_queue.push_back(producer);
    }
  };

  // A phantom spill coordinate has no edge to its uncommitted anchor; scan for producers above-left.
  std::size_t potential_coordinate_head = 0;
  while (expanded_potential_producers) {
    expanded_potential_producers = false;
    if (dynamic.seed_pending_candidates(registry, closure, bfs_queue)) {
      expanded_potential_producers = true;
    }
    while (potential_coordinate_head < bfs_queue.size()) {
      const CellNodeId current = bfs_queue[potential_coordinate_head++];
      if (current.sheet_id >= sheet_count || !potential_spill_producers_.has_sheet(current.sheet_id)) {
        continue;
      }
      const Cell* current_cell = workbook.sheet(current.sheet_id).cell_at(current.row, current.col);
      if (blocks_spill(current_cell)) {
        continue;
      }
      const SpillTarget target(workbook.sheet(current.sheet_id), current.row, current.col);
      for (const auto& [producer, static_potential] : potential_spill_producers_.on_sheet(current.sheet_id)) {
        if (producer.sheet_id != current.sheet_id || !target.reachable_from(producer)) {
          continue;
        }
        enqueue_potential_producer(producer, static_potential);
      }
    }
    range_dependencies_.for_each_distinct_range(
        [&](std::uint32_t range_id, const CellRangeDependency& range, const std::vector<CellNodeId>& owners) {
          if (range.sheet_id >= sheet_count || !potential_spill_producers_.has_sheet(range.sheet_id)) {
            return;
          }
          if (processed_range_entries.count(range_id) != 0U) {
            return;
          }
          const bool watched_from_closure = std::any_of(
              owners.begin(), owners.end(), [&closure](CellNodeId owner) { return closure.count(owner) != 0U; });
          if (!watched_from_closure) {
            return;
          }
          processed_range_entries.insert(range_id);
          for (const auto& [producer, static_potential] : potential_spill_producers_.on_sheet(range.sheet_id)) {
            // A spill extends only down and right, so an anchor past either last bound cannot reach the range.
            if (producer.row > range.row_last || producer.col > range.col_last) {
              continue;
            }
            enqueue_potential_producer(producer, static_potential);
          }
        });
    while (bfs_head < bfs_queue.size()) {
      const CellNodeId current = bfs_queue[bfs_head++];
      for (CellNodeId predecessor : graph_.dependencies_of_ref(current)) {
        if (closure.insert(predecessor).second) {
          bfs_queue.push_back(predecessor);
          expanded_potential_producers = true;
        }
      }
    }
  }

  // ---- Phase 3: BFS-propagate dirtiness inside the closure only. ----
  // Volatile cells inside the closure are forced dirty (the caller
  // asked for fresh values within the viewport, and a volatile cell's
  // result is by definition stale). Volatile cells OUTSIDE the closure
  // are NOT touched: the user opted into a viewport-bounded recalc and
  // expects volatile cells in remote regions to wait for the next
  // full pass.
  volatiles_.for_each([&](CellNodeId cell) {
    if (closure.count(cell) == 0U) {
      return;
    }
    if (!dirty_.contains(cell)) {
      dirty_.mark(cell);
    }
    ++stats.volatile_cells;
  });

  // Now propagate dirtiness through the closure. We snapshot the dirty
  // cells that are inside the closure and BFS through reverse edges,
  // adding any newly-discovered dependents that themselves live inside
  // the closure (a dependent outside the closure is by definition
  // unreachable from the viewport, so we do not need to visit it).
  const auto for_each_dirty_in_closure = [&](auto&& visit) {
    dirty_.for_each([&](CellNodeId c) {
      if (closure.count(c) != 0U) {
        visit(c);
      }
    });
  };
  std::vector<CellNodeId> propagation_queue;
  propagation_queue.reserve(dirty_.size());
  for_each_dirty_in_closure([&](CellNodeId c) { propagation_queue.push_back(c); });
  std::size_t prop_head = 0;
  while (prop_head < propagation_queue.size()) {
    const CellNodeId current = propagation_queue[prop_head++];
    for (CellNodeId dependent : graph_.dependents_of_ref(current)) {
      if (closure.count(dependent) == 0U) {
        continue;
      }
      if (!dirty_.contains(dependent)) {
        dirty_.mark(dependent);
        propagation_queue.push_back(dependent);
      }
    }
  }

  // ---- Phase 4: Tarjan SCC + selective evaluation. ----
  // Restrict Tarjan to the dirty cells in the viewport closure. Cells
  // outside it are never visited and retain their dirty flag for a later
  // full / overlapping partial recalc.
  std::unordered_set<CellNodeId, CellNodeIdHash> dirty_closure;
  dirty_closure.reserve(propagation_queue.size());
  for_each_dirty_in_closure([&](CellNodeId c) { dirty_closure.insert(c); });
  std::vector<std::vector<CellNodeId>> sccs = graph_.tarjan_scc_subset(dirty_closure);
  dynamic.settle_sccs(sccs, dirty_closure);
  std::unordered_set<CellNodeId, CellNodeIdHash> visited_in_sccs;
  for (const std::vector<CellNodeId>& component : sccs) {
    // Skip components that have no overlap with the closure: their
    // cells are not transitively read by the viewport.
    bool any_in_closure = false;
    for (CellNodeId c : component) {
      if (closure.count(c) != 0U) {
        any_in_closure = true;
        break;
      }
    }
    if (!any_in_closure) {
      continue;
    }

    // Skip components that have no dirty member: nothing to do here.
    bool any_dirty = false;
    for (CellNodeId c : component) {
      if (dirty_.contains(c) && closure.count(c) != 0U) {
        any_dirty = true;
        break;
      }
    }
    if (!any_dirty) {
      continue;
    }

    if (is_cyclic_component(component, graph_)) {
      // A cycle that the viewport reaches must still be surfaced as a
      // cycle: the closure restriction never hides a circular reference.
      // Mirrors `recalc()`'s cycle handling exactly, with the iterative
      // solver wired to the same progress callback.
      for (CellNodeId c : component) {
        visited_in_sccs.insert(c);
      }

      RETURN_IF_ERROR(
          evaluate_cyclic_component_locked(pass, component, "evaluation arena exhausted during partial recalc"));
      continue;
    }

    // Plain singleton: only evaluate if the cell is in the closure AND
    // dirty. Cells outside the closure stay dirty for a later pass.
    const CellNodeId only = component.front();
    if (closure.count(only) == 0U || !dirty_.contains(only)) {
      continue;
    }
    visited_in_sccs.insert(only);
    RETURN_IF_ERROR(evaluate_formula_cell_locked(pass, only));
  }

  // ---- Phase 4b: standalone dirty cells inside the closure. ----
  // Isolated formula cells with no dep-graph entries do not appear in
  // Tarjan output; sweep the closure for any dirty cell we have not
  // already touched.
  std::vector<CellNodeId> standalone_dirty;
  dirty_.for_each([&](CellNodeId c) {
    if (closure.count(c) != 0U && visited_in_sccs.count(c) == 0U) {
      standalone_dirty.push_back(c);
    }
  });
  for (CellNodeId c : standalone_dirty) {
    RETURN_IF_ERROR(evaluate_formula_cell_locked(pass, c));
  }

  // Reconcile globally after this partial wave. A newly discovered
  // watcher-to-producer edge is retried here only when that watcher belongs
  // to this viewport closure; a remote watcher remains dirty for a later
  // full or overlapping partial pass.
  const DepGraph::DependencyDelta dependency_delta = reconcile_spill_dependencies_locked(workbook);
  bool dependency_retry_in_closure = false;
  for (const DepGraph::Edge& edge : dependency_delta.added) {
    dirty_.mark(edge.first);
    if (closure.count(edge.first) != 0U) {
      dependency_retry_in_closure = true;
    }
  }
  // Judged before the closure is unmarked, so a target this viewport left
  // dirty still reads as unevaluated.
  const std::vector<CellNodeId>& stale_readers = dynamic.end_wave(&closure, &registry);
  for (const CellNodeId reader : stale_readers) {
    if (closure.count(reader) != 0U) {
      dependency_retry_in_closure = true;
    }
  }

  // ---- Phase 5: clear only the closure's dirty entries. ----
  // Cells outside the closure must remain dirty so a subsequent
  // `recalc()` (or an overlapping `partial_recalc`) revisits them.
  // Snapshot the in-closure dirty entries first to avoid mutating the
  // underlying container during iteration.
  std::vector<CellNodeId> to_unmark;
  to_unmark.reserve(closure.size());
  for_each_dirty_in_closure([&](CellNodeId c) { to_unmark.push_back(c); });
  // Hand dirtiness on to dependents outside the closure so a later pass still reaches them.
  const auto mark_outside_closure = [&](CellNodeId dependent) {
    if (closure.count(dependent) == 0U) {
      dirty_.mark(dependent);
    }
  };
  for (CellNodeId c : to_unmark) {
    for (CellNodeId dependent : graph_.dependents_of_ref(c)) {
      mark_outside_closure(dependent);
    }
    range_dependencies_.for_each_owner_covering(c, mark_outside_closure);
  }
  for (CellNodeId c : to_unmark) {
    dirty_.unmark(c);
  }
  for (const DepGraph::Edge& edge : dependency_delta.added) {
    if (closure.count(edge.first) != 0U) {
      dirty_.mark(edge.first);
    }
  }
  dynamic.mark_stale_dirty();

  // A committed spill can release a pending producer while this viewport is
  // being evaluated. Preserve releases outside the closure for a later full
  // or overlapping partial pass; releases inside the closure get a true
  // dependency-ordered next wave before this partial call returns.
  const std::vector<CellNodeId> released = release_queue.take();
  if (!released.empty() || dependency_retry_in_closure) {
    const LockedMutator mutator = locked_mutator();
    mark_spill_release_wave(mutator, released, graph_);
    bool release_in_closure = false;
    for (const CellNodeId anchor : released) {
      if (closure.count(anchor) != 0U) {
        release_in_closure = true;
        break;
      }
    }
    if (release_in_closure || dependency_retry_in_closure) {
      if (!released.empty()) {
        wave_budget.count_release_wave();
        wave_budget.observe_release(workbook, released);
      }
      if (dependency_retry_in_closure) {
        wave_budget.count_dependency_wave();
      }
      if (wave_budget.exceeded(!released.empty())) {
        return make_error(FormulonErrorCode::kGraphScheduleFailed, "recalc waves made no progress",
                          "partial spill recovery or dynamic-reference retries exceeded the bounded wave budget");
      }
      goto partial_recalc_next_wave;
    }
  }
  return stats;
}

}  // namespace formulon::eval
