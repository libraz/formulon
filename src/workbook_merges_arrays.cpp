//
// Merged ranges and legacy array formulas for `Workbook`: the merge mutators
// with the spill-anchor wake-up they owe, and array-formula entry and lookup.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "data_table.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "io/xlsb/table_records.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"
#include "workbook_sheet_mutation.h"

namespace formulon {

namespace {

using detail::check_cell_target;

// Refusal of `target` as the extent of a new legacy array formula, on the
// first structure it collides with. A legacy array block wholly inside
// `target` is replaced, so only one it partially overlaps counts.
Expected<void, Error> check_array_target_free(const Workbook& wb, std::size_t sheet_index, const MergeRange& target) {
  const Sheet& sheet = wb.sheet(sheet_index);
  const auto refuse = [&](const char* what) {
    return make_error(FormulonErrorCode::kPreconditionFailed, what,
                      "row=" + std::to_string(target.first_row) + " col=" + std::to_string(target.first_col));
  };
  for (const CellAddress address : sheet.all_formula_cells()) {
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell == nullptr || cell->array_block_rows == 0U) {
      continue;
    }
    const MergeRange block{address.row, address.col, address.row + cell->array_block_rows - 1U,
                           address.col + cell->array_block_cols - 1U};
    if (intersects(block, target) && !contains(target, block)) {
      return refuse("range overlaps part of an array formula");
    }
  }
  for (const MergeRange& merge : sheet.merges()) {
    if (intersects(normalized(merge), target)) {
      return refuse("array formulas are not valid in merged cells");
    }
  }
  for (const DataTable& table : sheet.data_tables()) {
    if (intersects(normalized(table.ref), target)) {
      return refuse("range overlaps part of a data table");
    }
  }
  for (const std::unique_ptr<pivot::PivotTable>& pivot : sheet.pivot_tables()) {
    if (pivot == nullptr || pivot->span_rows() == 0U || pivot->span_cols() == 0U) {
      continue;
    }
    const MergeRange area{pivot->anchor_row(), pivot->anchor_col(), pivot->anchor_row() + pivot->span_rows() - 1U,
                          pivot->anchor_col() + pivot->span_cols() - 1U};
    if (intersects(area, target)) {
      return refuse("range overlaps part of a PivotTable");
    }
  }
  if (target.first_row != target.last_row || target.first_col != target.last_col) {
    for (const TableMetadata& table : wb.tables()) {
      const std::optional<MergeRange> ref = detail::table_rect_on_sheet(table, sheet_index);
      if (ref.has_value() && intersects(*ref, target)) {
        return refuse("range must not span several cells inside a table");
      }
    }
  }
  return Expected<void, Error>::Ok();
}

// Refusal of `merge` as a new merged area, on the first structure it
// intersects: another merge is `kAlreadyExists`, a structure Excel keeps
// merges out of is `kPreconditionFailed`.
Expected<void, Error> check_merge_target_free(const Workbook& wb, std::size_t sheet_index, const MergeRange& merge) {
  const Sheet& sheet = wb.sheet(sheet_index);
  const auto context = [&] {
    std::string ref;
    append_a1_rectangle(ref, merge);
    return "ref=" + ref;
  };
  const auto refuse = [&](const char* what) {
    return make_error(FormulonErrorCode::kPreconditionFailed, what, context());
  };
  for (const MergeRange& existing : sheet.merges()) {
    if (intersects(normalized(existing), merge)) {
      return make_error(FormulonErrorCode::kAlreadyExists, "range overlaps an existing merge", context());
    }
  }
  for (const TableMetadata& table : wb.tables()) {
    const std::optional<MergeRange> ref = detail::table_rect_on_sheet(table, sheet_index);
    if (ref.has_value() && intersects(*ref, merge)) {
      return refuse("merge overlaps a table");
    }
  }
  for (const io::xlsb::XlsbRetainedTable& kept : io::xlsb::xlsb_retained_tables(wb)) {
    const std::optional<MergeRange> kept_rect = parse_a1_rectangle(kept.ref);
    if (kept.sheet_index == sheet_index && kept_rect && intersects(*kept_rect, merge)) {
      return refuse("merge overlaps a table");
    }
  }
  for (const std::unique_ptr<pivot::PivotTable>& pivot : sheet.pivot_tables()) {
    if (pivot != nullptr && intersects(detail::pivot_occupied_extent(*pivot), merge)) {
      return refuse("merge overlaps part of a PivotTable");
    }
  }
  for (const CellAddress address : sheet.formula_cells_in(0U, 0U, merge.last_row, merge.last_col)) {
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell == nullptr || cell->array_block_rows == 0U) {
      continue;
    }
    const MergeRange block{address.row, address.col, address.row + cell->array_block_rows - 1U,
                           address.col + cell->array_block_cols - 1U};
    if (intersects(block, merge)) {
      return refuse("array formulas are not valid in merged cells");
    }
  }
  for (const DataTable& table : sheet.data_tables()) {
    if (intersects(normalized(table.ref), merge)) {
      return refuse("merge overlaps part of a data table");
    }
  }
  return Expected<void, Error>::Ok();
}

/// `Sheet::blocked_spill_anchors_intersecting` or
/// `Sheet::committed_spill_anchors_intersecting`.
using SpillAnchorQuery = std::vector<CellAddress> (Sheet::*)(std::uint32_t, std::uint32_t, std::uint32_t,
                                                             std::uint32_t) const;

void mark_spill_anchors_intersecting_merge(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                           const std::vector<Sheet>& sheets, const MergeRange& merge,
                                           SpillAnchorQuery query) {
  if (merge.first_row > merge.last_row || merge.first_col > merge.last_col ||
      !Sheet::coord_in_grid(merge.first_row, merge.first_col) ||
      !Sheet::coord_in_grid(merge.last_row, merge.last_col)) {
    return;
  }
  const std::uint32_t rows = merge.last_row - merge.first_row + 1U;
  const std::uint32_t cols = merge.last_col - merge.first_col + 1U;
  for (const CellAddress anchor : (sheets[sheet_index].*query)(merge.first_row, merge.first_col, rows, cols)) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
}

}  // namespace

std::optional<Workbook::ArrayFormulaInfo> Workbook::array_formula_at(std::size_t sheet_index, std::uint32_t row,
                                                                     std::uint32_t col) const {
  if (sheet_index >= sheets_.size() || !Sheet::coord_in_grid(row, col)) {
    return std::nullopt;
  }
  const Sheet& sheet = sheets_[sheet_index];
  std::optional<SpillFootprint> block = sheet.array_block_covering(row, col);
  if (!block.has_value()) {
    // Before a recalc distributes the block, only its anchor is known to the
    // sheet: look for the anchor above and to the left.
    for (const CellAddress address : sheet.formula_cells_in(0U, 0U, row, col)) {
      const Cell* anchor = sheet.cell_at(address.row, address.col);
      if (anchor != nullptr && anchor->array_block_rows != 0U && row - address.row < anchor->array_block_rows &&
          col - address.col < anchor->array_block_cols) {
        block = SpillFootprint{address.row, address.col, anchor->array_block_rows, anchor->array_block_cols};
        break;
      }
    }
  }
  if (!block.has_value()) {
    return std::nullopt;
  }
  Sheet::CellRead read;
  sheet.read_formula_cell(block->anchor_row, block->anchor_col, read);
  return ArrayFormulaInfo{block->anchor_row, block->anchor_col, block->rows, block->cols,
                          std::string(read.formula_text())};
}

Expected<void, Error> Workbook::set_array_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                  std::uint32_t rows, std::uint32_t cols, std::string formula,
                                                  FormulaSource source) {
  RETURN_IF_ERROR(check_cell_target(sheet_index, sheets_.size(), row, col));
  if (rows == 0U || cols == 0U) {
    return make_error(FormulonErrorCode::kInvalidArgument, "rows and cols must be >= 1",
                      "rows=" + std::to_string(rows) + " cols=" + std::to_string(cols));
  }
  if (rows > Sheet::kMaxRows - row || cols > Sheet::kMaxCols - col ||
      static_cast<std::uint64_t>(rows) * cols > kMaxRangeExpansionCells) {
    return make_error(FormulonErrorCode::kInvalidArgument, "array outside grid or too large",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col) + " rows=" + std::to_string(rows) +
                          " cols=" + std::to_string(cols));
  }
  const MergeRange target{row, col, row + rows - 1U, col + cols - 1U};
  RETURN_IF_ERROR(check_array_target_free(*this, sheet_index, target));

  // The block's own cells must hold nothing: a stored value or formula would
  // make the distributed result a #SPILL!. Only populated coordinates are
  // visited, so the cost follows what is there, not the rectangle's area.
  struct Cleared {
    std::vector<CellAddress> cells;
  } cleared;
  std::uint64_t cursor = 0;
  while (cursor != Sheet::kCellCursorEnd) {
    cursor = sheets_[sheet_index].cells_in_range(
        target.first_row, target.first_col, target.last_row, target.last_col, cursor, 4096U,
        [](const Sheet::RangeCell& cell, void* ctx) {
          static_cast<Cleared*>(ctx)->cells.push_back(CellAddress{cell.row, cell.col});
        },
        &cleared);
  }
  for (const CellAddress address : cleared.cells) {
    if (address.row == row && address.col == col) {
      continue;
    }
    // Clearing one cell can take a whole block or spill with it.
    if (sheets_[sheet_index].cell_at(address.row, address.col) == nullptr &&
        !sheets_[sheet_index].committed_spill_footprint_covering(address.row, address.col).has_value()) {
      continue;
    }
    put_cell_value(sheet_index, address.row, address.col, Value::blank());
  }
  put_cell_formula(sheet_index, row, col, std::move(formula), source, rows, cols);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::add_merge(std::size_t sheet_index, MergeRange merge) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  merge = normalized(merge);
  if (!Sheet::rect_in_grid(merge.first_row, merge.first_col, merge.last_row, merge.last_col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "range out of grid",
                      "first_row=" + std::to_string(merge.first_row) + " first_col=" + std::to_string(merge.first_col) +
                          " last_row=" + std::to_string(merge.last_row) +
                          " last_col=" + std::to_string(merge.last_col));
  }
  RETURN_IF_ERROR(check_merge_target_free(*this, sheet_index, merge));
  put_merge(sheet_index, merge);
  return Expected<void, Error>::Ok();
}

void Workbook::put_merge(std::size_t sheet_index, const MergeRange& merge) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  sheets_[sheet_index].add_merge(merge);
  // A merge added over a committed spill becomes a blocker for the next
  // evaluation. Keep the anchor dirty so recalc clears the old rectangle and
  // records the resulting #SPILL! state under the normal commit contract.
  mark_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, merge,
                                        &Sheet::committed_spill_anchors_intersecting);
}

Expected<void, Error> Workbook::remove_merges_intersecting(std::size_t sheet_index, MergeRange merge) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  merge = normalized(merge);
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const std::vector<MergeRange> removed = sheets_[sheet_index].remove_merges_intersecting(merge);
  for (const MergeRange& erased : removed) {
    mark_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, erased,
                                          &Sheet::blocked_spill_anchors_intersecting);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_merge_at(std::size_t sheet_index, std::size_t index) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  MergeRange removed;
  if (!sheets_[sheet_index].remove_merge_at(index, &removed)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "index out of range", "index=" + std::to_string(index));
  }
  mark_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, removed,
                                        &Sheet::blocked_spill_anchors_intersecting);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::clear_merges(std::size_t sheet_index) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  for (const CellAddress anchor : sheets_[sheet_index].blocked_spill_anchors()) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
  sheets_[sheet_index].clear_merges();
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
