// @size-budget: 6 KB
//
// The occupancy rule a pivot report's range answers to, the recovery of
// blocked reports after a recalc pass, and the cut-and-paste move of a
// report.
//
// A report range is occupied by a value or formula cell, a spill range cell,
// a merge, a legacy array block, a data-table result cell, a table or another
// PivotTable's current range, even one only straddling it. A blocked report
// keeps its needed range relative to its anchor and renders on the first
// recalc pass that finds the range free, earlier-created reports first. A
// move clears what a paste would overwrite, arrays and data tables inside the
// destination included, but refuses another report, an array or data table
// crossing the destination's edge, and the sheet edge; it never blocks.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "data_table.h"
#include "drawing/slicer_edit.h"
#include "merge_range.h"
#include "pivot/field_remap.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/status_macros.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {
namespace {

using Occupant = detail::ReportOccupant;

// The top-left cell `a` and `b` share; they must intersect.
Occupant occupant_at(Occupant::Kind kind, const MergeRange& a, const MergeRange& b) {
  return Occupant{kind, a.first_row > b.first_row ? a.first_row : b.first_row,
                  a.first_col > b.first_col ? a.first_col : b.first_col};
}

std::optional<Occupant> pivot_in(const Sheet& sheet, std::size_t pivot_index, const MergeRange& target) {
  for (std::size_t i = 0; i < sheet.pivot_tables().size(); ++i) {
    const std::unique_ptr<pivot::PivotTable>& other = sheet.pivot_tables()[i];
    if (i == pivot_index || other == nullptr) {
      continue;
    }
    const MergeRange extent = detail::pivot_occupied_extent(*other);
    if (intersects(extent, target)) {
      return occupant_at(Occupant::Kind::kPivot, extent, target);
    }
  }
  return std::nullopt;
}

std::optional<Occupant> data_table_in(const Sheet& sheet, const MergeRange& target) {
  for (const DataTable& table : sheet.data_tables()) {
    const MergeRange ref = normalized(table.ref);
    if (intersects(ref, target)) {
      return occupant_at(Occupant::Kind::kDataTable, ref, target);
    }
  }
  return std::nullopt;
}

std::optional<Occupant> array_in(const Sheet& sheet, const MergeRange& target) {
  for (const CellAddress address : sheet.formula_cells_in(0U, 0U, target.last_row, target.last_col)) {
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell == nullptr || cell->array_block_rows == 0U) {
      continue;
    }
    const MergeRange block{address.row, address.col, address.row + cell->array_block_rows - 1U,
                           address.col + cell->array_block_cols - 1U};
    if (intersects(block, target)) {
      return occupant_at(Occupant::Kind::kArray, block, target);
    }
  }
  return std::nullopt;
}

struct CellScan {
  MergeRange own;
  std::optional<CellAddress> found;
};

void note_cell(const Sheet::RangeCell& cell, void* ctx) {
  CellScan& scan = *static_cast<CellScan*>(ctx);
  const bool owned = cell.row >= scan.own.first_row && cell.row <= scan.own.last_row &&
                     cell.col >= scan.own.first_col && cell.col <= scan.own.last_col;
  if (!scan.found.has_value() && !owned) {
    scan.found = CellAddress{cell.row, cell.col};
  }
}

// The first value, formula or spill cell of `target` outside `own`, row-major.
std::optional<Occupant> cell_in(const Sheet& sheet, const MergeRange& target, const MergeRange& own) {
  constexpr std::uint32_t kBatch = 64;
  CellScan scan{own, std::nullopt};
  std::uint64_t cursor = 0;
  while (!scan.found.has_value() && cursor != Sheet::kCellCursorEnd) {
    cursor = sheet.cells_in_range(target.first_row, target.first_col, target.last_row, target.last_col, cursor, kBatch,
                                  &note_cell, &scan);
  }
  if (!scan.found.has_value()) {
    return std::nullopt;
  }
  const Cell* cell = sheet.cell_at(scan.found->row, scan.found->col);
  const bool stored = cell != nullptr && (!cell->formula_text.empty() || !cell->cached_value.is_blank());
  return Occupant{stored ? Occupant::Kind::kCell : Occupant::Kind::kSpill, scan.found->row, scan.found->col};
}

// The sheet's pivot indices in creation order, earliest first.
std::vector<std::uint32_t> creation_order_of(const Sheet& sheet) {
  const std::vector<std::unique_ptr<pivot::PivotTable>>& tables = sheet.pivot_tables();
  const auto earlier = [&tables](std::uint32_t a, std::uint32_t b) {
    return tables[a] != nullptr && tables[b] != nullptr && detail::created_before(*tables[a], a, *tables[b], b);
  };
  std::vector<std::uint32_t> order;
  sorted_index_order(order, static_cast<std::uint32_t>(tables.size()), make_index_less(earlier));
  return order;
}

std::size_t count_blocked(const std::vector<std::unique_ptr<pivot::PivotTable>>& tables) {
  std::size_t count = 0;
  for (const std::unique_ptr<pivot::PivotTable>& table : tables) {
    count += table != nullptr && table->blocked().has_value() ? 1U : 0U;
  }
  return count;
}

}  // namespace

std::optional<detail::ReportOccupant> detail::find_report_occupant(const Workbook& wb, std::size_t sheet_index,
                                                                   std::size_t pivot_index, const MergeRange& target,
                                                                   const MergeRange& own) {
  const Sheet& sheet = wb.sheet(sheet_index);
  if (std::optional<Occupant> hit = pivot_in(sheet, pivot_index, target)) {
    return hit;
  }
  for (const TableMetadata& table : wb.tables()) {
    const std::optional<MergeRange> ref = table_rect_on_sheet(table, sheet_index);
    if (ref.has_value() && intersects(normalized(*ref), target)) {
      return occupant_at(Occupant::Kind::kTable, normalized(*ref), target);
    }
  }
  if (std::optional<Occupant> hit = data_table_in(sheet, target)) {
    return hit;
  }
  if (std::optional<Occupant> hit = array_in(sheet, target)) {
    return hit;
  }
  for (const MergeRange& merge : sheet.merges()) {
    const MergeRange m = normalized(merge);
    if (intersects(m, target) && !contains(own, m)) {
      return occupant_at(Occupant::Kind::kMerge, m, target);
    }
  }
  return cell_in(sheet, target, own);
}

void Workbook::adopt_saved_blocked_pivots() {
  const auto zero = [](std::optional<std::uint32_t> v) { return v.has_value() && *v == 0U; };
  for (std::size_t s = 0; s < sheets_.size(); ++s) {
    for (const std::unique_ptr<pivot::PivotTable>& table : sheets_[s].mutable_pivot_tables()) {
      if (table == nullptr || table->blocked().has_value() || !table->span_read() || table->span_rows() != 1U ||
          table->span_cols() != 1U || !zero(table->location_first_header_row()) ||
          !zero(table->location_first_data_row()) || !zero(table->location_first_data_col())) {
        continue;
      }
      const std::uint32_t row = table->anchor_row();
      const std::uint32_t col = table->anchor_col();
      const Cell* anchor = sheets_[s].cell_at(row, col);
      if (anchor == nullptr || !anchor->formula_text.empty() || !anchor->cached_value.is_error()) {
        continue;
      }
      pivot::BlockedReport state;
      state.size_known = false;
      table->set_projected_extent(pivot::PivotRect{row, col, row, col});
      table->set_blocked(state);
      put_cell_value(s, row, col, Value::error(ErrorCode::Spill));
    }
  }
}

std::size_t Workbook::settle_blocked_pivots() {
  std::size_t projected = 0;
  bool rendered = false;
  for (std::size_t s = 0; s < sheets_.size(); ++s) {
    if (count_blocked(sheets_[s].pivot_tables()) == 0U) {
      continue;
    }
    // Creation order decides which of two colliding reports renders: the earlier-created one recovers first.
    for (const std::uint32_t p : creation_order_of(sheets_[s])) {
      pivot::PivotTable* table = sheets_[s].mutable_pivot_tables()[p].get();
      if (table == nullptr || !table->blocked().has_value()) {
        continue;
      }
      const pivot::BlockedReport state = *table->blocked();
      if (state.size_known && !table->projection_pending()) {
        const std::uint32_t above = state.rows_above < table->anchor_row() ? state.rows_above : table->anchor_row();
        const std::uint64_t last_row = static_cast<std::uint64_t>(table->anchor_row()) - above + state.rows - 1U;
        const std::uint64_t last_col = static_cast<std::uint64_t>(table->anchor_col()) + state.cols - 1U;
        // A report clipped at the edge still reaches it until its anchor moves away from it.
        const bool at_edge = state.past_edge && (last_row + 1U >= Sheet::kMaxRows || last_col + 1U >= Sheet::kMaxCols);
        if (at_edge || last_row >= Sheet::kMaxRows || last_col >= Sheet::kMaxCols) {
          continue;
        }
        const MergeRange needed{table->anchor_row() - above, table->anchor_col(), static_cast<std::uint32_t>(last_row),
                                static_cast<std::uint32_t>(last_col)};
        if (detail::find_report_occupant(*this, s, p, needed, detail::pivot_occupied_extent(*table)).has_value()) {
          continue;
        }
      }
      if (project_pivot(s, p, ReportPlacement::kBlock)) {
        ++projected;
        if (!table->blocked().has_value()) {
          rendered = true;
          pivot::drop_derived_location(*table);
        }
      }
    }
  }
  // A recovered report's cells may be read through references that saw no cell write.
  if (rendered) {
    mark_all_formulas_dirty();
  }
  return projected;
}

void Workbook::remove_tables_meeting(std::size_t sheet_index, const MergeRange& rect) {
  std::vector<std::string> names;
  for (std::size_t i = tables_.size(); i-- > 0U;) {
    const std::optional<MergeRange> ref = detail::table_rect_on_sheet(tables_[i], sheet_index);
    if (!ref.has_value() || !intersects(normalized(*ref), rect)) {
      continue;
    }
    names.push_back(tables_[i].name);
    names.push_back(tables_[i].display_name);
    remove_table_slicers(*this, tables_[i].id);
    tables_.erase(tables_.begin() + static_cast<std::ptrdiff_t>(i));
  }
  if (names.empty()) {
    return;
  }
  reindex_formulas_for_table_change(names);
  mark_row_visibility_dependents_dirty();
}

Expected<void, Error> Workbook::move_pivot(std::size_t sheet_index, std::size_t pivot_index, std::uint32_t row,
                                           std::uint32_t col) {
  if (sheet_index >= sheets_.size() || pivot_index >= sheets_[sheet_index].pivot_tables().size() ||
      sheets_[sheet_index].pivot_tables()[pivot_index] == nullptr || !Sheet::coord_in_grid(row, col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "index or coordinate out of range",
                      "sheet=" + std::to_string(sheet_index) + " pivot=" + std::to_string(pivot_index));
  }
  pivot::PivotTable& table = *sheets_[sheet_index].mutable_pivot_tables()[pivot_index];
  const auto refuse = [&](const char* what) {
    return make_error(FormulonErrorCode::kPreconditionFailed, std::string(what),
                      "sheet=" + std::to_string(sheet_index) + " pivot=" + std::to_string(pivot_index));
  };
  // The destination is the rectangle the report holds now, moved with its anchor.
  const MergeRange own = detail::pivot_occupied_extent(table);
  const std::int64_t dr = static_cast<std::int64_t>(row) - table.anchor_row();
  const std::int64_t dc = static_cast<std::int64_t>(col) - table.anchor_col();
  const std::int64_t top = own.first_row + dr;
  const std::int64_t left = own.first_col + dc;
  const std::int64_t bottom = own.last_row + dr;
  const std::int64_t right = own.last_col + dc;
  if (top < 0 || left < 0 || bottom >= Sheet::kMaxRows || right >= Sheet::kMaxCols) {
    return refuse("The report would extend past the edge of the sheet");
  }
  const MergeRange dest{static_cast<std::uint32_t>(top), static_cast<std::uint32_t>(left),
                        static_cast<std::uint32_t>(bottom), static_cast<std::uint32_t>(right)};
  const Sheet& sheet = sheets_[sheet_index];
  if (pivot_in(sheet, pivot_index, dest).has_value()) {
    return refuse("The report would overlap another PivotTable");
  }
  // An array block or data table wholly inside is pasted over; one crossing the destination's edge refuses.
  for (const CellAddress address : sheet.formula_cells_in(0U, 0U, dest.last_row, dest.last_col)) {
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell == nullptr || cell->array_block_rows == 0U) {
      continue;
    }
    const MergeRange block{address.row, address.col, address.row + cell->array_block_rows - 1U,
                           address.col + cell->array_block_cols - 1U};
    if (intersects(block, dest) && !contains(dest, block)) {
      return refuse("The report would overlap an array formula");
    }
  }
  std::vector<CellAddress> pasted_data_tables;
  for (const DataTable& data_table : sheet.data_tables()) {
    const MergeRange ref = normalized(data_table.ref);
    if (!intersects(ref, dest)) {
      continue;
    }
    if (!contains(dest, ref)) {
      return refuse("The report would overlap a data table");
    }
    pasted_data_tables.push_back(CellAddress{ref.first_row, ref.first_col});
  }
  {
    std::unique_ptr<pivot::PivotTable> moved = table.clone();
    moved->set_anchor(row, col, 1U, 1U);
    RETURN_IF_ERROR(detail::plan_pivot_report(*this, *moved));
  }

  // The paste clears the destination, then the report leaves its old place.
  for (const CellAddress at : pasted_data_tables) {
    RETURN_IF_ERROR(remove_data_table(sheet_index, at.row, at.col));
  }
  remove_tables_meeting(sheet_index, dest);
  RETURN_IF_ERROR(remove_merges_intersecting(sheet_index, dest));
  // Pasting over a spill drops it; its producer reports #SPILL! on the next recalc.
  for (const CellAddress anchor : sheets_[sheet_index].committed_spill_anchors_intersecting(
           dest.first_row, dest.first_col, dest.last_row - dest.first_row + 1U, dest.last_col - dest.first_col + 1U)) {
    sheets_[sheet_index].clear_spill(anchor.row, anchor.col);
  }
  erase_cells_in(sheet_index, dest);
  std::vector<MergeRange> own_merges;
  for (const MergeRange& merge : sheets_[sheet_index].merges()) {
    if (contains(own, normalized(merge))) {
      own_merges.push_back(normalized(merge));
    }
  }
  for (const MergeRange& merge : own_merges) {
    RETURN_IF_ERROR(remove_merges_intersecting(sheet_index, merge));
  }
  erase_cells_in(sheet_index, own);
  table.mutable_projected_merges().clear();
  table.set_projected_extent(std::nullopt);
  table.set_anchor(row, col, 1U, 1U);
  pivot::clear_derived_state(table);
  mark_all_formulas_dirty();
  RETURN_IF_ERROR(project_pivot(sheet_index, pivot_index, ReportPlacement::kBlock));
  if (!table.blocked().has_value()) {
    pivot::drop_derived_location(table);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
