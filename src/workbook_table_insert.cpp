// Table.insertRows / Table.insertColumns: a band insert over the table's own
// columns (rows), then the table's extent, column entries and calculated
// columns brought in line with it.

#include <cstddef>
#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <utility>

#include "auto_filter.h"
#include "eval/recalc_engine.h"
#include "merge_range.h"
#include "new_workbook_defaults.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {

namespace {

Error invalid_argument(std::string message, std::string context = {}) {
  Error err = make_error(FormulonErrorCode::kInvalidArgument, std::move(message));
  err.context = std::move(context);
  return err;
}

/// The extent of table `table_index`, which must be a table of `sheet_index`.
Expected<MergeRange, Error> table_extent(const Workbook& wb, std::size_t sheet_index, std::size_t table_index) {
  if (sheet_index >= wb.sheet_count()) {
    return invalid_argument("sheet index out of range", "sheet=" + std::to_string(sheet_index));
  }
  if (table_index >= wb.tables().size() || wb.tables()[table_index].sheet_index != sheet_index) {
    return invalid_argument("table index out of range", "table=" + std::to_string(table_index));
  }
  const TableMetadata& table = wb.tables()[table_index];
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (!rect.has_value()) {
    return invalid_argument("table ref is not a range", "ref=" + table.ref);
  }
  return *rect;
}

/// Checks the inclusive position span `[first, last]` against `count`
/// existing positions and the grid, for an insert that starts `start`
/// cells into an axis of `axis_extent` and pushes the table's far edge,
/// `table_last`, along by the span's length. Returns the span's length.
Expected<std::uint64_t, Error> check_insert_span(std::uint32_t first, std::uint32_t last, std::uint64_t count,
                                                 const char* noun, std::uint64_t start, std::uint64_t table_last,
                                                 std::uint64_t axis_extent) {
  if (first > count) {
    return invalid_argument("first " + std::to_string(first) + " is out of range (" + std::to_string(count) + " " +
                                noun + (count == 1U ? ")" : "s)"),
                            "first=" + std::to_string(first));
  }
  if (last < first) {
    return invalid_argument("last " + std::to_string(last) + " is before first " + std::to_string(first),
                            "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  const std::uint64_t length = static_cast<std::uint64_t>(last) - first + 1U;
  // An append against the grid's edge has no cell to start in: content would go off the grid.
  if (start >= axis_extent) {
    return make_error(FormulonErrorCode::kPreconditionFailed,
                      "first " + std::to_string(first) + " is outside the grid");
  }
  if (start + length > axis_extent) {
    return invalid_argument("last " + std::to_string(last) + " is outside the grid", "last=" + std::to_string(last));
  }
  if (table_last + length >= axis_extent) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "last would push the table off the grid");
  }
  return length;
}

}  // namespace

Expected<void, Error> Workbook::insert_table_rows(std::size_t sheet_index, std::size_t table_index, std::uint32_t first,
                                                  std::uint32_t last) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const Expected<MergeRange, Error> extent = table_extent(*this, sheet_index, table_index);
  if (!extent.has_value()) {
    return extent.error();
  }
  const MergeRange rect = extent.value();
  const bool has_header = tables_[table_index].header_row;
  const bool has_totals = tables_[table_index].totals_row;
  const std::uint64_t first_data = static_cast<std::uint64_t>(rect.first_row) + (has_header ? 1U : 0U);
  const std::uint64_t end_data = static_cast<std::uint64_t>(rect.last_row) + 1U - (has_totals ? 1U : 0U);
  if (end_data < first_data) {
    return invalid_argument("table has no room for its header and totals rows", "ref=" + tables_[table_index].ref);
  }
  const std::uint64_t data_rows = end_data - first_data;
  const std::uint64_t start = first_data + first;
  const Expected<std::uint64_t, Error> span =
      check_insert_span(first, last, data_rows, "row", start, rect.last_row, Sheet::kMaxRows);
  if (!span.has_value()) {
    return span.error();
  }
  const std::uint32_t count = static_cast<std::uint32_t>(span.value());
  const std::uint32_t band_first = static_cast<std::uint32_t>(start);
  RETURN_IF_ERROR(apply_band_edit_locked(
      sheet_index, MergeRange{band_first, rect.first_col, band_first + count - 1U, rect.last_col}, CellShift::kDown,
      /*is_delete=*/false));

  // The band edit fills the inserted rows itself when its rewrite grew the table.
  TableMetadata& table = tables_[table_index];
  const std::optional<MergeRange> rewritten = parse_a1_rectangle(table.ref);
  const bool filled =
      rewritten.has_value() && rewritten->first_row == rect.first_row && rewritten->last_row > rect.last_row;
  table.ref.clear();
  append_a1_rectangle(table.ref, MergeRange{rect.first_row, rect.first_col, rect.last_row + count, rect.last_col});
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  if (!filled) {
    detail::fill_calculated_columns_locked(*this, sheets_[sheet_index], mutator, table, band_first,
                                           band_first + count - 1U, 0U, table.columns.size() - 1U);
  }
  detail::finish_table_extent_change_locked(*this, sheets_, mutator, defined_names_, table_index);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::insert_table_columns(std::size_t sheet_index, std::size_t table_index,
                                                     std::uint32_t first, std::uint32_t last) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const Expected<MergeRange, Error> extent = table_extent(*this, sheet_index, table_index);
  if (!extent.has_value()) {
    return extent.error();
  }
  const MergeRange rect = extent.value();
  const std::uint64_t width = static_cast<std::uint64_t>(rect.last_col) - rect.first_col + 1U;
  const std::uint64_t start = static_cast<std::uint64_t>(rect.first_col) + first;
  const Expected<std::uint64_t, Error> span =
      check_insert_span(first, last, width, "column", start, rect.last_col, Sheet::kMaxCols);
  if (!span.has_value()) {
    return span.error();
  }
  const std::uint32_t count = static_cast<std::uint32_t>(span.value());
  const std::uint32_t band_first = static_cast<std::uint32_t>(start);
  const std::size_t columns_before = tables_[table_index].columns.size();
  RETURN_IF_ERROR(apply_band_edit_locked(
      sheet_index, MergeRange{rect.first_row, band_first, rect.last_row, band_first + count - 1U}, CellShift::kRight,
      /*is_delete=*/false));

  // The band edit adds the entries itself when the insert fell inside the table.
  TableMetadata& table = tables_[table_index];
  const MergeRange grown{rect.first_row, rect.first_col, rect.last_row, rect.last_col + count};
  table.ref.clear();
  append_a1_rectangle(table.ref, grown);
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  if (table.columns.size() == columns_before) {
    detail::add_table_columns_locked(table, first, count, grown, sheets_[sheet_index], sheet_index, mutator,
                                     new_workbook_defaults(excel_profile_.locale).column_word);
  }
  detail::finish_table_extent_change_locked(*this, sheets_, mutator, defined_names_, table_index);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
