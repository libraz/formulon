// Header-row toggle of a worksheet table: turning the header row off and on.

#include <cassert>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "cf/auto_filter_eval.h"
#include "eval/recalc_engine.h"
#include "io/auto_filter_xml.h"
#include "io/xml_escape.h"
#include "merge_range.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {

// @size-budget: 4 KB
Expected<std::optional<PendingShift>, Error> Workbook::check_table_header_row(
    std::size_t index, bool shown, const TableShape& assumed, const std::vector<PendingShift>& pending) const {
  if (index >= tables_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "index out of range", "index=" + std::to_string(index));
  }
  if (shown == assumed.header_row) {
    return std::optional<PendingShift>();
  }
  const TableMetadata& table = tables_[index];
  const std::size_t sheet_index = table.sheet_index;
  const MergeRange& rect = assumed.rect;
  if (!shown) {
    const std::uint32_t rows = rect.last_row - rect.first_row + 1U;
    if (rows <= 1U + (assumed.totals_row ? 1U : 0U)) {
      return make_error(FormulonErrorCode::kPreconditionFailed,
                        "patch.header_row cannot be hidden when the table has no data row", "ref=" + table.ref);
    }
    return std::optional<PendingShift>();
  }
  // The names go into an empty row directly above the data when the grown range overlaps nothing.
  if (rect.first_row > 0U) {
    const std::uint32_t above = rect.first_row - 1U;
    bool empty = true;
    for (std::uint32_t col = rect.first_col; col <= rect.last_col && empty; ++col) {
      // A cell a pending shift inserts is empty.
      const std::optional<CellAddress> at = detail::cell_before_shifts(sheet_index, above, col, pending);
      empty = !at.has_value() || !detail::cell_is_occupied(sheets_[sheet_index], at->row, at->col);
    }
    MergeRange grown = rect;
    grown.first_row = above;
    if (empty && detail::check_table_range_free(*this, sheet_index, grown, index, /*unmerge_inside=*/false, pending)
                     .has_value()) {
      return std::optional<PendingShift>();
    }
  }
  const MergeRange band{rect.first_row, rect.first_col, rect.first_row, rect.last_col};
  RETURN_IF_ERROR(detail::check_band_edit_after(*this, sheet_index, band, CellShift::kDown, pending));
  return std::optional<PendingShift>(PendingShift{sheet_index, band, CellShift::kDown});
}

Expected<void, Error> Workbook::set_table_header_row(std::size_t index, bool shown) {
  if (index >= tables_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "index out of range", "index=" + std::to_string(index));
  }
  const std::optional<MergeRange> rect = parse_a1_rectangle(tables_[index].ref);
  if (!rect) {
    return make_error(FormulonErrorCode::kInvalidArgument, "table ref is not a range", "ref=" + tables_[index].ref);
  }
  if (shown == tables_[index].header_row) {
    return Expected<void, Error>::Ok();
  }
  const std::size_t sheet_index = tables_[index].sheet_index;
  // OFF drops the filter with its criteria and shows the rows they hid; a filter without criteria hid none.
  std::optional<AutoFilter> shown_rows;
  if (!shown) {
    const AutoFilter* filter = tables_[index].auto_filter_xml.get();
    if (filter != nullptr && filter->has_criteria()) {
      shown_rows.emplace();
      shown_rows->range = filter->range;
    }
  }
  RETURN_IF_ERROR(set_table_header_row_cells(index, shown, *rect));
  if (shown_rows.has_value()) {
    return apply_auto_filter(*this, sheet_index, *shown_rows);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_table_header_row_cells(std::size_t index, bool shown, const MergeRange& rect) {
  const std::size_t sheet_index = tables_[index].sheet_index;
  // Cell writes, the optional insert and the range change are one edit under a single hold of the engine mutex.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const TableShape current{rect, tables_[index].header_row, tables_[index].totals_row};
  ASSIGN_OR_RETURN(const std::optional<PendingShift> shift, check_table_header_row(index, shown, current, {}));
  MergeRange next = rect;
  if (!shown) {
    // The header cells go outright, formats included, and nothing moves.
    for (std::uint32_t col = rect.first_col; col <= rect.last_col; ++col) {
      if (sheets_[sheet_index].cell_at(rect.first_row, col) != nullptr) {
        put_cell_value_locked(sheet_index, rect.first_row, col, Value::blank());
        sheets_[sheet_index].erase_cell(rect.first_row, col);
      }
    }
    next.first_row = rect.first_row + 1U;
  } else if (shift.has_value()) {
    // The insert moves the whole table down one row; the header takes the vacated first data row.
    const Expected<void, Error> inserted = apply_band_edit_locked(sheet_index, shift->range, shift->shift, false);
    assert(inserted.has_value());
    RETURN_IF_ERROR(inserted);
    next.last_row = rect.last_row + 1U;
  } else {
    next.first_row = rect.first_row - 1U;
  }
  TableMetadata& table = tables_[index];
  table.ref.clear();
  append_a1_rectangle(table.ref, next);
  table.header_row = shown;
  if (shown) {
    std::string escaped_ref;
    io::AppendXmlAttrEscaped(escaped_ref, table.ref);
    table.auto_filter_xml.set(io::auto_filter_from_xml("<autoFilter ref=\"" + escaped_ref + "\"/>"));
    for (std::size_t i = 0; i < table.columns.size(); ++i) {
      put_cell_text_locked(sheet_index, next.first_row, next.first_col + static_cast<std::uint32_t>(i),
                           table.columns[i].name);
    }
  } else {
    table.auto_filter_xml.reset();
  }
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  detail::finish_table_extent_change_locked(*this, sheets_, mutator, defined_names_, index);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
