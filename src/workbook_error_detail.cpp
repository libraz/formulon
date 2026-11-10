// @size-budget: 3 KB
//
// Load-time restore of the rich errors a file saved: a blocked formula
// spill, an error cell's detail and a blocked PivotTable report.

#include <cstddef>
#include <cstdint>
#include <memory>

#include "pivot/pivot_layout.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {

namespace {

// The subType a rich `#SPILL!` carries for a report blocked past the sheet edge.
constexpr std::uint32_t kPastEdgeSubType = 3U;

}  // namespace

Expected<void, Error> Workbook::restore_blocked_spill(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                      std::uint32_t rows, std::uint32_t cols, SpillCause cause) {
  RETURN_IF_ERROR(detail::check_cell_target(sheet_index, sheets_.size(), row, col));
  if (rows == 0U || cols == 0U) {
    return make_error(FormulonErrorCode::kInvalidArgument, "a blocked spill needs a non-empty footprint");
  }
  sheets_[sheet_index].reject_spill_footprint(row, col, rows, cols, cause);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::restore_error_detail(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                     ErrorCode code, CellErrorDetail error_detail) {
  RETURN_IF_ERROR(detail::check_cell_target(sheet_index, sheets_.size(), row, col));
  Sheet& sheet = sheets_[sheet_index];
  sheet.set_cell_cached_value(row, col, Value::error(code));
  sheet.set_error_detail(row, col, error_detail);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::restore_blocked_pivot(std::size_t sheet_index, std::size_t pivot_index,
                                                      std::uint32_t body_rows, std::uint32_t cols,
                                                      std::uint32_t sub_type) {
  if (sheet_index >= sheets_.size() || pivot_index >= sheets_[sheet_index].pivot_tables().size() ||
      sheets_[sheet_index].pivot_tables()[pivot_index] == nullptr) {
    return make_error(FormulonErrorCode::kInvalidArgument, "index out of range");
  }
  pivot::PivotTable& table = *sheets_[sheet_index].mutable_pivot_tables()[pivot_index];
  const std::uint32_t row = table.anchor_row();
  const std::uint32_t col = table.anchor_col();
  // The page-field block sits above the body, pushed down only by an anchor too near the top.
  const std::uint32_t page_rows = pivot::page_block_rows(table, table.page_field_order().size());
  pivot::BlockedReport state;
  state.rows_above = page_rows < row ? page_rows : row;
  state.rows = body_rows + state.rows_above;
  state.cols = cols;
  state.past_edge = sub_type == kPastEdgeSubType;
  state.size_known = false;
  state.cause = state.past_edge ? SpillCause::kBeyondEdge : SpillCause::kOccupied;
  table.set_projected_extent(pivot::PivotRect{row, col, row, col});
  table.set_blocked(state);
  put_cell_value(sheet_index, row, col, Value::error(ErrorCode::Spill));
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
