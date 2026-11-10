// Stable C ABI argument range checks: every setter taking an index, count or
// level rejects an out-of-range value with kInvalidArgument and leaves the
// workbook unchanged.

#include <string>
#include <utility>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);

fm_sheet_view_info_t View(const WorkbookGuard& wb, std::size_t sheet) {
  fm_sheet_view_info_t view = FM_STRUCT_INIT(fm_sheet_view_info_t);
  EXPECT_EQ(fm_sheet_get_view(wb.handle, sheet, &view), 0) << fm_get_last_error_message();
  return view;
}

std::uint32_t ActiveSheet(const WorkbookGuard& wb) {
  std::size_t index = 99U;
  EXPECT_EQ(fm_workbook_get_active_sheet(wb.handle, &index), 0) << fm_get_last_error_message();
  return static_cast<std::uint32_t>(index);
}

std::size_t CellCount(const WorkbookGuard& wb) {
  std::size_t cells = 99U;
  EXPECT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &cells), 0) << fm_get_last_error_message();
  return cells;
}

fm_status_t SetFreeze(const WorkbookGuard& wb, std::size_t sheet, std::uint32_t rows, std::uint32_t cols) {
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.freeze_rows_op = FM_PATCH_OP_SET;
  patch.freeze_rows = rows;
  patch.freeze_cols_op = FM_PATCH_OP_SET;
  patch.freeze_cols = cols;
  return fm_sheet_update_view(wb.handle, sheet, &patch);
}

fm_status_t SetVisibility(const WorkbookGuard& wb, std::size_t sheet, std::int32_t visibility) {
  return fm_sheet_set_visibility(wb.handle, sheet, visibility);
}

std::int32_t Visibility(const WorkbookGuard& wb, std::size_t sheet) {
  std::int32_t visibility = -1;
  EXPECT_EQ(fm_sheet_get_visibility(wb.handle, sheet, &visibility), 0) << fm_get_last_error_message();
  return visibility;
}

fm_status_t SetTopLeft(const WorkbookGuard& wb, std::uint32_t row, std::uint32_t col) {
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.top_left_row_op = FM_PATCH_OP_SET;
  patch.top_left_row = row;
  patch.top_left_col_op = FM_PATCH_OP_SET;
  patch.top_left_col = col;
  return fm_sheet_update_view(wb.handle, 0, &patch);
}

void SaveAndReload(const WorkbookGuard& wb, std::int32_t format, WorkbookGuard* reloaded) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, format, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded->handle), 0) << fm_get_last_error_message();
}

}  // namespace

TEST(FormulonCApiArgumentRange, FreezeOutsideTheGridIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetFreeze(wb, 0, 2U, 1U), 0);

  EXPECT_EQ(SetFreeze(wb, 0, 0U, formulon::Sheet::kMaxCols), kInvalidArgument);
  EXPECT_EQ(SetFreeze(wb, 0, formulon::Sheet::kMaxRows, 0U), kInvalidArgument);
  EXPECT_EQ(SetFreeze(wb, 0, formulon::Sheet::kMaxRows, formulon::Sheet::kMaxCols), kInvalidArgument);
  const fm_sheet_view_info_t unchanged = View(wb, 0);
  EXPECT_EQ(unchanged.freeze_rows, 2U);
  EXPECT_EQ(unchanged.freeze_cols, 1U);
}

TEST(FormulonCApiArgumentRange, FreezeAtTheLastGridCellRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetFreeze(wb, 0, formulon::Sheet::kMaxRows - 1U, formulon::Sheet::kMaxCols - 1U), 0);

  WorkbookGuard reloaded;
  SaveAndReload(wb, FM_WORKBOOK_FORMAT_XLSX, &reloaded);
  const fm_sheet_view_info_t view = View(reloaded, 0);
  EXPECT_EQ(view.freeze_rows, formulon::Sheet::kMaxRows - 1U);
  EXPECT_EQ(view.freeze_cols, formulon::Sheet::kMaxCols - 1U);
}

TEST(FormulonCApiArgumentRange, OutlineLevelAboveSevenIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // 263 is 7 in its low byte, so a narrower parameter would accept it.
  for (const std::uint32_t level : {8U, 255U, 263U}) {
    EXPECT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 3U, 3U, level), kInvalidArgument) << level;
    EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 4U, level), kInvalidArgument) << level;
  }
  std::size_t rows = 99U;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &rows), 0);
  EXPECT_EQ(rows, 0U);
  std::size_t columns = 99U;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &columns), 0);
  EXPECT_EQ(columns, 0U);
}

TEST(FormulonCApiArgumentRange, OutlineLevelSevenRoundTripsThroughBothFormats) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 3U, 3U, 7U), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 2U, 7U), 0);

  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    WorkbookGuard reloaded;
    SaveAndReload(wb, format, &reloaded);
    fm_row_layout_info_t row = FM_STRUCT_INIT(fm_row_layout_info_t);
    ASSERT_EQ(fm_sheet_get_row_layout_at(reloaded.handle, 0, 0U, &row), 0) << format;
    EXPECT_EQ(row.row, 3U) << format;
    EXPECT_EQ(row.outline_level, 7U) << format;
    fm_column_layout_info_t column = FM_STRUCT_INIT(fm_column_layout_info_t);
    ASSERT_EQ(fm_sheet_get_column_layout_at(reloaded.handle, 0, 0U, &column), 0) << format;
    EXPECT_EQ(column.first, 2U) << format;
    EXPECT_EQ(column.outline_level, 7U) << format;
  }
}

TEST(FormulonCApiArgumentRange, CellXfIndexPastTheTableIsRejectedBeforeAnyWrite) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3U);
  std::size_t xf_count = 0U;
  ASSERT_EQ(fm_workbook_get_cell_xf_count(wb.handle, &xf_count), 0);
  ASSERT_EQ(xf_count, 3U);
  const auto count = static_cast<std::uint32_t>(xf_count);

  EXPECT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 1U, 1U, count), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 2U, 2U}), count),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(CellCount(wb), 0U);

  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 1U, 1U, count - 1U), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, record_arg(fm_cell_rect_t{3U, 0U, 4U, 0U}), count - 1U), 0)
      << fm_get_last_error_message();
  WorkbookGuard reloaded;
  SaveAndReload(wb, FM_WORKBOOK_FORMAT_XLSX, &reloaded);
  const std::pair<std::uint32_t, std::uint32_t> kStyled[] = {{1U, 1U}, {3U, 0U}, {4U, 0U}};
  for (const auto& [row, col] : kStyled) {
    std::int32_t present = 0;
    std::uint32_t xf = 0U;
    ASSERT_EQ(fm_sheet_get_cell_xf_index(reloaded.handle, 0, row, col, &present, &xf), 0);
    EXPECT_EQ(present, 1) << row << "," << col;
    EXPECT_EQ(xf, count - 1U) << row << "," << col;
  }
}

TEST(FormulonCApiArgumentRange, PivotDataFieldIndicesAreCheckedWithAnExplicitName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0U;
  std::size_t pivot = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot), 0) << fm_get_last_error_message();
  std::size_t field_count = 0U;
  ASSERT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0, pivot, &field_count), 0);
  ASSERT_EQ(field_count, 2U);

  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  spec.name = "Explicit";
  spec.field_index = 1U;
  spec.aggregation = FM_PIVOT_AGGREGATION_SUM;
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;

  fm_pivot_data_field_t bad_field = spec;
  bad_field.field_index = static_cast<std::uint32_t>(field_count);
  fm_pivot_data_field_t bad_base = spec;
  bad_base.has_show_as_base_field = 1;
  bad_base.show_as_base_field = static_cast<std::uint32_t>(field_count);
  fm_pivot_data_field_t far_base = spec;
  far_base.has_show_as_base_field = 1;
  far_base.show_as_base_field = 0xFFFFFFFEU;
  for (const fm_pivot_data_field_t* rejected : {&bad_field, &bad_base, &far_base}) {
    std::size_t idx = 99U;
    EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot, rejected, &idx), FM_STATUS_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot, 0U, rejected), FM_STATUS_INDEX_OUT_OF_RANGE);
  }
  std::size_t data_fields = 0U;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0, pivot, &data_fields), 0);
  EXPECT_EQ(data_fields, 1U);
  fm_pivot_data_field_t read = FM_STRUCT_INIT(fm_pivot_data_field_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0, pivot, 0U, &read), 0);
  EXPECT_STREQ(read.name, "Sum of Amount");

  spec.has_show_as_base_field = 1;
  spec.show_as_base_field = 0U;
  std::size_t added = 99U;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot, &spec, &added), 0) << fm_get_last_error_message();
  EXPECT_EQ(added, 1U);
}

TEST(FormulonCApiArgumentRange, TheOnlyVisibleSheetCannotBeHidden) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Hidden", append_position(wb.handle), ignored_index()), 0);
  ASSERT_EQ(SetVisibility(wb, 1, FM_SHEET_VISIBILITY_HIDDEN), 0);

  EXPECT_EQ(SetVisibility(wb, 0, FM_SHEET_VISIBILITY_HIDDEN), kPreconditionFailed);
  EXPECT_EQ(SetVisibility(wb, 0, FM_SHEET_VISIBILITY_VERY_HIDDEN), kPreconditionFailed);
  EXPECT_EQ(Visibility(wb, 0), FM_SHEET_VISIBILITY_VISIBLE);
  EXPECT_EQ(ActiveSheet(wb), 0U);

  // Re-stating a hidden sheet's state leaves the visible one alone, so it passes.
  EXPECT_EQ(SetVisibility(wb, 1, FM_SHEET_VISIBILITY_VERY_HIDDEN), 0);
}

TEST(FormulonCApiArgumentRange, TheOnlyVisibleSheetCannotBeRemoved) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Hidden", append_position(wb.handle), ignored_index()), 0);
  ASSERT_EQ(SetVisibility(wb, 1, FM_SHEET_VISIBILITY_HIDDEN), 0);

  EXPECT_EQ(fm_workbook_remove_sheet(wb.handle, 0), kPreconditionFailed);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 2U);
  EXPECT_EQ(Visibility(wb, 0), FM_SHEET_VISIBILITY_VISIBLE);
  EXPECT_EQ(ActiveSheet(wb), 0U);

  // A hidden sheet can still go, which leaves the visible one in place.
  EXPECT_EQ(fm_workbook_remove_sheet(wb.handle, 1), 0) << fm_get_last_error_message();
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);
}

TEST(FormulonCApiArgumentRange, CountHelpersClearThePreviousError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetFreeze(wb, 99U, 1U, 1U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  ASSERT_STRNE(fm_get_last_error_message(), "");

  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);
  EXPECT_STREQ(fm_get_last_error_message(), "");
  EXPECT_STREQ(fm_get_last_error_context(), "");

  fm_pagination_t* pagination = nullptr;
  ASSERT_EQ(fm_sheet_paginate(wb.handle, 0, &pagination), 0);
  ASSERT_EQ(SetFreeze(wb, 99U, 1U, 1U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(count_of(fm_pagination_get_page_count, pagination), 0U);
  EXPECT_STREQ(fm_get_last_error_message(), "");
  fm_pagination_destroy(pagination);
}

TEST(FormulonCApiArgumentRange, CountHelpersRejectANullHandle) {
  constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  std::size_t count = 99U;
  EXPECT_EQ(fm_workbook_get_sheet_count(nullptr, &count), kNullPointer);
  EXPECT_STRNE(fm_get_last_error_message(), "");
  EXPECT_EQ(fm_pagination_get_page_count(nullptr, &count), kNullPointer);
  EXPECT_STRNE(fm_get_last_error_message(), "");
  EXPECT_EQ(count, 99U);
}

namespace {

constexpr std::uint32_t kPastRow = formulon::Sheet::kMaxRows;
constexpr std::uint32_t kPastCol = formulon::Sheet::kMaxCols;
constexpr std::uint32_t kLastRow = formulon::Sheet::kMaxRows - 1U;
constexpr std::uint32_t kLastCol = formulon::Sheet::kMaxCols - 1U;
constexpr fm_error_code_t kErrorNa = static_cast<fm_error_code_t>(formulon::ErrorCode::NA);

}  // namespace

TEST(FormulonCApiArgumentRange, CellSettersRejectAnOutOfGridCoordinate) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const auto& [row, col] : {std::pair{kPastRow, 0U}, std::pair{0U, kPastCol}, std::pair{kPastRow, kPastCol}}) {
    EXPECT_EQ(fm_sheet_set_number(wb.handle, 0, row, col, 1.0), kInvalidArgument) << row << "," << col;
    EXPECT_EQ(fm_sheet_set_bool(wb.handle, 0, row, col, 1), kInvalidArgument) << row << "," << col;
    EXPECT_EQ(fm_sheet_set_error(wb.handle, 0, row, col, kErrorNa), kInvalidArgument) << row << "," << col;
    EXPECT_EQ(fm_sheet_set_text(wb.handle, 0, row, col, "x"), kInvalidArgument) << row << "," << col;
    EXPECT_EQ(fm_sheet_set_blank(wb.handle, 0, row, col), kInvalidArgument) << row << "," << col;
    EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, row, col, "=1"), kInvalidArgument) << row << "," << col;
  }
  EXPECT_EQ(CellCount(wb), 0U);
}

TEST(FormulonCApiArgumentRange, CellSettersAcceptTheLastGridCell) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_number(wb.handle, 0, kLastRow, kLastCol, 1.0), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_bool(wb.handle, 0, kLastRow, kLastCol, 1), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_error(wb.handle, 0, kLastRow, kLastCol, kErrorNa), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_text(wb.handle, 0, kLastRow, kLastCol, "x"), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_blank(wb.handle, 0, kLastRow, kLastCol), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, kLastRow, kLastCol, "=1"), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiArgumentRange, CellGettersRejectAnOutOfGridCoordinate) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const auto& [row, col] : {std::pair{kPastRow, 0U}, std::pair{0U, kPastCol}}) {
    const char* text = nullptr;
    std::size_t count = 0;
    fm_phonetic_run_t run{};
    fm_phonetic_properties_t properties = FM_STRUCT_INIT(fm_phonetic_properties_t);
    std::int32_t flag = 0;
    const char* lambda = nullptr;
    fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
    fm_data_table_t table = FM_STRUCT_INIT(fm_data_table_t);
    fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
    fm_comment_t comment = FM_STRUCT_INIT(fm_comment_t);
    std::uint32_t xf = 0;
    EXPECT_EQ(fm_sheet_get_cell_phonetic(wb.handle, 0, row, col, &flag, &text), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_phonetic_run_count(wb.handle, 0, row, col, &count), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_phonetic_run_at(wb.handle, 0, row, col, 0, &run), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_phonetic_properties(wb.handle, 0, row, col, &properties), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_show_phonetic(wb.handle, 0, row, col, &flag), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, row, col, &flag, &lambda), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_spill(wb.handle, 0, row, col, &flag, &spill), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_data_table(wb.handle, 0, row, col, &flag, &table), kInvalidArgument);
    EXPECT_EQ(fm_sheet_remove_data_table(wb.handle, 0, row, col), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, row, col, &flag, &link), kInvalidArgument);
    EXPECT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, row, col, nullptr), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_comment(wb.handle, 0, row, col, &flag, &comment), kInvalidArgument);
    EXPECT_EQ(fm_sheet_set_comment(wb.handle, 0, row, col, nullptr), kInvalidArgument);
    std::int32_t visible = 0;
    EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 0, row, col, &flag, &visible), kInvalidArgument);
    EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 0, row, col, 1), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, row, col, &flag, &xf), kInvalidArgument);
    EXPECT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, row, col), kInvalidArgument);
    fm_cell_nodes_t* nodes = nullptr;
    EXPECT_EQ(fm_sheet_get_precedents(wb.handle, 0, row, col, 1U, &nodes), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_dependents(wb.handle, 0, row, col, 1U, &nodes), kInvalidArgument);
    EXPECT_EQ(nodes, nullptr);
    fm_value_t proposed{};
    proposed.kind = FM_VALUE_KIND_NUMBER;
    proposed.u.number = 1.0;
    fm_validation_outcome_info_t outcome = FM_STRUCT_INIT(fm_validation_outcome_info_t);
    EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, row, col, &proposed, &outcome), kInvalidArgument);
    EXPECT_EQ(SetTopLeft(wb, row, col), kInvalidArgument);
  }
}

TEST(FormulonCApiArgumentRange, CellGettersAcceptTheLastGridCell) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* text = nullptr;
  std::size_t count = 0;
  fm_phonetic_properties_t properties = FM_STRUCT_INIT(fm_phonetic_properties_t);
  std::int32_t flag = 0;
  const char* lambda = nullptr;
  fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
  fm_data_table_t table = FM_STRUCT_INIT(fm_data_table_t);
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  fm_comment_t comment = FM_STRUCT_INIT(fm_comment_t);
  std::uint32_t xf = 0;
  EXPECT_EQ(fm_sheet_get_cell_phonetic(wb.handle, 0, kLastRow, kLastCol, &flag, &text), 0);
  EXPECT_EQ(fm_sheet_get_cell_phonetic_run_count(wb.handle, 0, kLastRow, kLastCol, &count), 0);
  EXPECT_EQ(fm_sheet_get_cell_phonetic_properties(wb.handle, 0, kLastRow, kLastCol, &properties), 0);
  EXPECT_EQ(properties.stated, 0);
  EXPECT_EQ(fm_sheet_get_cell_show_phonetic(wb.handle, 0, kLastRow, kLastCol, &flag), 0);
  EXPECT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, kLastRow, kLastCol, &flag, &lambda), 0);
  EXPECT_EQ(fm_sheet_get_spill(wb.handle, 0, kLastRow, kLastCol, &flag, &spill), 0);
  EXPECT_EQ(flag, 0);
  EXPECT_EQ(fm_sheet_get_data_table(wb.handle, 0, kLastRow, kLastCol, &flag, &table), 0);
  EXPECT_EQ(fm_sheet_remove_data_table(wb.handle, 0, kLastRow, kLastCol), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, kLastRow, kLastCol, &flag, &link), 0);
  EXPECT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, kLastRow, kLastCol, nullptr), 0);
  EXPECT_EQ(fm_sheet_get_comment(wb.handle, 0, kLastRow, kLastCol, &flag, &comment), 0);
  EXPECT_EQ(fm_sheet_set_comment(wb.handle, 0, kLastRow, kLastCol, nullptr), 0);
  std::int32_t visible = 0;
  EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 0, kLastRow, kLastCol, &flag, &visible), 0);
  EXPECT_EQ(flag, 0);
  EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 0, kLastRow, kLastCol, 1), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, kLastRow, kLastCol, &flag, &xf), 0);
  EXPECT_EQ(flag, 0);
  EXPECT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, kLastRow, kLastCol), 0);
  fm_cell_nodes_t* nodes = nullptr;
  ASSERT_EQ(fm_sheet_get_precedents(wb.handle, 0, kLastRow, kLastCol, 1U, &nodes), 0);
  fm_cell_nodes_destroy(nodes);
  nodes = nullptr;
  ASSERT_EQ(fm_sheet_get_dependents(wb.handle, 0, kLastRow, kLastCol, 1U, &nodes), 0);
  fm_cell_nodes_destroy(nodes);
  fm_value_t proposed{};
  proposed.kind = FM_VALUE_KIND_NUMBER;
  proposed.u.number = 1.0;
  fm_validation_outcome_info_t outcome = FM_STRUCT_INIT(fm_validation_outcome_info_t);
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, kLastRow, kLastCol, &proposed, &outcome), 0);
  EXPECT_EQ(SetTopLeft(wb, kLastRow, kLastCol), 0);
}

namespace {

constexpr fm_status_t kNotFound = static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound);
constexpr fm_status_t kAlreadyExists = static_cast<fm_status_t>(formulon::FormulonErrorCode::kAlreadyExists);
constexpr fm_status_t kIndexOutOfRange = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange);

fm_pivot_cache_options_t CacheOptions(const WorkbookGuard& wb, std::uint32_t cache_id) {
  fm_pivot_cache_options_t options = FM_STRUCT_INIT(fm_pivot_cache_options_t);
  EXPECT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id, &options), 0) << fm_get_last_error_message();
  return options;
}

fm_pivot_options_t PivotOptions(const WorkbookGuard& wb, std::size_t pivot) {
  fm_pivot_options_t options = FM_STRUCT_INIT(fm_pivot_options_t);
  EXPECT_EQ(fm_sheet_get_pivot_options(wb.handle, 0, pivot, &options), 0) << fm_get_last_error_message();
  return options;
}

}  // namespace

TEST(FormulonCApiArgumentRange, PivotCacheOptionsPatchWritesOnlyItsSetFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
  const fm_pivot_cache_options_t before = CacheOptions(wb, cache_id);
  EXPECT_EQ(before.has_missing_items_limit, 0);

  fm_pivot_cache_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  EXPECT_EQ(fm_workbook_update_pivot_cache_options(wb.handle, cache_id, &patch), 0) << fm_get_last_error_message();
  patch.missing_items_limit_op = FM_PATCH_OP_SET;
  patch.missing_items_limit = 32500U;
  ASSERT_EQ(fm_workbook_update_pivot_cache_options(wb.handle, cache_id, &patch), 0) << fm_get_last_error_message();
  fm_pivot_cache_options_t after = CacheOptions(wb, cache_id);
  EXPECT_EQ(after.has_missing_items_limit, 1);
  EXPECT_EQ(after.missing_items_limit, 32500U);
  EXPECT_EQ(after.save_data, before.save_data);
  EXPECT_EQ(after.refresh_on_load, before.refresh_on_load);

  patch.missing_items_limit_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_workbook_update_pivot_cache_options(wb.handle, cache_id, &patch), 0) << fm_get_last_error_message();
  const fm_pivot_cache_options_t cleared = CacheOptions(wb, cache_id);
  EXPECT_EQ(cleared.has_missing_items_limit, 0);
}

TEST(FormulonCApiArgumentRange, PivotCacheOptionsPatchRefusesABadFieldAndChangesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
  const fm_pivot_cache_options_t before = CacheOptions(wb, cache_id);

  fm_pivot_cache_options_patch_t clear_flag = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  clear_flag.save_data_op = FM_PATCH_OP_CLEAR;
  fm_pivot_cache_options_patch_t bad_op = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  bad_op.refresh_on_load_op = static_cast<fm_patch_op_t>(3);
  bad_op.refresh_on_load = 1;
  fm_pivot_cache_options_patch_t bad_limit = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  bad_limit.refresh_on_load_op = FM_PATCH_OP_SET;
  bad_limit.refresh_on_load = before.refresh_on_load != 0 ? 0 : 1;
  bad_limit.missing_items_limit_op = FM_PATCH_OP_SET;
  bad_limit.missing_items_limit = 1048577U;
  for (const fm_pivot_cache_options_patch_t* rejected : {&clear_flag, &bad_op, &bad_limit}) {
    EXPECT_EQ(fm_workbook_update_pivot_cache_options(wb.handle, cache_id, rejected), kInvalidArgument);
  }
  const fm_pivot_cache_options_t after = CacheOptions(wb, cache_id);
  EXPECT_EQ(after.save_data, before.save_data);
  EXPECT_EQ(after.has_missing_items_limit, before.has_missing_items_limit);
  EXPECT_EQ(after.missing_items_limit, before.missing_items_limit);
  EXPECT_EQ(after.refresh_on_load, before.refresh_on_load);
}

TEST(FormulonCApiArgumentRange, PivotOptionsPatchClearsACaptionAndRefusesAnUnclearableField) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot), 0) << fm_get_last_error_message();
  const fm_pivot_options_t before = PivotOptions(wb, pivot);

  fm_pivot_options_patch_t set_caption = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  set_caption.error_caption_op = FM_PATCH_OP_SET;
  set_caption.error_caption = "n/a";
  ASSERT_EQ(fm_sheet_update_pivot_options(wb.handle, 0, pivot, &set_caption), 0) << fm_get_last_error_message();
  fm_pivot_options_t after = PivotOptions(wb, pivot);
  ASSERT_NE(after.error_caption, nullptr);
  EXPECT_STREQ(after.error_caption, "n/a");
  EXPECT_EQ(after.indent, before.indent);
  EXPECT_EQ(after.show_error, before.show_error);

  fm_pivot_options_patch_t clear_caption = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  clear_caption.error_caption_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_pivot_options(wb.handle, 0, pivot, &clear_caption), 0) << fm_get_last_error_message();
  EXPECT_EQ(PivotOptions(wb, pivot).error_caption, nullptr);

  fm_pivot_options_patch_t clear_indent = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  clear_indent.indent_op = FM_PATCH_OP_CLEAR;
  fm_pivot_options_patch_t null_caption = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  null_caption.alt_text_op = FM_PATCH_OP_SET;
  fm_pivot_options_patch_t bad_indent = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  bad_indent.show_error_op = FM_PATCH_OP_SET;
  bad_indent.show_error = before.show_error != 0 ? 0 : 1;
  bad_indent.indent_op = FM_PATCH_OP_SET;
  bad_indent.indent = 127U;
  for (const fm_pivot_options_patch_t* rejected : {&clear_indent, &bad_indent}) {
    EXPECT_EQ(fm_sheet_update_pivot_options(wb.handle, 0, pivot, rejected), kInvalidArgument);
  }
  EXPECT_EQ(fm_sheet_update_pivot_options(wb.handle, 0, pivot, &null_caption), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_STREQ(fm_get_last_error_message(), "alt_text must not be NULL");
  after = PivotOptions(wb, pivot);
  EXPECT_EQ(after.show_error, before.show_error);
  EXPECT_EQ(after.indent, before.indent);
  EXPECT_EQ(after.alt_text, nullptr);
}

TEST(FormulonCApiArgumentRange, PivotGrandTotalsPatchKeepsTheOtherBand) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot), 0) << fm_get_last_error_message();
  fm_pivot_grand_totals_t totals = FM_STRUCT_INIT(fm_pivot_grand_totals_t);
  ASSERT_EQ(fm_sheet_get_pivot_grand_totals(wb.handle, 0, pivot, &totals), 0);
  const std::int32_t columns_before = totals.columns;

  fm_pivot_grand_totals_patch_t patch = FM_STRUCT_INIT(fm_pivot_grand_totals_patch_t);
  patch.rows_op = FM_PATCH_OP_SET;
  patch.rows = totals.rows != 0 ? 0 : 1;
  ASSERT_EQ(fm_sheet_update_pivot_grand_totals(wb.handle, 0, pivot, &patch), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_pivot_grand_totals(wb.handle, 0, pivot, &totals), 0);
  EXPECT_EQ(totals.rows, patch.rows);
  EXPECT_EQ(totals.columns, columns_before);

  fm_pivot_grand_totals_patch_t clear = FM_STRUCT_INIT(fm_pivot_grand_totals_patch_t);
  clear.columns_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_pivot_grand_totals(wb.handle, 0, pivot, &clear), kInvalidArgument);
  ASSERT_EQ(fm_sheet_get_pivot_grand_totals(wb.handle, 0, pivot, &totals), 0);
  EXPECT_EQ(totals.columns, columns_before);
}

TEST(FormulonCApiArgumentRange, PivotFieldOrderFillsTheCallerBufferAndReportsTheFullCount) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot), 0) << fm_get_last_error_message();

  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot, nullptr, 0U, &count), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(count, 1U);
  std::uint32_t order[2] = {77U, 77U};
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot, order, 2U, &count), 0);
  EXPECT_EQ(count, 1U);
  EXPECT_EQ(order[0], 0U);
  EXPECT_EQ(order[1], 77U);
  EXPECT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot, nullptr, 1U, &count),
            static_cast<fm_status_t>(FM_STATUS_BINDING_NULL_POINTER));

  ASSERT_EQ(fm_sheet_get_pivot_column_field_order(wb.handle, 0, pivot, nullptr, 0U, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiArgumentRange, PivotCacheStatusesFollowTheSectionRules) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();

  // A cache a pivot draws from cannot go; an id another cache holds is taken.
  EXPECT_EQ(fm_workbook_remove_pivot_cache(wb.handle, cache_id), kPreconditionFailed);
  std::uint32_t taken = 0U;
  EXPECT_EQ(fm_workbook_add_pivot_cache(wb.handle, cache_id, &taken), kAlreadyExists);

  // An unknown cache id is NOT_FOUND; an item past the shared items is INDEX_OUT_OF_RANGE.
  fm_pivot_cache_options_t options = FM_STRUCT_INIT(fm_pivot_cache_options_t);
  EXPECT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id + 100U, &options), kNotFound);
  std::int32_t is_date = -1;
  EXPECT_EQ(fm_workbook_is_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0U, 99U, &is_date),
            kIndexOutOfRange);
  fm_pivot_calculated_item_info_t item = FM_STRUCT_INIT(fm_pivot_calculated_item_info_t);
  EXPECT_EQ(fm_workbook_get_pivot_cache_calculated_item_at(wb.handle, cache_id, 0U, &item), kIndexOutOfRange);
}

TEST(FormulonCApiArgumentRange, ViewPatchWritesOnlyItsSetFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sheet_view_info_t before = View(wb, 0);
  const std::int32_t visibility_before = Visibility(wb, 0);

  fm_sheet_view_patch_t keep = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  ASSERT_EQ(fm_sheet_update_view(wb.handle, 0, &keep), 0) << fm_get_last_error_message();
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.show_grid_lines_op = FM_PATCH_OP_SET;
  patch.show_grid_lines = before.show_grid_lines != 0 ? 0 : 1;
  ASSERT_EQ(fm_sheet_update_view(wb.handle, 0, &patch), 0) << fm_get_last_error_message();

  const fm_sheet_view_info_t after = View(wb, 0);
  EXPECT_EQ(after.show_grid_lines, patch.show_grid_lines);
  EXPECT_EQ(after.zoom_pct, before.zoom_pct);
  EXPECT_EQ(after.show_zeros, before.show_zeros);
  EXPECT_EQ(after.freeze_rows, before.freeze_rows);
  EXPECT_EQ(after.view_mode, before.view_mode);
  EXPECT_EQ(Visibility(wb, 0), visibility_before);
}

TEST(FormulonCApiArgumentRange, ViewPatchRefusingOneFieldWritesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sheet_view_info_t before = View(wb, 0);

  fm_sheet_view_patch_t freeze_and_split = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  freeze_and_split.zoom_pct_op = FM_PATCH_OP_SET;
  freeze_and_split.zoom_pct = 150U;
  freeze_and_split.freeze_rows_op = FM_PATCH_OP_SET;
  freeze_and_split.freeze_rows = 1U;
  freeze_and_split.split_x_twips_op = FM_PATCH_OP_SET;
  freeze_and_split.split_x_twips = 1200.0;
  fm_sheet_view_patch_t clear = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  clear.zoom_pct_op = FM_PATCH_OP_CLEAR;
  fm_sheet_view_patch_t bad_op = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  bad_op.show_zeros_op = static_cast<fm_patch_op_t>(3);
  fm_sheet_view_patch_t bad_mode = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  bad_mode.zoom_pct_op = FM_PATCH_OP_SET;
  bad_mode.zoom_pct = 150U;
  bad_mode.view_mode_op = FM_PATCH_OP_SET;
  bad_mode.view_mode = 3;
  fm_sheet_view_patch_t pane_without_panes = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  pane_without_panes.pane_top_left_row_op = FM_PATCH_OP_SET;
  pane_without_panes.pane_top_left_row = 4U;
  for (const fm_sheet_view_patch_t* rejected : {&freeze_and_split, &clear, &bad_op, &bad_mode, &pane_without_panes}) {
    EXPECT_EQ(fm_sheet_update_view(wb.handle, 0, rejected), kInvalidArgument);
  }
  const fm_sheet_view_info_t after = View(wb, 0);
  EXPECT_EQ(after.zoom_pct, before.zoom_pct);
  EXPECT_EQ(after.freeze_rows, 0U);
  EXPECT_EQ(after.split_x_twips, 0.0);
  EXPECT_EQ(after.view_mode, before.view_mode);
}

TEST(FormulonCApiArgumentRange, ZoomOutsideTenToFourHundredIsRejectedNotClamped) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.zoom_pct_op = FM_PATCH_OP_SET;
  for (const std::uint32_t zoom : {10U, 400U}) {
    patch.zoom_pct = zoom;
    ASSERT_EQ(fm_sheet_update_view(wb.handle, 0, &patch), 0) << zoom << fm_get_last_error_message();
    EXPECT_EQ(View(wb, 0).zoom_pct, zoom);
  }
  for (const std::uint32_t zoom : {0U, 9U, 401U}) {
    patch.zoom_pct = zoom;
    EXPECT_EQ(fm_sheet_update_view(wb.handle, 0, &patch), kInvalidArgument) << zoom;
    EXPECT_EQ(View(wb, 0).zoom_pct, 400U) << zoom;
  }
}

TEST(FormulonCApiArgumentRange, RowHeightAndColumnWidthStopAtExcelsLimits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, 1U, 1U, 409.0), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_row_height(wb.handle, 0, 2U, 2U, 409.25), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_row_auto_height(wb.handle, 0, 2U, 2U, 409.25), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 1U, 255.0), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_column_width(wb.handle, 0, 2U, 2U, 255.5), kInvalidArgument);

  std::size_t rows = 99U;
  ASSERT_EQ(fm_sheet_get_row_layout_count(wb.handle, 0, &rows), 0);
  EXPECT_EQ(rows, 1U);
  std::size_t columns = 99U;
  ASSERT_EQ(fm_sheet_get_column_layout_count(wb.handle, 0, &columns), 0);
  EXPECT_EQ(columns, 1U);
}

TEST(FormulonCApiArgumentRange, FontSizeOutsideOneToFourHundredNineIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_font_t font{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, 0, &font), 0);
  // The record's string views are transient, so the name is kept by value.
  const std::string font_name = font.name != nullptr ? font.name : "";
  font.name = font_name.c_str();
  const std::size_t fonts = count_of(fm_workbook_get_font_count, wb.handle);
  for (const double size : {0.5, 409.5, -1.0}) {
    fm_font_t sized = font;
    sized.size_pt = size;
    sized.has_size_pt = 1;
    std::uint32_t index = 99U;
    EXPECT_EQ(fm_workbook_add_font(wb.handle, &sized, &index), kInvalidArgument) << size;
    EXPECT_EQ(fm_workbook_set_font(wb.handle, 0, &sized), kInvalidArgument) << size;
    fm_dxf_t dxf{};
    dxf.has_font = 1;
    dxf.font = sized;
    dxf.font.has_size_pt = 1;
    std::uint32_t dxf_index = 99U;
    EXPECT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &dxf_index), kInvalidArgument) << size;
  }
  EXPECT_EQ(count_of(fm_workbook_get_font_count, wb.handle), fonts);
  for (const double size : {1.0, 409.0}) {
    fm_font_t sized = font;
    sized.size_pt = size;
    sized.has_size_pt = 1;
    std::uint32_t index = 99U;
    EXPECT_EQ(fm_workbook_add_font(wb.handle, &sized, &index), 0) << size << fm_get_last_error_message();
  }
}

TEST(FormulonCApiArgumentRange, FormatDefaultsPatchClearsADefaultAndKeepsTheRest) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_format_defaults_t before = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &before), 0);
  fm_sheet_format_defaults_patch_t set = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  set.default_row_height_pt_op = FM_PATCH_OP_SET;
  set.default_row_height_pt = 20.0;
  set.zero_height_op = FM_PATCH_OP_SET;
  set.zero_height = 1;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &set), 0) << fm_get_last_error_message();
  fm_sheet_format_defaults_t read = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_default_row_height_pt, 1);
  EXPECT_EQ(read.default_row_height_pt, 20.0);
  EXPECT_EQ(read.zero_height, 1);

  fm_sheet_format_defaults_patch_t clear = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  clear.default_row_height_pt_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &clear), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_default_row_height_pt, 0);
  EXPECT_EQ(read.zero_height, 1);

  fm_sheet_format_defaults_patch_t clear_flag = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  clear_flag.zero_height_op = FM_PATCH_OP_CLEAR;
  fm_sheet_format_defaults_patch_t negative = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  negative.custom_height_op = FM_PATCH_OP_SET;
  negative.custom_height = 1;
  negative.base_column_width_chars_op = FM_PATCH_OP_SET;
  negative.base_column_width_chars = -1.0;
  for (const fm_sheet_format_defaults_patch_t* rejected : {&clear_flag, &negative}) {
    EXPECT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, rejected), kInvalidArgument);
  }
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.zero_height, 1);
  EXPECT_EQ(read.custom_height, before.custom_height);
  EXPECT_EQ(read.base_column_width_chars, before.base_column_width_chars);
}

TEST(FormulonCApiArgumentRange, TableStyleDefaultsPatchSetsAndClearsEachName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_table_style_defaults_t before = FM_STRUCT_INIT(fm_table_style_defaults_t);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &before), 0);
  const std::string pivot_default = before.pivot_style_name != nullptr ? before.pivot_style_name : "";
  fm_table_style_defaults_patch_t set = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  set.table_style_name_op = FM_PATCH_OP_SET;
  set.table_style_name = "TableStyleLight9";
  ASSERT_EQ(fm_workbook_update_table_style_defaults(wb.handle, &set), 0) << fm_get_last_error_message();
  fm_table_style_defaults_t read = FM_STRUCT_INIT(fm_table_style_defaults_t);
  read.pivot_style_name = "stale";
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &read), 0);
  ASSERT_NE(read.table_style_name, nullptr);
  EXPECT_STREQ(read.table_style_name, "TableStyleLight9");
  EXPECT_EQ(read.pivot_style_name != nullptr ? std::string(read.pivot_style_name) : std::string(), pivot_default);

  fm_table_style_defaults_patch_t null_name = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  null_name.pivot_style_name_op = FM_PATCH_OP_SET;
  EXPECT_EQ(fm_workbook_update_table_style_defaults(wb.handle, &null_name), FM_STATUS_INVALID_ARGUMENT);

  fm_table_style_defaults_patch_t clear = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  clear.table_style_name_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_workbook_update_table_style_defaults(wb.handle, &clear), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &read), 0);
  EXPECT_EQ(read.table_style_name, nullptr);
}

TEST(FormulonCApiArgumentRange, IterationCapAboveExcelLimitIsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  for (const std::uint32_t rejected : {40000U, 0U}) {
    fm_iterative_settings_patch_t patch = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
    patch.max_iterations_op = FM_PATCH_OP_SET;
    patch.max_iterations = rejected;
    EXPECT_EQ(fm_workbook_update_iterative(wb.handle, &patch), kInvalidArgument) << rejected;
  }

  fm_iterative_settings_patch_t limit = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  limit.max_iterations_op = FM_PATCH_OP_SET;
  limit.max_iterations = 32767U;
  EXPECT_EQ(fm_workbook_update_iterative(wb.handle, &limit), 0) << fm_get_last_error_message();
  fm_iterative_settings_info_t info = FM_STRUCT_INIT(fm_iterative_settings_info_t);
  ASSERT_EQ(fm_workbook_get_iterative(wb.handle, &info), 0) << fm_get_last_error_message();
  EXPECT_EQ(info.max_iterations, 32767U);
}
