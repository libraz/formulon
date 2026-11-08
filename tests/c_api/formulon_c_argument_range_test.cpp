// Stable C ABI argument range checks: every setter taking an index, count or
// level rejects an out-of-range value with kInvalidArgument and leaves the
// workbook unchanged.

#include <utility>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

fm_sheet_view_t View(const WorkbookGuard& wb, std::size_t sheet) {
  fm_sheet_view_t view = FM_STRUCT_INIT(fm_sheet_view_t);
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

void SaveAndReload(const WorkbookGuard& wb, std::int32_t format, WorkbookGuard* reloaded) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save_as(wb.handle, format, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded->handle), 0) << fm_get_last_error_message();
}

}  // namespace

TEST(FormulonCApiArgumentRange, FreezeOutsideTheGridIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 0, 2U, 1U), 0);

  EXPECT_EQ(fm_sheet_set_freeze(wb.handle, 0, 0U, formulon::Sheet::kMaxCols), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_freeze(wb.handle, 0, formulon::Sheet::kMaxRows, 0U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_freeze(wb.handle, 0, formulon::Sheet::kMaxRows, formulon::Sheet::kMaxCols), kInvalidArgument);
  const fm_sheet_view_t unchanged = View(wb, 0);
  EXPECT_EQ(unchanged.freeze_rows, 2U);
  EXPECT_EQ(unchanged.freeze_cols, 1U);
}

TEST(FormulonCApiArgumentRange, FreezeAtTheLastGridCellRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, formulon::Sheet::kMaxCols - 1U), 0);

  WorkbookGuard reloaded;
  SaveAndReload(wb, FM_WORKBOOK_FORMAT_XLSX, &reloaded);
  const fm_sheet_view_t view = View(reloaded, 0);
  EXPECT_EQ(view.freeze_rows, formulon::Sheet::kMaxRows - 1U);
  EXPECT_EQ(view.freeze_cols, formulon::Sheet::kMaxCols - 1U);
}

TEST(FormulonCApiArgumentRange, OutlineLevelAboveSevenIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const std::uint8_t level : {std::uint8_t{8U}, std::uint8_t{255U}}) {
    EXPECT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 3U, 3U, level), kInvalidArgument) << int{level};
    EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 4U, level), kInvalidArgument) << int{level};
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
    fm_row_layout_t row = FM_STRUCT_INIT(fm_row_layout_t);
    ASSERT_EQ(fm_sheet_get_row_layout_at(reloaded.handle, 0, 0U, &row), 0) << format;
    EXPECT_EQ(row.row, 3U) << format;
    EXPECT_EQ(row.outline_level, 7U) << format;
    fm_column_layout_t column = FM_STRUCT_INIT(fm_column_layout_t);
    ASSERT_EQ(fm_sheet_get_column_layout_at(reloaded.handle, 0, 0U, &column), 0) << format;
    EXPECT_EQ(column.first, 2U) << format;
    EXPECT_EQ(column.outline_level, 7U) << format;
  }
}

TEST(FormulonCApiArgumentRange, CellXfIndexPastTheTableIsRejectedBeforeAnyWrite) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3U);
  std::size_t count = 0U;
  ASSERT_EQ(fm_workbook_get_cell_xf_count(wb.handle, &count), 0);
  ASSERT_EQ(count, 3U);

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
    std::size_t xf = 0U;
    ASSERT_EQ(fm_sheet_get_cell_xf_index(reloaded.handle, 0, row, col, &xf), 0);
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

  fm_pivot_data_field_spec_t spec = FM_STRUCT_INIT(fm_pivot_data_field_spec_t);
  spec.name = "Explicit";
  spec.field_index = 1U;
  spec.aggregation = FM_PIVOT_AGGREGATION_SUM;
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  spec.show_as_base_field = -1;
  spec.show_as_base_item = -1;

  fm_pivot_data_field_spec_t bad_field = spec;
  bad_field.field_index = static_cast<std::uint32_t>(field_count);
  fm_pivot_data_field_spec_t bad_base = spec;
  bad_base.show_as_base_field = static_cast<std::int32_t>(field_count);
  fm_pivot_data_field_spec_t negative_base = spec;
  negative_base.show_as_base_field = -2;
  for (const fm_pivot_data_field_spec_t* rejected : {&bad_field, &bad_base, &negative_base}) {
    std::size_t idx = 99U;
    EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot, rejected, &idx), FM_STATUS_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot, 0U, rejected), FM_STATUS_INDEX_OUT_OF_RANGE);
  }
  std::size_t data_fields = 0U;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0, pivot, &data_fields), 0);
  EXPECT_EQ(data_fields, 1U);
  fm_pivot_data_field_spec_t read = FM_STRUCT_INIT(fm_pivot_data_field_spec_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field(wb.handle, 0, pivot, 0U, &read), 0);
  EXPECT_STREQ(read.name, "Sum of Amount");

  spec.show_as_base_field = 0;
  std::size_t added = 99U;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot, &spec, &added), 0) << fm_get_last_error_message();
  EXPECT_EQ(added, 1U);
}

TEST(FormulonCApiArgumentRange, TheOnlyVisibleSheetCannotBeHidden) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Hidden", ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 1, FM_SHEET_VISIBILITY_HIDDEN), 0);

  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VISIBILITY_HIDDEN), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VISIBILITY_VERY_HIDDEN), kInvalidArgument);
  EXPECT_EQ(View(wb, 0).visibility, FM_SHEET_VISIBILITY_VISIBLE);
  EXPECT_EQ(ActiveSheet(wb), 0U);

  // Re-stating a hidden sheet's state leaves the visible one alone, so it passes.
  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 1, FM_SHEET_VISIBILITY_VERY_HIDDEN), 0);
}

TEST(FormulonCApiArgumentRange, TheOnlyVisibleSheetCannotBeRemoved) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Hidden", ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 1, FM_SHEET_VISIBILITY_HIDDEN), 0);

  EXPECT_EQ(fm_workbook_remove_sheet(wb.handle, 0), kInvalidArgument);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 2U);
  EXPECT_EQ(View(wb, 0).visibility, FM_SHEET_VISIBILITY_VISIBLE);
  EXPECT_EQ(ActiveSheet(wb), 0U);

  // A hidden sheet can still go, which leaves the visible one in place.
  EXPECT_EQ(fm_workbook_remove_sheet(wb.handle, 1), 0) << fm_get_last_error_message();
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);
}

TEST(FormulonCApiArgumentRange, CountHelpersClearThePreviousError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 99U, 1U, 1U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  ASSERT_STRNE(fm_get_last_error_message(), "");

  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);
  EXPECT_STREQ(fm_get_last_error_message(), "");
  EXPECT_STREQ(fm_get_last_error_context(), "");

  fm_pagination_t* pagination = nullptr;
  ASSERT_EQ(fm_sheet_get_pagination(wb.handle, 0, &pagination), 0);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 99U, 1U, 1U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
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
    std::size_t font = 0;
    std::int32_t type = 0;
    std::int32_t alignment = 0;
    std::int32_t flag = 0;
    const char* lambda = nullptr;
    fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
    fm_data_table_t table = FM_STRUCT_INIT(fm_data_table_t);
    fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
    fm_comment_t comment = FM_STRUCT_INIT(fm_comment_t);
    std::size_t xf = 0;
    EXPECT_EQ(fm_sheet_get_cell_phonetic(wb.handle, 0, row, col, &text), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_phonetic_run_count(wb.handle, 0, row, col, &count), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_phonetic_run(wb.handle, 0, row, col, 0, &run), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_phonetic_properties(wb.handle, 0, row, col, &font, &type, &alignment),
              kInvalidArgument);
    EXPECT_EQ(fm_sheet_has_cell_phonetic_properties(wb.handle, 0, row, col, &flag), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_show_phonetic(wb.handle, 0, row, col, &flag), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, row, col, &flag, &lambda), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_spill_info(wb.handle, 0, row, col, &spill), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_data_table(wb.handle, 0, row, col, &flag, &table), kInvalidArgument);
    EXPECT_EQ(fm_sheet_remove_data_table(wb.handle, 0, row, col), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, row, col, &flag, &link), kInvalidArgument);
    EXPECT_EQ(fm_sheet_remove_hyperlink(wb.handle, 0, row, col), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_comment(wb.handle, 0, row, col, &flag, &comment), kInvalidArgument);
    EXPECT_EQ(fm_sheet_remove_comment(wb.handle, 0, row, col), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 0, row, col, &flag), kInvalidArgument);
    EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 0, row, col, 1), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, row, col, &xf), kInvalidArgument);
    EXPECT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, row, col), kInvalidArgument);
    fm_cell_nodes_t* nodes = nullptr;
    EXPECT_EQ(fm_sheet_get_precedents(wb.handle, 0, row, col, 1U, &nodes), kInvalidArgument);
    EXPECT_EQ(fm_sheet_get_dependents(wb.handle, 0, row, col, 1U, &nodes), kInvalidArgument);
    EXPECT_EQ(nodes, nullptr);
    fm_value_t proposed{};
    proposed.kind = FM_VALUE_KIND_NUMBER;
    proposed.u.number = 1.0;
    fm_validation_outcome_t outcome = FM_STRUCT_INIT(fm_validation_outcome_t);
    EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, row, col, &proposed, &outcome), kInvalidArgument);
    EXPECT_EQ(fm_sheet_set_top_left(wb.handle, 0, row, col), kInvalidArgument);
  }
}

TEST(FormulonCApiArgumentRange, CellGettersAcceptTheLastGridCell) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* text = nullptr;
  std::size_t count = 0;
  std::size_t font = 0;
  std::int32_t type = 0;
  std::int32_t alignment = 0;
  std::int32_t flag = 0;
  const char* lambda = nullptr;
  fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
  fm_data_table_t table = FM_STRUCT_INIT(fm_data_table_t);
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  fm_comment_t comment = FM_STRUCT_INIT(fm_comment_t);
  std::size_t xf = 0;
  EXPECT_EQ(fm_sheet_get_cell_phonetic(wb.handle, 0, kLastRow, kLastCol, &text), 0);
  EXPECT_EQ(fm_sheet_get_cell_phonetic_run_count(wb.handle, 0, kLastRow, kLastCol, &count), 0);
  EXPECT_EQ(fm_sheet_get_cell_phonetic_properties(wb.handle, 0, kLastRow, kLastCol, &font, &type, &alignment), 0);
  EXPECT_EQ(fm_sheet_has_cell_phonetic_properties(wb.handle, 0, kLastRow, kLastCol, &flag), 0);
  EXPECT_EQ(fm_sheet_get_cell_show_phonetic(wb.handle, 0, kLastRow, kLastCol, &flag), 0);
  EXPECT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, kLastRow, kLastCol, &flag, &lambda), 0);
  EXPECT_EQ(fm_sheet_get_spill_info(wb.handle, 0, kLastRow, kLastCol, &spill), 0);
  EXPECT_EQ(fm_sheet_get_data_table(wb.handle, 0, kLastRow, kLastCol, &flag, &table), 0);
  EXPECT_EQ(fm_sheet_remove_data_table(wb.handle, 0, kLastRow, kLastCol), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, kLastRow, kLastCol, &flag, &link), 0);
  EXPECT_EQ(fm_sheet_remove_hyperlink(wb.handle, 0, kLastRow, kLastCol), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_comment(wb.handle, 0, kLastRow, kLastCol, &flag, &comment), 0);
  EXPECT_EQ(fm_sheet_remove_comment(wb.handle, 0, kLastRow, kLastCol), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 0, kLastRow, kLastCol, &flag), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 0, kLastRow, kLastCol, 1), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, kLastRow, kLastCol, &xf), 0);
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
  fm_validation_outcome_t outcome = FM_STRUCT_INIT(fm_validation_outcome_t);
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, kLastRow, kLastCol, &proposed, &outcome), 0);
  EXPECT_EQ(fm_sheet_set_top_left(wb.handle, 0, kLastRow, kLastCol), 0);
}
