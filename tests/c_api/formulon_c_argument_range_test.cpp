// Stable C ABI argument range checks: every setter taking an index, count or
// level rejects an out-of-range value with kInvalidArgument and leaves the
// workbook unchanged.

#include <utility>

#include "formulon_c_pivot_test_helpers.h"
#include "sheet.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

fm_sheet_view_t View(const WorkbookGuard& wb, std::size_t sheet) {
  fm_sheet_view_t view{};
  EXPECT_EQ(fm_sheet_get_view(wb.handle, sheet, &view), 0) << fm_last_error_message();
  return view;
}

std::uint32_t ActiveSheet(const WorkbookGuard& wb) {
  std::uint32_t index = 99U;
  EXPECT_EQ(fm_workbook_get_active_sheet(wb.handle, &index), 0) << fm_last_error_message();
  return index;
}

std::size_t CellCount(const WorkbookGuard& wb) {
  std::size_t cells = 99U;
  EXPECT_EQ(fm_workbook_cell_count(wb.handle, 0, &cells), 0) << fm_last_error_message();
  return cells;
}

void SaveAndReload(const WorkbookGuard& wb, std::int32_t format, WorkbookGuard* reloaded) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save_as(wb.handle, format, &saved.data, &saved.len), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded->handle), 0) << fm_last_error_message();
}

}  // namespace

TEST(FormulonCApiArgumentRange, FreezeOutsideTheGridIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, formulon::Sheet::kMaxCols - 1U), 0);

  WorkbookGuard reloaded;
  SaveAndReload(wb, FM_WORKBOOK_FORMAT_XLSX, &reloaded);
  const fm_sheet_view_t view = View(reloaded, 0);
  EXPECT_EQ(view.freeze_rows, formulon::Sheet::kMaxRows - 1U);
  EXPECT_EQ(view.freeze_cols, formulon::Sheet::kMaxCols - 1U);
}

TEST(FormulonCApiArgumentRange, OutlineLevelAboveSevenIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  for (const std::uint8_t level : {std::uint8_t{8U}, std::uint8_t{255U}}) {
    EXPECT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 3U, level), kInvalidArgument) << int{level};
    EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 4U, level), kInvalidArgument) << int{level};
  }
  std::size_t rows = 99U;
  ASSERT_EQ(fm_sheet_get_row_override_count(wb.handle, 0, &rows), 0);
  EXPECT_EQ(rows, 0U);
  std::size_t columns = 99U;
  ASSERT_EQ(fm_sheet_get_column_count(wb.handle, 0, &columns), 0);
  EXPECT_EQ(columns, 0U);
}

TEST(FormulonCApiArgumentRange, OutlineLevelSevenRoundTripsThroughBothFormats) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 3U, 7U), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 2U, 7U), 0);

  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    WorkbookGuard reloaded;
    SaveAndReload(wb, format, &reloaded);
    fm_row_layout_t row{};
    ASSERT_EQ(fm_sheet_get_row_override(reloaded.handle, 0, 0U, &row), 0) << format;
    EXPECT_EQ(row.row, 3U) << format;
    EXPECT_EQ(row.outline_level, 7U) << format;
    fm_column_layout_t column{};
    ASSERT_EQ(fm_sheet_get_column(reloaded.handle, 0, 0U, &column), 0) << format;
    EXPECT_EQ(column.first, 2U) << format;
    EXPECT_EQ(column.outline_level, 7U) << format;
  }
}

TEST(FormulonCApiArgumentRange, CellXfIndexPastTheTableIsRejectedBeforeAnyWrite) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3U);
  std::uint32_t count = 0U;
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &count), 0);
  ASSERT_EQ(count, 3U);

  EXPECT_EQ(fm_cell_set_xf_index(wb.handle, 0, 1U, 1U, count), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, 0U, 0U, 2U, 2U, count), kInvalidArgument);
  EXPECT_EQ(CellCount(wb), 0U);

  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 1U, 1U, count - 1U), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, 3U, 0U, 4U, 0U, count - 1U), 0) << fm_last_error_message();
  WorkbookGuard reloaded;
  SaveAndReload(wb, FM_WORKBOOK_FORMAT_XLSX, &reloaded);
  const std::pair<std::uint32_t, std::uint32_t> kStyled[] = {{1U, 1U}, {3U, 0U}, {4U, 0U}};
  for (const auto& [row, col] : kStyled) {
    std::uint32_t xf = 0U;
    ASSERT_EQ(fm_cell_get_xf_index(reloaded.handle, 0, row, col, &xf), 0);
    EXPECT_EQ(xf, count - 1U) << row << "," << col;
  }
}

TEST(FormulonCApiArgumentRange, PivotDataFieldIndicesAreCheckedWithAnExplicitName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0U;
  std::size_t pivot = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot), 0) << fm_last_error_message();
  std::size_t field_count = 0U;
  ASSERT_EQ(fm_workbook_pivot_field_count(wb.handle, 0, pivot, &field_count), 0);
  ASSERT_EQ(field_count, 2U);

  fm_pivot_data_field_spec_t spec{};
  spec.name = "Explicit";
  spec.field_index = 1U;
  spec.aggregation = FM_PIVOT_AGG_SUM;
  spec.show_as = FM_PIVOT_SHOW_AS_NORMAL;
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
    EXPECT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot, rejected, &idx), kInvalidArgument);
    EXPECT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot, 0U, rejected), kInvalidArgument);
  }
  std::size_t data_fields = 0U;
  ASSERT_EQ(fm_workbook_pivot_data_field_count(wb.handle, 0, pivot, &data_fields), 0);
  EXPECT_EQ(data_fields, 1U);
  fm_pivot_data_field_spec_t read{};
  ASSERT_EQ(fm_workbook_pivot_data_field_at(wb.handle, 0, pivot, 0U, &read), 0);
  EXPECT_STREQ(read.name, "Sum of Amount");

  spec.show_as_base_field = 0;
  std::size_t added = 99U;
  EXPECT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot, &spec, &added), 0) << fm_last_error_message();
  EXPECT_EQ(added, 1U);
}

TEST(FormulonCApiArgumentRange, TheOnlyVisibleSheetCannotBeHidden) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Hidden"), 0);
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 1, FM_SHEET_HIDDEN), 0);

  EXPECT_EQ(fm_sheet_set_tab_hidden(wb.handle, 0, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_HIDDEN), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VERY_HIDDEN), kInvalidArgument);
  EXPECT_EQ(View(wb, 0).visibility, FM_SHEET_VISIBLE);
  EXPECT_EQ(ActiveSheet(wb), 0U);

  // Re-stating a hidden sheet's state leaves the visible one alone, so it passes.
  EXPECT_EQ(fm_sheet_set_visibility(wb.handle, 1, FM_SHEET_VERY_HIDDEN), 0);
  EXPECT_EQ(fm_sheet_set_tab_hidden(wb.handle, 1, 1), 0);
}

TEST(FormulonCApiArgumentRange, TheOnlyVisibleSheetCannotBeRemoved) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Hidden"), 0);
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 1, FM_SHEET_HIDDEN), 0);

  EXPECT_EQ(fm_workbook_remove_sheet(wb.handle, 0), kInvalidArgument);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 2U);
  EXPECT_EQ(View(wb, 0).visibility, FM_SHEET_VISIBLE);
  EXPECT_EQ(ActiveSheet(wb), 0U);

  // A hidden sheet can still go, which leaves the visible one in place.
  EXPECT_EQ(fm_workbook_remove_sheet(wb.handle, 1), 0) << fm_last_error_message();
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 1U);
}

TEST(FormulonCApiArgumentRange, CountHelpersClearThePreviousError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 99U, 1U, 1U), kInvalidArgument);
  ASSERT_STRNE(fm_last_error_message(), "");

  EXPECT_EQ(fm_workbook_sheet_count(nullptr), 0U);
  EXPECT_STREQ(fm_last_error_message(), "");
  EXPECT_STREQ(fm_last_error_context(), "");

  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 99U, 1U, 1U), kInvalidArgument);
  EXPECT_EQ(fm_pagination_page_count(nullptr), 0U);
  EXPECT_STREQ(fm_last_error_message(), "");
}
