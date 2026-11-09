//
// C ABI tests for the projected PivotTable report handle (fm_pivot_report_*),
// the one-cell pivot anchor, and the style a new pivot starts with.

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"
#include "formulon_c_pivot_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kIndexOutOfRange = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange);

/// Sets the PivotTable style default, or clears it for "".
fm_status_t SetPivotStyleDefault(fm_workbook_t* wb, const char* pivot_style) {
  fm_table_style_defaults_patch_t patch = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  patch.pivot_style_op = pivot_style[0] == '\0' ? FM_PATCH_OP_CLEAR : FM_PATCH_OP_SET;
  patch.pivot_style = pivot_style;
  return fm_workbook_update_table_style_defaults(wb, &patch);
}

TEST(FormulonCApiPivotReport, BoundsCoverTheProjectedCells) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot_index), 0) << fm_get_last_error_message();

  PivotCellsGuard report;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &report.handle), 0) << fm_get_last_error_message();
  ASSERT_NE(report.handle, nullptr);

  std::size_t count = 0;
  ASSERT_EQ(fm_pivot_report_get_cell_count(report.handle, &count), 0);
  ASSERT_GT(count, 0U);

  int32_t present = 0;
  fm_cell_rect_t rect{};
  ASSERT_EQ(fm_pivot_report_get_bounds(report.handle, &present, &rect), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(rect.first_row, 0U);
  EXPECT_EQ(rect.first_col, 3U) << "the scratch pivot is anchored at D1";

  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_info_t cell = FM_STRUCT_INIT(fm_pivot_cell_info_t);
    ASSERT_EQ(fm_pivot_report_get_cell_at(report.handle, i, &cell), 0);
    EXPECT_GE(cell.row, rect.first_row);
    EXPECT_LE(cell.row, rect.last_row);
    EXPECT_GE(cell.col, rect.first_col);
    EXPECT_LE(cell.col, rect.last_col);
  }
}

TEST(FormulonCApiPivotReport, ReadsRefuseNullAndOutOfRangeArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot_index), 0) << fm_get_last_error_message();
  PivotCellsGuard report;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &report.handle), 0);

  std::size_t count = 0;
  int32_t present = 0;
  fm_cell_rect_t rect{};
  fm_pivot_cell_info_t cell = FM_STRUCT_INIT(fm_pivot_cell_info_t);
  EXPECT_EQ(fm_pivot_report_get_cell_count(nullptr, &count), kBindingNullPointer);
  EXPECT_EQ(fm_pivot_report_get_cell_count(report.handle, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_pivot_report_get_bounds(nullptr, &present, &rect), kBindingNullPointer);
  EXPECT_EQ(fm_pivot_report_get_bounds(report.handle, nullptr, &rect), kBindingNullPointer);
  EXPECT_EQ(fm_pivot_report_get_bounds(report.handle, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_pivot_report_get_cell_at(nullptr, 0, &cell), kBindingNullPointer);
  EXPECT_EQ(fm_pivot_report_get_cell_at(report.handle, 0, nullptr), kBindingNullPointer);

  ASSERT_EQ(fm_pivot_report_get_cell_count(report.handle, &count), 0);
  EXPECT_EQ(fm_pivot_report_get_cell_at(report.handle, count, &cell), kIndexOutOfRange);
  fm_pivot_report_destroy(nullptr);  // a NULL handle is a no-op
}

TEST(FormulonCApiPivotReport, AnchorMovesTheReportToOneCell) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot_index), 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_sheet_set_pivot_anchor(wb.handle, 0, pivot_index, 12U, 6U), 0) << fm_get_last_error_message();
  PivotCellsGuard report;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &report.handle), 0) << fm_get_last_error_message();
  int32_t present = 0;
  fm_cell_rect_t rect{};
  ASSERT_EQ(fm_pivot_report_get_bounds(report.handle, &present, &rect), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(rect.first_row, 12U);
  EXPECT_EQ(rect.first_col, 6U);
}

TEST(FormulonCApiPivotReport, AnchorRefusesBadIndices) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_pivot_anchor(wb.handle, 9, pivot_index, 0U, 0U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
  EXPECT_EQ(fm_sheet_set_pivot_anchor(nullptr, 0, pivot_index, 0U, 0U), kBindingNullPointer);
}

TEST(FormulonCApiPivotReport, NewPivotTakesThePivotStyleDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetPivotStyleDefault(wb.handle, "PivotStyleMedium5"), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot_index), 0) << fm_get_last_error_message();
  fm_pivot_style_options_t style = FM_STRUCT_INIT(fm_pivot_style_options_t);
  ASSERT_EQ(fm_sheet_get_pivot_style_options(wb.handle, 0, pivot_index, &style), 0);
  EXPECT_EQ(style.present, 1);
  EXPECT_STREQ(style.style_name, "PivotStyleMedium5");
  EXPECT_EQ(style.show_row_headers, 1);
  EXPECT_EQ(style.show_column_headers, 1);
  EXPECT_EQ(style.show_last_column, 1);
  EXPECT_EQ(style.show_row_stripes, 0);
  EXPECT_EQ(style.show_column_stripes, 0);
}

TEST(FormulonCApiPivotReport, NewPivotFallsBackToTheBuiltInStyleWhenTheDefaultIsRemoved) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetPivotStyleDefault(wb.handle, "PivotStyleMedium5"), 0);
  ASSERT_EQ(SetPivotStyleDefault(wb.handle, ""), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, &pivot_index), 0) << fm_get_last_error_message();
  fm_pivot_style_options_t style = FM_STRUCT_INIT(fm_pivot_style_options_t);
  ASSERT_EQ(fm_sheet_get_pivot_style_options(wb.handle, 0, pivot_index, &style), 0);
  EXPECT_STREQ(style.style_name, "PivotStyleLight16");
}

TEST(FormulonCApiPivotReport, FieldSpecSubtotalTopDefaultsToTheGroupHead) {
  const fm_tristate_t values[] = {FM_TRISTATE_DEFAULT, FM_TRISTATE_FALSE, FM_TRISTATE_TRUE};
  const int32_t expected_read[] = {FM_TRISTATE_TRUE, FM_TRISTATE_FALSE, FM_TRISTATE_TRUE};
  for (std::size_t i = 0; i < 3U; ++i) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    std::uint32_t cache_id = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), 0);
    std::size_t cache_field = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "F", &cache_field), 0);
    std::size_t pivot_index = 99;
    ASSERT_EQ(AddPivot(wb.handle, 0, "PT", cache_id, 0U, 0U, &pivot_index), 0);

    fm_pivot_field_spec_t spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
    spec.source_name = "F";
    spec.custom_name = "";
    spec.axis = FM_PIVOT_AXIS_ROW;
    spec.subtotal_top = values[i];
    spec.number_format = "";
    std::size_t field = 99;
    ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_index, &spec, &field), 0) << fm_get_last_error_message();
    fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
    ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, field, &info), 0);
    EXPECT_EQ(info.subtotal_top, expected_read[i]) << "spec value " << values[i];
  }
}

TEST(FormulonCApiPivotReport, FieldSpecSubtotalTopOutsideTheTristateIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), 0);
  std::size_t cache_field = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "F", &cache_field), 0);
  std::size_t pivot_index = 99;
  ASSERT_EQ(AddPivot(wb.handle, 0, "PT", cache_id, 0U, 0U, &pivot_index), 0);

  fm_pivot_field_spec_t bad = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  bad.source_name = "F";
  bad.custom_name = "";
  bad.axis = FM_PIVOT_AXIS_ROW;
  bad.subtotal_top = 3;
  bad.number_format = "";
  std::size_t field = 99;
  EXPECT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_index, &bad, &field),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

}  // namespace
