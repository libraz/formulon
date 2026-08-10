//
// Stable C ABI smoke tests for the sheet view / layout entry points
// added in `formulon_c.h`. The driver is C++ for gtest convenience but
// only touches the pure-C surface.

#include <cstdint>
#include <cstring>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

}  // namespace

TEST(FormulonCApiSheetLayout, GetViewDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_sheet_view_t v{};
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.zoom_scale, 100U);
  EXPECT_EQ(v.freeze_rows, 0U);
  EXPECT_EQ(v.freeze_cols, 0U);
  EXPECT_EQ(v.tab_hidden, 0);
}

TEST(FormulonCApiSheetLayout, SetViewSetters) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_zoom(wb.handle, 0, 175U), 0);
  ASSERT_EQ(fm_sheet_set_freeze(wb.handle, 0, 4U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_tab_hidden(wb.handle, 0, 1), 0);
  fm_sheet_view_t v{};
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.zoom_scale, 175U);
  EXPECT_EQ(v.freeze_rows, 4U);
  EXPECT_EQ(v.freeze_cols, 2U);
  EXPECT_EQ(v.tab_hidden, 1);
}

TEST(FormulonCApiSheetLayout, GetViewExDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_sheet_view_ex_t v{};
  ASSERT_EQ(fm_sheet_get_view_ex(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.zoom_scale, 100U);
  EXPECT_EQ(v.freeze_rows, 0U);
  EXPECT_EQ(v.freeze_cols, 0U);
  EXPECT_EQ(v.tab_hidden, 0);
  EXPECT_EQ(v.show_grid_lines, 1);
  EXPECT_EQ(v.show_row_col_headers, 1);
  EXPECT_EQ(v.show_zeros, 1);
  EXPECT_EQ(v.right_to_left, 0);
  EXPECT_EQ(v.tab_selected, 0);
  ASSERT_NE(v.view_mode, nullptr);
  EXPECT_STREQ(v.view_mode, "");
}

TEST(FormulonCApiSheetLayout, SetViewExSetters) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_show_grid_lines(wb.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_set_show_row_col_headers(wb.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_set_show_zeros(wb.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_set_right_to_left(wb.handle, 0, 1), 0);
  ASSERT_EQ(fm_sheet_set_tab_selected(wb.handle, 0, 1), 0);
  ASSERT_EQ(fm_sheet_set_view_mode(wb.handle, 0, "pageBreakPreview"), 0);
  fm_sheet_view_ex_t v{};
  ASSERT_EQ(fm_sheet_get_view_ex(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.show_grid_lines, 0);
  EXPECT_EQ(v.show_row_col_headers, 0);
  EXPECT_EQ(v.show_zeros, 0);
  EXPECT_EQ(v.right_to_left, 1);
  EXPECT_EQ(v.tab_selected, 1);
  ASSERT_NE(v.view_mode, nullptr);
  EXPECT_STREQ(v.view_mode, "pageBreakPreview");
}

TEST(FormulonCApiSheetLayout, AutoFilterXmlRoundTripsAndClears) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const char* xml =
      "<autoFilter ref=\"A1:C10\"><filterColumn colId=\"1\"><filters><filter val=\"east\"/>"
      "</filters></filterColumn></autoFilter>";
  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, xml), 0);

  const char* actual = nullptr;
  ASSERT_EQ(fm_sheet_get_auto_filter_xml(wb.handle, 0, &actual), 0);
  ASSERT_NE(actual, nullptr);
  EXPECT_STREQ(actual, xml);

  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, ""), 0);
  ASSERT_EQ(fm_sheet_get_auto_filter_xml(wb.handle, 0, &actual), 0);
  ASSERT_NE(actual, nullptr);
  EXPECT_STREQ(actual, "");
}

TEST(FormulonCApiSheetLayout, AutoFilterXmlSurvivesXlsxSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* xml =
      "<autoFilter ref=\"A1:C10\"><filterColumn colId=\"1\"><filters><filter val=\"east\"/>"
      "</filters></filterColumn></autoFilter>";
  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, xml), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  const char* actual = nullptr;
  ASSERT_EQ(fm_sheet_get_auto_filter_xml(reloaded.handle, 0, &actual), 0);
  ASSERT_NE(actual, nullptr);
  EXPECT_STREQ(actual, xml);
}

TEST(FormulonCApiSheetLayout, AutoFilterXmlRejectsWrongRoot) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<filterColumn colId=\"0\"/>"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, nullptr), 0);
  // A longer element name that merely starts with "autoFilter".
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilterColumn colId=\"0\"/>"), 0);
  // An unterminated element would break the whole worksheet part.
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\">"), 0);
  EXPECT_NE(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter"), 0);
  EXPECT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\"/>"), 0);
}

TEST(FormulonCApiSheetLayout, SetViewModeRejectsNullPointer) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_sheet_set_view_mode(wb.handle, 0, nullptr), 0);
}

TEST(FormulonCApiSheetLayout, ZoomClampedToBounds) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_zoom(wb.handle, 0, 5U), 0);
  fm_sheet_view_t v{};
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.zoom_scale, 10U);
  ASSERT_EQ(fm_sheet_set_zoom(wb.handle, 0, 9999U), 0);
  ASSERT_EQ(fm_sheet_get_view(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.zoom_scale, 400U);
}

TEST(FormulonCApiSheetLayout, ColumnWidthAndOutline) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 0U, 4U, 18.5), 0);
  ASSERT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 6U, 8U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 10U, 10U, 1), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_column_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 3U);

  fm_column_layout_t a{};
  fm_column_layout_t b{};
  fm_column_layout_t c{};
  ASSERT_EQ(fm_sheet_get_column(wb.handle, 0, 0, &a), 0);
  ASSERT_EQ(fm_sheet_get_column(wb.handle, 0, 1, &b), 0);
  ASSERT_EQ(fm_sheet_get_column(wb.handle, 0, 2, &c), 0);
  EXPECT_EQ(a.first, 0U);
  EXPECT_EQ(a.last, 4U);
  EXPECT_DOUBLE_EQ(a.width, 18.5);
  EXPECT_EQ(b.first, 6U);
  EXPECT_EQ(b.last, 8U);
  EXPECT_EQ(b.outline_level, 2U);
  EXPECT_EQ(c.first, 10U);
  EXPECT_EQ(c.last, 10U);
  EXPECT_EQ(c.hidden, 1);
}

TEST(FormulonCApiSheetLayout, RowOverridesUpsert) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 4U, 22.0), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 9U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 4U, 3U), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_row_override_count(wb.handle, 0, &count), 0);
  // Two distinct rows even though row 4 received two updates: the C
  // ABI's row setters upsert by row index.
  EXPECT_EQ(count, 2U);

  fm_row_layout_t row4{};
  fm_row_layout_t row9{};
  // Iteration order is implementation-defined (insertion order for
  // upsert on missing row); inspect both entries by content.
  fm_row_layout_t entry0{};
  fm_row_layout_t entry1{};
  ASSERT_EQ(fm_sheet_get_row_override(wb.handle, 0, 0, &entry0), 0);
  ASSERT_EQ(fm_sheet_get_row_override(wb.handle, 0, 1, &entry1), 0);
  if (entry0.row == 4U) {
    row4 = entry0;
    row9 = entry1;
  } else {
    row4 = entry1;
    row9 = entry0;
  }
  EXPECT_EQ(row4.row, 4U);
  EXPECT_DOUBLE_EQ(row4.height, 22.0);
  EXPECT_EQ(row4.outline_level, 3U);
  EXPECT_EQ(row4.hidden, 0);
  EXPECT_EQ(row9.row, 9U);
  EXPECT_EQ(row9.hidden, 1);
}

TEST(FormulonCApiSheetLayout, InvalidSheetIndexRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_sheet_view_t v{};
  EXPECT_NE(fm_sheet_get_view(wb.handle, 99U, &v), 0);
  EXPECT_NE(fm_sheet_set_zoom(wb.handle, 99U, 100U), 0);
  EXPECT_NE(fm_sheet_set_freeze(wb.handle, 99U, 1U, 1U), 0);
  EXPECT_NE(fm_sheet_set_tab_hidden(wb.handle, 99U, 1), 0);
  EXPECT_NE(fm_sheet_set_column_width(wb.handle, 99U, 0U, 1U, 10.0), 0);
}

TEST(FormulonCApiSheetLayout, ColumnSetterRejectsInverseSpan) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_sheet_set_column_width(wb.handle, 0, 5U, 3U, 10.0), 0);
  EXPECT_NE(fm_sheet_set_column_hidden(wb.handle, 0, 5U, 3U, 1), 0);
  EXPECT_NE(fm_sheet_set_column_outline(wb.handle, 0, 5U, 3U, 2U), 0);
}
