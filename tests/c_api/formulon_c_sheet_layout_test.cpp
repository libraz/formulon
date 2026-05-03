// Copyright 2026 libraz. Licensed under the MIT License.
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
