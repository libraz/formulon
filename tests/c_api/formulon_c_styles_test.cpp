// Copyright 2026 libraz. Licensed under the MIT License.
//
// Stable C ABI tests for the styles surface
// (`fm_cell_get_xf_index`, `fm_cell_set_xf_index`,
// `fm_styles_get_cell_xf`, `fm_styles_get_font`,
// `fm_styles_get_num_fmt_string`).
//
// The test driver is C++ for gtest convenience but everything it
// touches across the boundary is the pure-C surface.

#include <cstdint>
#include <cstring>
#include <string>
#include <vector>

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

TEST(FormulonCApiStyles, CellXfIndexDefaultsToZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t xf = 99;
  EXPECT_EQ(fm_cell_get_xf_index(wb.handle, 0, 0, 0, &xf), 0);
  EXPECT_EQ(xf, 0U);
}

TEST(FormulonCApiStyles, CellXfIndexRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_cell_set_xf_index(wb.handle, 0, 5, 7, 42), 0);
  uint32_t xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb.handle, 0, 5, 7, &xf), 0);
  EXPECT_EQ(xf, 42U);
}

TEST(FormulonCApiStyles, SetXfIndexRejectsBadSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // sheet 99 does not exist.
  EXPECT_NE(fm_cell_set_xf_index(wb.handle, 99, 0, 0, 1), 0);
}

TEST(FormulonCApiStyles, NullArgumentRejected) {
  uint32_t xf = 0;
  EXPECT_NE(fm_cell_get_xf_index(nullptr, 0, 0, 0, &xf), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_cell_get_xf_index(wb.handle, 0, 0, 0, nullptr), 0);
}

TEST(FormulonCApiStyles, BuiltinNumFmtResolves) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  EXPECT_EQ(fm_styles_get_num_fmt_string(wb.handle, 0, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "General");
  EXPECT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "mm-dd-yy");
}

TEST(FormulonCApiStyles, UnknownNumFmtIdRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  // Reserved built-in slot (id 5) without a custom override surfaces
  // an error rather than an empty string.
  EXPECT_NE(fm_styles_get_num_fmt_string(wb.handle, 5, &s), 0);
}

TEST(FormulonCApiStyles, GetCellXfRejectsOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  // A fresh workbook has an empty styles table, so any xf_index is
  // out of range. Calling with index 0 still fails because the table
  // has zero entries.
  EXPECT_NE(fm_styles_get_cell_xf(wb.handle, 0, &xf), 0);
}

TEST(FormulonCApiStyles, SetThenSaveLoadPreservesXfIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Stamp xf_index = 7 on cell A1 and save.
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 3.14), 0);
  EXPECT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, 7), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  uint32_t xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 0, 0, &xf), 0);
  EXPECT_EQ(xf, 7U);
}
