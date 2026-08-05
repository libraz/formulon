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
#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "io/styles_reader.h"
#include "workbook.h"

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

TEST(FormulonCApiStyles, CustomNumFmtOverridingBuiltinIdWins) {
  // A file may define a custom <numFmt> whose numFmtId collides with a
  // built-in slot; Excel honours the file's definition. Inject such an
  // override and confirm the getter returns the custom string, not the
  // built-in.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  formulon::io::StylesTable& styles = wb.handle->workbook().mutable_styles();
  formulon::io::NumFmtRecord rec;
  rec.id = 14;  // Built-in id whose default is "mm-dd-yy".
  rec.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("yyyy\"年\"m\"月\"d\"日\"");
  styles.num_fmts.push_back(rec);

  const char* s = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "yyyy\"年\"m\"月\"d\"日\"");
}

TEST(FormulonCApiStyles, UnknownNumFmtIdRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  // Reserved built-in slot (id 5) without a custom override surfaces
  // an error rather than an empty string.
  EXPECT_NE(fm_styles_get_num_fmt_string(wb.handle, 5, &s), 0);
}

TEST(FormulonCApiStyles, DifferentialFormatGetterExposesCfDxfTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  formulon::io::StylesTable& styles = wb.handle->workbook().mutable_styles();
  formulon::io::DifferentialFormat dxf;
  dxf.has_font = true;
  dxf.font.bold = true;
  dxf.font.color_argb = 0xFFFF0000U;
  dxf.has_fill = true;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFFFFFF00U;
  dxf.has_border = true;
  dxf.border.left.style = 1;
  dxf.border.left.color_argb = 0xFF000000U;
  dxf.has_num_fmt = true;
  dxf.num_fmt_id = 164;
  dxf.num_fmt_code = "0.00";
  styles.dxfs.push_back(dxf);

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_dxf_record out{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, 0, &out), 0);
  EXPECT_EQ(out.font_engaged, 1);
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.font.color_argb, 0xFFFF0000U);
  EXPECT_EQ(out.fill_engaged, 1);
  EXPECT_EQ(out.fill.pattern, 1U);
  EXPECT_EQ(out.fill.fg_argb, 0xFFFFFF00U);
  EXPECT_EQ(out.border_engaged, 1);
  EXPECT_EQ(out.border.left.style, 1U);
  EXPECT_EQ(out.border.left.color_argb, 0xFF000000U);
  EXPECT_EQ(out.num_fmt_engaged, 1);
  EXPECT_EQ(out.num_fmt_id, 164U);
  ASSERT_NE(out.num_fmt_code, nullptr);
  EXPECT_STREQ(out.num_fmt_code, "0.00");

  EXPECT_NE(fm_styles_get_dxf(wb.handle, 1, &out), 0);
}

TEST(FormulonCApiStyles, AddDxfDedupsAndReadsBack) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_dxf_record dxf{};
  dxf.font_engaged = 1;
  dxf.font.name = "Arial";
  dxf.font.size = 12.0;
  dxf.font.bold = 1;
  dxf.font.color_argb = 0xFFFF0000U;
  dxf.fill_engaged = 1;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFFFFFF00U;
  dxf.num_fmt_engaged = 1;
  dxf.num_fmt_id = 164;
  dxf.num_fmt_code = "0.00";

  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &a), 0);
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &b), 0);
  EXPECT_EQ(a, b);

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_dxf_record out{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, a, &out), 0);
  EXPECT_EQ(out.font_engaged, 1);
  EXPECT_STREQ(out.font.name, "Arial");
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.font.color_argb, 0xFFFF0000U);
  EXPECT_EQ(out.fill_engaged, 1);
  EXPECT_EQ(out.fill.pattern, 1U);
  EXPECT_EQ(out.fill.fg_argb, 0xFFFFFF00U);
  EXPECT_EQ(out.num_fmt_engaged, 1);
  EXPECT_EQ(out.num_fmt_id, 164U);
  ASSERT_NE(out.num_fmt_code, nullptr);
  EXPECT_STREQ(out.num_fmt_code, "0.00");
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

// ---- Add-side dedup primitives -------------------------------------

namespace {

fm_font_record MakeArial() {
  fm_font_record r{};
  r.name = "Arial";
  r.size = 12.0;
  r.color_argb = 0xFF112233U;
  r.bold = 1;
  r.italic = 0;
  r.strike = 0;
  r.underline = 0;
  return r;
}

fm_fill_record MakeRedFill() {
  fm_fill_record r{};
  r.pattern = 1;  // solid
  r.fg_argb = 0xFFFF0000U;
  r.bg_argb = 0xFF000000U;
  return r;
}

fm_border_record MakeThinBoxBorder() {
  fm_border_record r{};
  r.left.style = 1;  // thin
  r.left.color_argb = 0xFF000000U;
  r.right = r.left;
  r.top = r.left;
  r.bottom = r.left;
  r.diagonal.style = 0;
  r.diagonal.color_argb = 0;
  r.diagonal_up = 0;
  r.diagonal_down = 0;
  return r;
}

}  // namespace

TEST(FormulonCApiStyles, AddFontDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &a), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddFontDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_font_record arial = MakeArial();
  fm_font_record calibri = MakeArial();
  calibri.name = "Calibri";
  ASSERT_EQ(fm_styles_add_font(wb.handle, arial, &a), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, calibri, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddFontGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  // Round-trip: the freshly added font should read back equal.
  fm_font_record out{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &out), 0);
  EXPECT_STREQ(out.name, "Arial");
  EXPECT_DOUBLE_EQ(out.size, 12.0);
  EXPECT_EQ(out.color_argb, 0xFF112233U);
  EXPECT_EQ(out.bold, 1);
}

TEST(FormulonCApiStyles, AddFontExPreservesVerticalAlignmentWithoutChangingLegacyRecord) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record_ex superscript{};
  superscript.base = MakeArial();
  superscript.vert_align = 1;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font_ex(wb.handle, superscript, &idx), 0);

  fm_font_record_ex loaded{};
  ASSERT_EQ(fm_styles_get_font_ex(wb.handle, idx, &loaded), 0);
  EXPECT_STREQ(loaded.base.name, "Arial");
  EXPECT_EQ(loaded.vert_align, 1U);

  fm_font_record legacy{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &legacy), 0);
  EXPECT_STREQ(legacy.name, "Arial");
  EXPECT_EQ(legacy.underline, 0U);
}

TEST(FormulonCApiStyles, AddFillDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &a), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddFillDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_fill_record red = MakeRedFill();
  fm_fill_record blue = MakeRedFill();
  blue.fg_argb = 0xFF0000FFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, red, &a), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, blue, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddFillGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  fm_fill_record out{};
  ASSERT_EQ(fm_styles_get_fill(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.pattern, 1U);
  EXPECT_EQ(out.fg_argb, 0xFFFF0000U);
}

TEST(FormulonCApiStyles, AddBorderDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &a), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddBorderDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_border_record thin = MakeThinBoxBorder();
  fm_border_record dashed = MakeThinBoxBorder();
  dashed.left.style = 3;  // dashed
  ASSERT_EQ(fm_styles_add_border(wb.handle, thin, &a), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, dashed, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddBorderGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  fm_border_record out{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.left.style, 1U);
  EXPECT_EQ(out.right.style, 1U);
  EXPECT_EQ(out.diagonal_up, 0);
  EXPECT_EQ(out.diagonal_down, 0);
}

TEST(FormulonCApiStyles, AddNumFmtBuiltinReturnsBuiltinId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "General", &id), 0);
  EXPECT_EQ(id, 0U);
  // Adding a built-in must not create a custom entry.
  uint32_t font_count = 7;  // unrelated, just ensure other tables untouched
  EXPECT_EQ(fm_styles_get_font_count(wb.handle, &font_count), 0);
  EXPECT_EQ(font_count, 0U);
}

TEST(FormulonCApiStyles, AddNumFmtCustomReturnsCustomId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t id = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &id), 0);
  EXPECT_GE(id, 164U);
  // Resolves through the read-side getter.
  const char* s = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, id, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "\"USD\" #,##0");
}

TEST(FormulonCApiStyles, AddNumFmtCustomDedup) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t a = 0;
  uint16_t b = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &a), 0);
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddCellXfDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;  // built-in General
  xf.horizontal_align = 1;
  xf.vertical_align = 2;
  xf.wrap_text = 1;

  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &a), 0);
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddCellXfDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  xf.horizontal_align = 1;
  xf.vertical_align = 2;
  xf.wrap_text = 1;

  uint32_t a = 0;
  uint32_t b = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &a), 0);
  // Flip wrap_text — distinct record.
  xf.wrap_text = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddCellXfGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &before), 0);
  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it.
  xf.wrap_text = 1;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 2U);
  EXPECT_EQ(idx, 1U);
}

TEST(FormulonCApiStyles, AddCellXfOnFreshWorkbookKeepsZeroAsDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  // A fresh workbook's tables start empty, so the first font/fill/border
  // added legitimately becomes index 0 — there is no anonymous placeholder
  // seeded ahead of it.
  EXPECT_EQ(font_idx, 0U);
  EXPECT_EQ(fill_idx, 0U);
  EXPECT_EQ(border_idx, 0U);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it.
  xf.wrap_text = 1;

  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);
  EXPECT_NE(xf_idx, 0U);

  fm_cell_xf default_xf{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, 0, &default_xf), 0);
  EXPECT_EQ(default_xf.font_index, 0U);
  EXPECT_EQ(default_xf.fill_index, 0U);
  EXPECT_EQ(default_xf.border_index, 0U);
  EXPECT_EQ(default_xf.num_fmt_id, 0U);
  EXPECT_EQ(default_xf.wrap_text, 0);
}

TEST(FormulonCApiStyles, AddCellXfRejectsOutOfRangeIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Empty styles table: any non-zero index is out of range.
  fm_cell_xf xf{};
  xf.font_index = 5;  // intentionally OOR
  xf.fill_index = 0;
  xf.border_index = 0;
  xf.num_fmt_id = 0;
  uint32_t idx = 0xFFFFFFFFU;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
}

TEST(FormulonCApiStyles, AddCellXfRejectsUnknownNumFmt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 200;  // not a registered custom and not a documented built-in
  uint32_t idx = 0xFFFFFFFFU;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
}

TEST(FormulonCApiStyles, AddNullArgumentRejected) {
  uint32_t idx = 0;
  EXPECT_NE(fm_styles_add_font(nullptr, MakeArial(), &idx), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_styles_add_font(wb.handle, MakeArial(), nullptr), 0);
  uint16_t id = 0;
  EXPECT_NE(fm_styles_add_num_fmt(nullptr, "General", &id), 0);
  EXPECT_NE(fm_styles_add_num_fmt(wb.handle, "General", nullptr), 0);
  EXPECT_NE(fm_styles_add_dxf(nullptr, fm_dxf_record{}, &idx), 0);
  EXPECT_NE(fm_styles_add_dxf(wb.handle, fm_dxf_record{}, nullptr), 0);
}

TEST(FormulonCApiStyles, FullLifecycleSurvivesSaveLoad) {
  // Build font -> xf -> stamp on cell -> save -> reload -> the font is
  // resolvable via the read-side accessors.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);

  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.5), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  uint32_t reread_xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 0, 0, &reread_xf), 0);
  fm_cell_xf reloaded{};
  EXPECT_EQ(fm_styles_get_cell_xf(wb2.handle, reread_xf, &reloaded), 0);
  fm_font_record reloaded_font{};
  ASSERT_EQ(fm_styles_get_font(wb2.handle, reloaded.font_index, &reloaded_font), 0);
  EXPECT_STREQ(reloaded_font.name, "Arial");
  EXPECT_DOUBLE_EQ(reloaded_font.size, 12.0);
  EXPECT_EQ(reloaded_font.bold, 1);
}

// Regression coverage for the externally-reported symptom fixed by removing
// the spurious `ensure_default_style_roots` seeding from the font/fill/
// border adders: a non-default font + fill + custom numFmt combined into an
// `<xf>` must produce an `<xf>` index that is distinct from (and does not
// dedup to) the all-zero placeholder xf, and that index must survive
// `setCellXfIndex` -> save -> reload with the font, fill, and numFmt content
// intact on the reloaded workbook. Note that a fresh workbook's font/fill
// tables legitimately start at index 0 (that is the behavior the fix
// restored), so this test asserts on round-tripped *content*, not on the
// font/fill indices being non-zero. The WASM (`src/wasm/parts/
// workbook_styles.cpp`) and Node addon (`src/node_addon/parts/styles.cc`)
// bindings marshal the identical field set into `fm_cell_xf` and call this
// same `fm_styles_add_cell_xf` entry point, so this test also exercises
// their shared code path end to end.
TEST(FormulonCApiStyles, AddXfWithNonDefaultFontFillNumFmtRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t font_idx = 0xFFFFFFFFU;
  uint32_t fill_idx = 0xFFFFFFFFU;
  uint16_t num_fmt_id = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "0.00%", &num_fmt_id), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = 0;
  xf.num_fmt_id = num_fmt_id;
  uint32_t xf_idx = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);
  ASSERT_NE(xf_idx, 0U) << "non-default font/fill/numFmt must not dedup to the "
                           "all-zero placeholder xf at index 0";

  // Stability: re-adding the identical record must return the same index.
  uint32_t xf_idx_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx_again), 0);
  EXPECT_EQ(xf_idx, xf_idx_again);

  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 1, 0.4225), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 1, 1, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);

  uint32_t reread_xf = 0;
  ASSERT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 1, 1, &reread_xf), 0);
  ASSERT_EQ(reread_xf, xf_idx);

  fm_cell_xf reloaded{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb2.handle, reread_xf, &reloaded), 0);
  EXPECT_EQ(reloaded.font_index, font_idx);
  EXPECT_EQ(reloaded.fill_index, fill_idx);
  EXPECT_EQ(reloaded.num_fmt_id, num_fmt_id);

  fm_font_record reloaded_font{};
  ASSERT_EQ(fm_styles_get_font(wb2.handle, reloaded.font_index, &reloaded_font), 0);
  EXPECT_STREQ(reloaded_font.name, "Arial");
  EXPECT_DOUBLE_EQ(reloaded_font.size, 12.0);
  EXPECT_EQ(reloaded_font.bold, 1);

  fm_fill_record reloaded_fill{};
  ASSERT_EQ(fm_styles_get_fill(wb2.handle, reloaded.fill_index, &reloaded_fill), 0);
  EXPECT_EQ(reloaded_fill.pattern, 1U);
  EXPECT_EQ(reloaded_fill.fg_argb, 0xFFFF0000U);

  const char* reloaded_fmt = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb2.handle, reloaded.num_fmt_id, &reloaded_fmt), 0);
  ASSERT_NE(reloaded_fmt, nullptr);
  EXPECT_STREQ(reloaded_fmt, "0.00%");
}
