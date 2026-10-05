// Stable C ABI cell XF index and mutation tests.

#include "formulon_c_styles_test_helpers.h"

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

TEST(FormulonCApiStyles, GetCellXfRejectsOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  // A fresh workbook seeds one default xf, so index 0 resolves and index
  // 1 is the first out-of-range one.
  EXPECT_EQ(fm_styles_get_cell_xf(wb.handle, 0, &xf), 0);
  EXPECT_NE(fm_styles_get_cell_xf(wb.handle, 1, &xf), 0);
}

TEST(FormulonCApiStyles, CellXfGetterDiagnosticsUseInvokedApi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  auto expect_prefix = [](fm_status_t status, const char* prefix) {
    EXPECT_NE(status, 0);
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(prefix, 0), 0U) << message;
  };

  // Index 1 is past both seeded single-entry tables, so each getter takes
  // its out-of-range path and must name itself in the diagnostic.
  fm_cell_xf cell_xf{};
  expect_prefix(fm_styles_get_cell_xf(wb.handle, 1, &cell_xf), "fm_styles_get_cell_xf:");
  expect_prefix(fm_styles_get_cell_style_xf(wb.handle, 1, &cell_xf), "fm_styles_get_cell_style_xf:");
}

TEST(FormulonCApiStyles, CellXfAdderDiagnosticsUseInvokedApi) {
  auto expect_prefix = [](fm_status_t status, const char* prefix) {
    EXPECT_NE(status, 0);
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(prefix, 0), 0U) << message;
  };

  uint32_t index = 0;
  fm_cell_xf cell_xf{};
  expect_prefix(fm_styles_add_cell_xf(nullptr, cell_xf, &index), "fm_styles_add_cell_xf:");

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  cell_xf.font_index = 1U;
  expect_prefix(fm_styles_add_cell_xf(wb.handle, cell_xf, &index), "fm_styles_add_cell_xf:");

  cell_xf = fm_cell_xf{};
  cell_xf.xf_id = 1U;
  expect_prefix(fm_styles_add_cell_xf(wb.handle, cell_xf, &index), "fm_styles_add_cell_xf:");

  cell_xf = fm_cell_xf{};
  cell_xf.has_text_rotation = 1;
  cell_xf.text_rotation = 181U;
  expect_prefix(fm_styles_add_cell_xf(wb.handle, cell_xf, &index), "fm_styles_add_cell_xf:");
}

TEST(FormulonCApiStyles, SetThenSaveLoadPreservesXfIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Register the xf records the stamp below names. A `<c s="7">` against
  // a shorter `<cellXfs>` resolves to no style, so the reader falls back
  // to the default rather than handing back an index whose record
  // `fm_styles_get_cell_xf` would then refuse to return.
  wb.handle->workbook().mutable_styles().cell_xfs.resize(8);

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

  xf.has_wrap_text = 1;

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
  // guaranteed to land at a fresh index instead of deduping to it. The
  // presence flag is what makes the value part of the record.
  xf.wrap_text = 1;
  xf.has_wrap_text = 1;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &after), 0);
  // Index 0 is the seeded default xf.
  EXPECT_EQ(before, 1U);
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

  // A fresh workbook seeds Excel's reserved slots, so the caller's first
  // record lands after them: one default font, one empty border, and the
  // two fills (`none`, `gray125`) Excel always writes first.
  EXPECT_EQ(font_idx, 1U);
  EXPECT_EQ(fill_idx, 2U);
  EXPECT_EQ(border_idx, 1U);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it. The
  // presence flag is what makes the value part of the record.
  xf.wrap_text = 1;
  xf.has_wrap_text = 1;

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

namespace {

// A record that sets every apply flag, quotePrefix and an explicit
// `<protection>` child whose values differ from the schema defaults.
fm_cell_xf MakeFlaggedXf() {
  fm_cell_xf xf{};
  xf.apply_number_format = 1;
  xf.apply_font = 1;
  xf.apply_fill = 1;
  xf.apply_border = 1;
  xf.apply_alignment = 1;
  xf.apply_protection = 1;
  xf.quote_prefix = 1;
  xf.has_protection = 1;
  xf.locked = 0;
  xf.hidden = 1;
  return xf;
}

void ExpectFlagged(const fm_cell_xf& xf) {
  EXPECT_EQ(xf.apply_number_format, 1);
  EXPECT_EQ(xf.apply_font, 1);
  EXPECT_EQ(xf.apply_fill, 1);
  EXPECT_EQ(xf.apply_border, 1);
  EXPECT_EQ(xf.apply_alignment, 1);
  EXPECT_EQ(xf.apply_protection, 1);
  EXPECT_EQ(xf.quote_prefix, 1);
  EXPECT_EQ(xf.has_protection, 1);
  EXPECT_EQ(xf.locked, 0);
  EXPECT_EQ(xf.hidden, 1);
}

}  // namespace

TEST(FormulonCApiStyles, XfApplyFlagsRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t xf_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, MakeFlaggedXf(), &xf_index), 0);
  ASSERT_NE(xf_index, 0U);
  uint32_t style_xf_id = 0;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, MakeFlaggedXf(), &style_xf_id), 0);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, xf_index, &got), 0);
  ExpectFlagged(got);

  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_index), 0);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  EXPECT_NE(styles_xml.find("applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" "
                            "applyAlignment=\"1\" applyProtection=\"1\" quotePrefix=\"1\""),
            std::string::npos)
      << styles_xml;
  EXPECT_NE(styles_xml.find("<protection locked=\"0\" hidden=\"1\"/>"), std::string::npos) << styles_xml;

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  uint32_t reloaded_index = 0;
  ASSERT_EQ(fm_cell_get_xf_index(reloaded.handle, 0, 0, 0, &reloaded_index), 0);
  fm_cell_xf after{};
  ASSERT_EQ(fm_styles_get_cell_xf(reloaded.handle, reloaded_index, &after), 0);
  ExpectFlagged(after);
  fm_cell_xf style_after{};
  ASSERT_EQ(fm_styles_get_cell_style_xf(reloaded.handle, style_xf_id, &style_after), 0);
  ExpectFlagged(style_after);
}

TEST(FormulonCApiStyles, XfProtectionDistinguishesDedup) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf locked{};
  locked.has_protection = 1;
  locked.locked = 1;
  fm_cell_xf unlocked = locked;
  unlocked.locked = 0;
  uint32_t a = 0;
  uint32_t b = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, locked, &a), 0);
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, unlocked, &b), 0);
  EXPECT_NE(a, b);

  // Without the child, `locked` / `hidden` are ignored: the record
  // deduplicates with the default XF and reads back the schema defaults.
  fm_cell_xf absent{};
  absent.locked = 0;
  absent.hidden = 1;
  uint32_t c = 99;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, absent, &c), 0);
  EXPECT_EQ(c, 0U);
  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, c, &got), 0);
  EXPECT_EQ(got.has_protection, 0);
  EXPECT_EQ(got.locked, 1);
  EXPECT_EQ(got.hidden, 0);
}
