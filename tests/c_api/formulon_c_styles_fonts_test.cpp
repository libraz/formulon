// Stable C ABI font tests.

#include "formulon_c_styles_test_helpers.h"

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
  // A new workbook seeds the default font at index 0, so the caller's
  // first font lands at index 1.
  EXPECT_EQ(before, 1U);
  EXPECT_EQ(after, 2U);
  EXPECT_EQ(idx, 1U);
  // Round-trip: the freshly added font should read back equal.
  fm_font_record out{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &out), 0);
  EXPECT_STREQ(out.name, "Arial");
  EXPECT_DOUBLE_EQ(out.size, 12.0);
  EXPECT_EQ(out.color_argb, 0xFF112233U);
  EXPECT_EQ(out.bold, 1);
}

TEST(FormulonCApiStyles, GetFontRoundTripIsIdentityForSuperscript) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record superscript = MakeArial();
  superscript.vert_align = 1;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, superscript, &idx), 0);

  fm_font_record loaded{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &loaded), 0);
  EXPECT_STREQ(loaded.name, "Arial");
  EXPECT_EQ(loaded.vert_align, 1U);

  uint32_t count_before = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_before), 0);
  uint32_t reindex = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, loaded, &reindex), 0);
  EXPECT_EQ(reindex, idx);
  uint32_t count_after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count_before);
}

TEST(FormulonCApiStyles, GetEditAddFontPreservesTheUneditedFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record superscript = MakeArial();
  superscript.vert_align = 1;
  superscript.has_family = 1;
  superscript.family = 2;
  superscript.has_charset = 1;
  superscript.charset = 128;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, superscript, &idx), 0);

  fm_font_record edited{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &edited), 0);
  edited.color_argb = 0xFF00FF00U;
  uint32_t recolored = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, edited, &recolored), 0);
  EXPECT_NE(recolored, idx);

  fm_font_record reread{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, recolored, &reread), 0);
  EXPECT_EQ(reread.color_argb, 0xFF00FF00U);
  EXPECT_EQ(reread.vert_align, 1U);
  EXPECT_EQ(reread.has_family, 1);
  EXPECT_EQ(reread.family, 2U);
  EXPECT_EQ(reread.has_charset, 1);
  EXPECT_EQ(reread.charset, 128U);
}

TEST(FormulonCApiStyles, AddFontIsIdentityAgainstAFileLoadedTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count), 0);
  ASSERT_EQ(count, 3U);
  for (uint32_t i = 0; i < count; ++i) {
    fm_font_record loaded{};
    ASSERT_EQ(fm_styles_get_font(wb.handle, i, &loaded), 0) << "font " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_font(wb.handle, loaded, &reindex), 0) << "font " << i;
    EXPECT_EQ(reindex, i);
  }
  uint32_t count_after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count);
}

TEST(FormulonCApiStyles, AddFontDoesNotAliasAThemeColourOntoLiteralRgb) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  // Font 0 is `<color theme="1"/>`. Its sibling ARGB is only a compatibility
  // fallback, so folding it together with a caller's literal RGB record
  // would lose the selector and change the template's body text semantics.
  fm_font_record themed{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, 0U, &themed), 0);
  ASSERT_EQ(themed.color.kind, static_cast<uint8_t>(kFmColorTheme));

  fm_font_record literal_rgb = themed;
  literal_rgb.color = fm_color_spec{};
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, literal_rgb, &index), 0);
  EXPECT_NE(index, 0U);

  // The theme record is untouched and still serialises as a theme colour.
  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<color theme=\"1\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<color theme=\"0\" tint=\"-0.25\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, AddFontDistinguishesAnExplicitOffToggleFromAnAbsentOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record absent = MakeArial();
  absent.bold = 0;
  absent.has_bold = 0;
  fm_font_record explicit_off = absent;
  explicit_off.has_bold = 1;

  uint32_t absent_index = 0;
  uint32_t explicit_index = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, absent, &absent_index), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, explicit_off, &explicit_index), 0);
  EXPECT_NE(absent_index, explicit_index);
}
