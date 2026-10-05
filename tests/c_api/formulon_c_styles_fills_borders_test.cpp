// Stable C ABI fill and border tests.

#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, AddFillAndBorderAreIdentityAgainstAFileLoadedTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  uint32_t fills = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &fills), 0);
  ASSERT_EQ(fills, 2U);
  for (uint32_t i = 0; i < fills; ++i) {
    fm_fill_record loaded{};
    ASSERT_EQ(fm_styles_get_fill(wb.handle, i, &loaded), 0) << "fill " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_fill(wb.handle, loaded, &reindex), 0) << "fill " << i;
    EXPECT_EQ(reindex, i);
  }

  uint32_t borders = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &borders), 0);
  ASSERT_EQ(borders, 2U);
  for (uint32_t i = 0; i < borders; ++i) {
    fm_border_record loaded{};
    ASSERT_EQ(fm_styles_get_border(wb.handle, i, &loaded), 0) << "border " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_border(wb.handle, loaded, &reindex), 0) << "border " << i;
    EXPECT_EQ(reindex, i);
  }

  uint32_t fills_after = 0;
  uint32_t borders_after = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &fills_after), 0);
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &borders_after), 0);
  EXPECT_EQ(fills_after, fills);
  EXPECT_EQ(borders_after, borders);
}

TEST(FormulonCApiStyles, AddFillAndBorderDistinguishColourSpecifications) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  fm_fill_record themed_fill{};
  ASSERT_EQ(fm_styles_get_fill(wb.handle, 1U, &themed_fill), 0);
  ASSERT_EQ(themed_fill.fg.kind, static_cast<uint8_t>(kFmColorTheme));
  fm_fill_record literal_fill = themed_fill;
  literal_fill.fg = fm_color_spec{};
  literal_fill.fg_argb = themed_fill.fg_argb;
  uint32_t fill_index = 0;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, literal_fill, &fill_index), 0);
  EXPECT_NE(fill_index, 1U);

  fm_border_record themed_border{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, 1U, &themed_border), 0);
  ASSERT_EQ(themed_border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  fm_border_record literal_border = themed_border;
  literal_border.left.color = fm_color_spec{};
  uint32_t border_index = 0;
  ASSERT_EQ(fm_styles_add_border(wb.handle, literal_border, &border_index), 0);
  EXPECT_NE(border_index, 1U);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<fgColor theme=\"4\" tint=\"0.5\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<bgColor indexed=\"64\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, SelectorColoursRemainAuthoritativeAcrossGetAddAndSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_color_spec theme{};
  theme.kind = static_cast<uint8_t>(kFmColorTheme);
  theme.theme = 3U;
  theme.tint = 0.5;
  fm_color_spec indexed{};
  indexed.kind = static_cast<uint8_t>(kFmColorIndexed);
  indexed.indexed = 9U;
  fm_color_spec automatic{};
  automatic.kind = static_cast<uint8_t>(kFmColorAuto);

  fm_font_record font = MakeArial();
  font.color_argb = 0x01020304U;  // compatibility fallback, not a render result
  font.color = theme;
  uint32_t font_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, font, &font_index), 0);
  fm_font_record got_font{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, font_index, &got_font), 0);
  EXPECT_EQ(got_font.color.kind, static_cast<uint8_t>(kFmColorTheme));
  EXPECT_EQ(got_font.color.theme, 3U);
  EXPECT_DOUBLE_EQ(got_font.color.tint, 0.5);
  EXPECT_EQ(got_font.color_argb, 0x01020304U);
  uint32_t font_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, got_font, &font_again), 0);
  EXPECT_EQ(font_again, font_index);

  fm_fill_record fill = MakeRedFill();
  fill.fg_argb = 0x05060708U;
  fill.bg_argb = 0x090A0B0CU;
  fill.fg = indexed;
  fill.bg = automatic;
  uint32_t fill_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, fill, &fill_index), 0);
  fm_fill_record got_fill{};
  ASSERT_EQ(fm_styles_get_fill(wb.handle, fill_index, &got_fill), 0);
  EXPECT_EQ(got_fill.fg.kind, static_cast<uint8_t>(kFmColorIndexed));
  EXPECT_EQ(got_fill.fg.indexed, 9U);
  EXPECT_EQ(got_fill.bg.kind, static_cast<uint8_t>(kFmColorAuto));
  uint32_t fill_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, got_fill, &fill_again), 0);
  EXPECT_EQ(fill_again, fill_index);

  fm_border_record border = MakeThinBoxBorder();
  border.left.color = theme;
  border.right.color = indexed;
  uint32_t border_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_border(wb.handle, border, &border_index), 0);
  fm_border_record got_border{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, border_index, &got_border), 0);
  EXPECT_EQ(got_border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  EXPECT_EQ(got_border.left.color.theme, 3U);
  EXPECT_EQ(got_border.right.color.kind, static_cast<uint8_t>(kFmColorIndexed));
  EXPECT_EQ(got_border.right.color.indexed, 9U);
  uint32_t border_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_border(wb.handle, got_border, &border_again), 0);
  EXPECT_EQ(border_again, border_index);

  fm_dxf_record dxf{};
  dxf.font_engaged = 1;
  dxf.font = font;
  dxf.font.color = automatic;
  dxf.fill_engaged = 1;
  dxf.fill = fill;
  dxf.border_engaged = 1;
  dxf.border = border;
  uint32_t dxf_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &dxf_index), 0);
  fm_dxf_record got_dxf{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, dxf_index, &got_dxf), 0);
  EXPECT_EQ(got_dxf.font.color.kind, static_cast<uint8_t>(kFmColorAuto));
  EXPECT_EQ(got_dxf.fill.fg.kind, static_cast<uint8_t>(kFmColorIndexed));
  EXPECT_EQ(got_dxf.border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  uint32_t dxf_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, got_dxf, &dxf_again), 0);
  EXPECT_EQ(dxf_again, dxf_index);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);

  fm_font_record reloaded_font{};
  ASSERT_EQ(fm_styles_get_font(reloaded.handle, font_index, &reloaded_font), 0);
  EXPECT_EQ(reloaded_font.color.kind, static_cast<uint8_t>(kFmColorTheme));
  EXPECT_EQ(reloaded_font.color.theme, 3U);
  EXPECT_DOUBLE_EQ(reloaded_font.color.tint, 0.5);
  fm_fill_record reloaded_fill{};
  ASSERT_EQ(fm_styles_get_fill(reloaded.handle, fill_index, &reloaded_fill), 0);
  EXPECT_EQ(reloaded_fill.fg.kind, static_cast<uint8_t>(kFmColorIndexed));
  EXPECT_EQ(reloaded_fill.fg.indexed, 9U);
  EXPECT_EQ(reloaded_fill.bg.kind, static_cast<uint8_t>(kFmColorAuto));
  fm_border_record reloaded_border{};
  ASSERT_EQ(fm_styles_get_border(reloaded.handle, border_index, &reloaded_border), 0);
  EXPECT_EQ(reloaded_border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  fm_dxf_record reloaded_dxf{};
  ASSERT_EQ(fm_styles_get_dxf(reloaded.handle, dxf_index, &reloaded_dxf), 0);
  EXPECT_EQ(reloaded_dxf.font.color.kind, static_cast<uint8_t>(kFmColorAuto));
  EXPECT_EQ(reloaded_dxf.border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
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
  // The seeded table reserves `none` and `gray125` at 0 and 1, matching
  // what Excel puts in every workbook, so the caller's fill starts at 2.
  EXPECT_EQ(before, 2U);
  EXPECT_EQ(after, 3U);
  EXPECT_EQ(idx, 2U);
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
  // Index 0 is the seeded empty border.
  EXPECT_EQ(before, 1U);
  EXPECT_EQ(after, 2U);
  EXPECT_EQ(idx, 1U);
  fm_border_record out{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.left.style, 1U);
  EXPECT_EQ(out.right.style, 1U);
  EXPECT_EQ(out.diagonal_up, 0);
  EXPECT_EQ(out.diagonal_down, 0);
}
