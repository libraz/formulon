// Stable C ABI fill and border tests.

#include "formulon_c_record_arg.h"
#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, AddFillAndBorderAreIdentityAgainstAFileLoadedTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  std::size_t fills = 0;
  ASSERT_EQ(fm_workbook_get_fill_count(wb.handle, &fills), 0);
  ASSERT_EQ(fills, 2U);
  for (uint32_t i = 0; i < fills; ++i) {
    fm_fill_t loaded{};
    ASSERT_EQ(fm_workbook_get_fill_at(wb.handle, i, &loaded), 0) << "fill " << i;
    std::size_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_workbook_add_fill(wb.handle, &loaded, &reindex), 0) << "fill " << i;
    EXPECT_EQ(reindex, i);
  }

  std::size_t borders = 0;
  ASSERT_EQ(fm_workbook_get_border_count(wb.handle, &borders), 0);
  ASSERT_EQ(borders, 2U);
  for (uint32_t i = 0; i < borders; ++i) {
    fm_border_t loaded{};
    ASSERT_EQ(fm_workbook_get_border_at(wb.handle, i, &loaded), 0) << "border " << i;
    std::size_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_workbook_add_border(wb.handle, &loaded, &reindex), 0) << "border " << i;
    EXPECT_EQ(reindex, i);
  }

  std::size_t fills_after = 0;
  std::size_t borders_after = 0;
  ASSERT_EQ(fm_workbook_get_fill_count(wb.handle, &fills_after), 0);
  ASSERT_EQ(fm_workbook_get_border_count(wb.handle, &borders_after), 0);
  EXPECT_EQ(fills_after, fills);
  EXPECT_EQ(borders_after, borders);
}

TEST(FormulonCApiStyles, AddFillAndBorderDistinguishColourSpecifications) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  fm_fill_t themed_fill{};
  ASSERT_EQ(fm_workbook_get_fill_at(wb.handle, 1U, &themed_fill), 0);
  ASSERT_EQ(themed_fill.foreground.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));
  fm_fill_t literal_fill = themed_fill;
  literal_fill.foreground = RgbColor(0xFF112233U);
  std::size_t fill_index = 0;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &literal_fill, &fill_index), 0);
  EXPECT_NE(fill_index, 1U);

  fm_border_t themed_border{};
  ASSERT_EQ(fm_workbook_get_border_at(wb.handle, 1U, &themed_border), 0);
  ASSERT_EQ(themed_border.left.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));
  fm_border_t literal_border = themed_border;
  literal_border.left.color = RgbColor(0xFF112233U);
  std::size_t border_index = 0;
  ASSERT_EQ(fm_workbook_add_border(wb.handle, &literal_border, &border_index), 0);
  EXPECT_NE(border_index, 1U);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<fgColor theme=\"4\" tint=\"0.5\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<bgColor indexed=\"64\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, SelectorColoursRemainAuthoritativeAcrossGetAddAndSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_color_t theme{};
  theme.kind = static_cast<uint8_t>(FM_COLOR_KIND_THEME);
  theme.theme = 3U;
  theme.tint = 0.5;
  fm_color_t indexed{};
  indexed.kind = static_cast<uint8_t>(FM_COLOR_KIND_INDEXED);
  indexed.indexed = 9U;
  fm_color_t automatic{};
  automatic.kind = static_cast<uint8_t>(FM_COLOR_KIND_AUTO);

  fm_font_t font = MakeArial();
  font.color = theme;
  std::size_t font_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &font, &font_index), 0);
  fm_font_t got_font{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, font_index, &got_font), 0);
  EXPECT_EQ(got_font.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));
  EXPECT_EQ(got_font.color.theme, 3U);
  EXPECT_DOUBLE_EQ(got_font.color.tint, 0.5);
  std::size_t font_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &got_font, &font_again), 0);
  EXPECT_EQ(font_again, font_index);

  fm_fill_t fill = MakeRedFill();
  fill.foreground = indexed;
  fill.background = automatic;
  std::size_t fill_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &fill, &fill_index), 0);
  fm_fill_t got_fill{};
  ASSERT_EQ(fm_workbook_get_fill_at(wb.handle, fill_index, &got_fill), 0);
  EXPECT_EQ(got_fill.foreground.kind, static_cast<uint8_t>(FM_COLOR_KIND_INDEXED));
  EXPECT_EQ(got_fill.foreground.indexed, 9U);
  EXPECT_EQ(got_fill.background.kind, static_cast<uint8_t>(FM_COLOR_KIND_AUTO));
  std::size_t fill_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &got_fill, &fill_again), 0);
  EXPECT_EQ(fill_again, fill_index);

  fm_border_t border = MakeThinBoxBorder();
  border.left.color = theme;
  border.right.color = indexed;
  std::size_t border_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_border(wb.handle, &border, &border_index), 0);
  fm_border_t got_border{};
  ASSERT_EQ(fm_workbook_get_border_at(wb.handle, border_index, &got_border), 0);
  EXPECT_EQ(got_border.left.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));
  EXPECT_EQ(got_border.left.color.theme, 3U);
  EXPECT_EQ(got_border.right.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_INDEXED));
  EXPECT_EQ(got_border.right.color.indexed, 9U);
  std::size_t border_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_border(wb.handle, &got_border, &border_again), 0);
  EXPECT_EQ(border_again, border_index);

  fm_dxf_t dxf{};
  dxf.font_engaged = 1;
  dxf.font = font;
  dxf.font.color = automatic;
  dxf.fill_engaged = 1;
  dxf.fill = fill;
  dxf.border_engaged = 1;
  dxf.border = border;
  std::uint32_t dxf_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &dxf_index), 0);
  fm_dxf_t got_dxf{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, dxf_index, &got_dxf), 0);
  EXPECT_EQ(got_dxf.font.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_AUTO));
  EXPECT_EQ(got_dxf.fill.foreground.kind, static_cast<uint8_t>(FM_COLOR_KIND_INDEXED));
  EXPECT_EQ(got_dxf.border.left.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));
  std::uint32_t dxf_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &got_dxf, &dxf_again), 0);
  EXPECT_EQ(dxf_again, dxf_index);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);

  fm_font_t reloaded_font{};
  ASSERT_EQ(fm_workbook_get_font_at(reloaded.handle, font_index, &reloaded_font), 0);
  EXPECT_EQ(reloaded_font.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));
  EXPECT_EQ(reloaded_font.color.theme, 3U);
  EXPECT_DOUBLE_EQ(reloaded_font.color.tint, 0.5);
  fm_fill_t reloaded_fill{};
  ASSERT_EQ(fm_workbook_get_fill_at(reloaded.handle, fill_index, &reloaded_fill), 0);
  EXPECT_EQ(reloaded_fill.foreground.kind, static_cast<uint8_t>(FM_COLOR_KIND_INDEXED));
  EXPECT_EQ(reloaded_fill.foreground.indexed, 9U);
  EXPECT_EQ(reloaded_fill.background.kind, static_cast<uint8_t>(FM_COLOR_KIND_AUTO));
  fm_border_t reloaded_border{};
  ASSERT_EQ(fm_workbook_get_border_at(reloaded.handle, border_index, &reloaded_border), 0);
  EXPECT_EQ(reloaded_border.left.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));
  fm_dxf_t reloaded_dxf{};
  ASSERT_EQ(fm_workbook_get_dxf_at(reloaded.handle, dxf_index, &reloaded_dxf), 0);
  EXPECT_EQ(reloaded_dxf.font.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_AUTO));
  EXPECT_EQ(reloaded_dxf.border.left.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));
}

TEST(FormulonCApiStyles, AddFillDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t a = 0xFFFFFFFFU;
  std::size_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &a), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddFillDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t a = 0;
  std::size_t b = 0;
  fm_fill_t red = MakeRedFill();
  fm_fill_t blue = MakeRedFill();
  blue.foreground = RgbColor(0xFF0000FFU);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &red, &a), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &blue, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddFillGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t before = 0;
  ASSERT_EQ(fm_workbook_get_fill_count(wb.handle, &before), 0);
  std::size_t idx = 0;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &idx), 0);
  std::size_t after = 0;
  ASSERT_EQ(fm_workbook_get_fill_count(wb.handle, &after), 0);
  // The seeded table reserves `none` and `gray125` at 0 and 1, matching
  // what Excel puts in every workbook, so the caller's fill starts at 2.
  EXPECT_EQ(before, 2U);
  EXPECT_EQ(after, 3U);
  EXPECT_EQ(idx, 2U);
  fm_fill_t out{};
  ASSERT_EQ(fm_workbook_get_fill_at(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.pattern, 1);
  EXPECT_EQ(out.foreground.argb, 0xFFFF0000U);
}

TEST(FormulonCApiStyles, AddBorderDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t a = 0xFFFFFFFFU;
  std::size_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &a), 0);
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddBorderDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t a = 0;
  std::size_t b = 0;
  fm_border_t thin = MakeThinBoxBorder();
  fm_border_t dashed = MakeThinBoxBorder();
  dashed.left.style = 3;  // dashed
  ASSERT_EQ(fm_workbook_add_border(wb.handle, &thin, &a), 0);
  ASSERT_EQ(fm_workbook_add_border(wb.handle, &dashed, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddBorderGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t before = 0;
  ASSERT_EQ(fm_workbook_get_border_count(wb.handle, &before), 0);
  std::size_t idx = 0;
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &idx), 0);
  std::size_t after = 0;
  ASSERT_EQ(fm_workbook_get_border_count(wb.handle, &after), 0);
  // Index 0 is the seeded empty border.
  EXPECT_EQ(before, 1U);
  EXPECT_EQ(after, 2U);
  EXPECT_EQ(idx, 1U);
  fm_border_t out{};
  ASSERT_EQ(fm_workbook_get_border_at(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.left.style, 1);
  EXPECT_EQ(out.right.style, 1);
  EXPECT_EQ(out.diagonal_up, 0);
  EXPECT_EQ(out.diagonal_down, 0);
}

TEST(FormulonCApiStyles, FontFillAndBorderOrdinalsOutsideTheirEnumerationsAreRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t fonts_before = 0;
  std::size_t fills_before = 0;
  std::size_t borders_before = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &fonts_before), 0);
  ASSERT_EQ(fm_workbook_get_fill_count(wb.handle, &fills_before), 0);
  ASSERT_EQ(fm_workbook_get_border_count(wb.handle, &borders_before), 0);
  std::size_t index = 0;

  for (int32_t underline : {-1, FM_UNDERLINE_STYLE_DOUBLE_ACCOUNTING + 1, 256}) {
    fm_font_t font{};
    font.size_pt = 11.0;
    font.underline = underline;
    EXPECT_EQ(fm_workbook_add_font(wb.handle, &font, &index), FM_STATUS_INVALID_ARGUMENT) << underline;
  }
  for (int32_t pattern : {-1, FM_FILL_PATTERN_GRAY0625 + 1, 256}) {
    fm_fill_t fill{};
    fill.pattern = pattern;
    EXPECT_EQ(fm_workbook_add_fill(wb.handle, &fill, &index), FM_STATUS_INVALID_ARGUMENT) << pattern;
  }
  for (int32_t style : {-1, FM_BORDER_STYLE_SLANT_DASH_DOT + 1, 256}) {
    fm_border_t border{};
    border.diagonal.style = style;
    EXPECT_EQ(fm_workbook_add_border(wb.handle, &border, &index), FM_STATUS_INVALID_ARGUMENT) << style;
  }
  fm_dxf_t dxf{};
  dxf.border_engaged = 1;
  dxf.border.left.style = 256;
  uint32_t dxf_index = 0;
  EXPECT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &dxf_index), FM_STATUS_INVALID_ARGUMENT);

  std::size_t fonts_after = 0;
  std::size_t fills_after = 0;
  std::size_t borders_after = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &fonts_after), 0);
  ASSERT_EQ(fm_workbook_get_fill_count(wb.handle, &fills_after), 0);
  ASSERT_EQ(fm_workbook_get_border_count(wb.handle, &borders_after), 0);
  EXPECT_EQ(fonts_after, fonts_before);
  EXPECT_EQ(fills_after, fills_before);
  EXPECT_EQ(borders_after, borders_before);
}
