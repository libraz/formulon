// Stable C ABI style-table lifetime and default-slot tests.

#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, CreateEmptySeedsReservedStyleSlots) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet1"), 0);

  uint32_t font_count = 0;
  uint32_t fill_count = 0;
  uint32_t border_count = 0;
  uint32_t cell_xf_count = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &font_count), 0);
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &fill_count), 0);
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &border_count), 0);
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &cell_xf_count), 0);
  EXPECT_EQ(font_count, 1U);
  EXPECT_EQ(fill_count, 2U);
  EXPECT_EQ(border_count, 1U);
  EXPECT_EQ(cell_xf_count, 1U);

  // Excel reserves fill 0 for `none` and fill 1 for `gray125`, so a caller's
  // own fill must never take either slot; the same holds for font 0 and
  // border 0. A non-zero index is what proves the reserved slots survived.
  fm_font_record font{};
  font.name = "Arial";
  font.size = 12.0;
  uint32_t font_index = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, font, &font_index), 0);
  EXPECT_GT(font_index, 0U);

  fm_fill_record fill{};
  fill.pattern = 1U; /* solid */
  fill.fg_argb = 0xFFFF0000U;
  uint32_t fill_index = 0;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, fill, &fill_index), 0);
  EXPECT_GT(fill_index, 1U);

  fm_border_record border{};
  border.left.style = 1U; /* thin */
  uint32_t border_index = 0;
  ASSERT_EQ(fm_styles_add_border(wb.handle, border, &border_index), 0);
  EXPECT_GT(border_index, 0U);

  fm_cell_xf xf{};
  xf.font_index = font_index;
  xf.fill_index = fill_index;
  xf.border_index = border_index;
  uint32_t xf_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_index), 0);
  EXPECT_GT(xf_index, 0U);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  EXPECT_NE(styles_xml.find("gray125"), std::string::npos);
  EXPECT_NE(styles_xml.find("name=\"Normal\""), std::string::npos);
  // Cell xf 0 stays the unformatted default: it must not pick up the
  // caller's fill.
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  const std::size_t first_xf = styles_xml.find("<xf", cell_xfs_begin);
  ASSERT_NE(first_xf, std::string::npos);
  const std::size_t first_xf_end = styles_xml.find('>', first_xf);
  ASSERT_NE(first_xf_end, std::string::npos);
  EXPECT_NE(styles_xml.substr(first_xf, first_xf_end - first_xf).find("fillId=\"0\""), std::string::npos);
}

TEST(FormulonCApiStyles, DefaultFontDeclaresWhatAnUnstyledCellSavesAs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "unstyled"), 0);

  // The seeded default is Excel's Calibri 11, and `add_font` can only append
  // beside it -- which is exactly why the declaring entry point exists.
  fm_font_record seeded{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, 0, &seeded), 0);
  EXPECT_STREQ(seeded.name, "Calibri");
  fm_font_record appended{};
  appended.name = "游ゴシック";
  appended.size = 11.0;
  uint32_t appended_index = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, appended, &appended_index), 0);
  EXPECT_GT(appended_index, 0U);

  fm_font_record declared{};
  declared.name = "游ゴシック";
  declared.size = 11.0;
  declared.has_family = 1;
  declared.family = 2;
  declared.has_charset = 1;
  declared.charset = 128;
  declared.color_argb = 0xFF000000U;
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, declared), 0);

  fm_font_record read_back{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, 0, &read_back), 0);
  EXPECT_STREQ(read_back.name, "游ゴシック");
  EXPECT_EQ(read_back.charset, 128U);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t fonts_begin = styles_xml.find("<fonts");
  ASSERT_NE(fonts_begin, std::string::npos);
  const std::size_t first_font = styles_xml.find("<font>", fonts_begin);
  ASSERT_NE(first_font, std::string::npos);
  const std::size_t first_font_end = styles_xml.find("</font>", first_font);
  ASSERT_NE(first_font_end, std::string::npos);
  const std::string first = styles_xml.substr(first_font, first_font_end - first_font);
  EXPECT_NE(first.find("<name val=\"游ゴシック\"/>"), std::string::npos);
  EXPECT_EQ(first.find("Calibri"), std::string::npos);
}

TEST(FormulonCApiStyles, SetFontOverwritesInPlaceAndRejectsAnAbsentIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record added{};
  added.name = "Meiryo";
  added.size = 12.0;
  uint32_t index = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, added, &index), 0);
  ASSERT_GT(index, 0U);

  fm_font_record replacement{};
  replacement.name = "MS Gothic";
  replacement.size = 9.0;
  ASSERT_EQ(fm_styles_set_font(wb.handle, index, replacement), 0);

  fm_font_record read_back{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, index, &read_back), 0);
  EXPECT_STREQ(read_back.name, "MS Gothic");
  EXPECT_DOUBLE_EQ(read_back.size, 9.0);

  // Overwriting does not grow the table, and an index past its end is
  // refused rather than filled in with default records.
  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count), 0);
  EXPECT_EQ(count, index + 1U);
  EXPECT_NE(fm_styles_set_font(wb.handle, count, replacement), 0);
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count), 0);
  EXPECT_EQ(count, index + 1U);

  EXPECT_NE(fm_styles_set_font(nullptr, 0, replacement), 0);
  EXPECT_NE(fm_workbook_set_default_font(nullptr, replacement), 0);
}

TEST(FormulonCApiStyles, FontSchemeSurvivesAReadModifyWriteThroughTheAbi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record body{};
  body.name = "游ゴシック";
  body.size = 11.0;
  body.has_charset = 1;
  body.charset = 128;
  body.scheme = 2;  // minor: the theme's body font
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, body), 0);

  // The link has to come back out of the projection, or a host that reads a
  // font, edits one field and writes it back silently unlinks it.
  fm_font_record read_back{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, 0, &read_back), 0);
  EXPECT_EQ(read_back.scheme, 2U);
  read_back.size = 12.0;
  ASSERT_EQ(fm_styles_set_font(wb.handle, 0, read_back), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  EXPECT_NE(styles_xml.find("<charset val=\"128\"/><scheme val=\"minor\"/>"), std::string::npos);

  // A record left at the default writes no element, so `add_font` sees a
  // distinct entry rather than folding the two together.
  fm_font_record unlinked{};
  unlinked.name = "游ゴシック";
  unlinked.size = 12.0;
  unlinked.has_charset = 1;
  unlinked.charset = 128;
  uint32_t unlinked_index = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, unlinked, &unlinked_index), 0);
  EXPECT_GT(unlinked_index, 0U);
}
