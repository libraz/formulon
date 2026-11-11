// Stable C ABI differential-format tests.

#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, DifferentialFormatGetterExposesCfDxfTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  formulon::StylesTable& styles = wb.handle->workbook().mutable_styles();
  formulon::DifferentialFormat dxf;
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

  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_dxf_t out{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, 0, &out), 0);
  EXPECT_EQ(out.has_font, 1);
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.font.color.argb, 0xFFFF0000U);
  EXPECT_EQ(out.has_fill, 1);
  EXPECT_EQ(out.fill.pattern, 1);
  EXPECT_EQ(out.fill.foreground.argb, 0xFFFFFF00U);
  EXPECT_EQ(out.has_border, 1);
  EXPECT_EQ(out.border.left.style, 1);
  EXPECT_EQ(out.border.left.color.argb, 0xFF000000U);
  ASSERT_NE(out.number_format, nullptr);
  EXPECT_STREQ(out.number_format, "0.00");

  EXPECT_NE(fm_workbook_get_dxf_at(wb.handle, 1, &out), 0);
}

TEST(FormulonCApiStyles, AddDxfDedupsAndReadsBack) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_dxf_t dxf{};
  dxf.has_font = 1;
  dxf.font.name = "Arial";
  dxf.font.size_pt = 12.0;
  dxf.font.bold = 1;
  dxf.font.color = RgbColor(0xFFFF0000U);
  dxf.has_fill = 1;
  dxf.fill.pattern = 1;
  dxf.fill.foreground = RgbColor(0xFFFFFF00U);
  dxf.number_format = "0.00";

  std::uint32_t a = 0xFFFFFFFFU;
  std::uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &a), 0);
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &b), 0);
  EXPECT_EQ(a, b);

  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_dxf_t out{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, a, &out), 0);
  EXPECT_EQ(out.has_font, 1);
  EXPECT_STREQ(out.font.name, "Arial");
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.font.color.argb, 0xFFFF0000U);
  EXPECT_EQ(out.has_fill, 1);
  EXPECT_EQ(out.fill.pattern, 1);
  EXPECT_EQ(out.fill.foreground.argb, 0xFFFFFF00U);
  ASSERT_NE(out.number_format, nullptr);
  EXPECT_STREQ(out.number_format, "0.00");
}

TEST(FormulonCApiStyles, DxfNumberFormatIsNullWhenTheDxfHasNone) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_dxf_t dxf{};
  dxf.has_font = 1;
  dxf.font.bold = 1;
  dxf.font.has_bold = 1;
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &index), 0);
  fm_dxf_t out{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, index, &out), 0);
  EXPECT_EQ(out.number_format, nullptr);
}

TEST(FormulonCApiStyles, DxfNumberFormatSurvivesSaveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_dxf_t dxf{};
  dxf.number_format = "0.000%";
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &index), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  fm_dxf_t out{};
  ASSERT_EQ(fm_workbook_get_dxf_at(reloaded.handle, index, &out), 0);
  ASSERT_NE(out.number_format, nullptr);
  EXPECT_STREQ(out.number_format, "0.000%");
}

TEST(FormulonCApiStyles, DxfFontRoundTripsVerticalAlignmentThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_dxf_t dxf{};
  dxf.has_font = 1;
  dxf.font.name = "Calibri";
  dxf.font.size_pt = 9.0;
  dxf.font.color = RgbColor(0xFF112233U);
  dxf.font.vert_align = 1;  // superscript
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &index), 0);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<vertAlign val=\"superscript\"/>"), std::string::npos);

  fm_dxf_t loaded{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, index, &loaded), 0);
  EXPECT_EQ(loaded.font.vert_align, 1);

  std::size_t count_before = 0;
  ASSERT_EQ(fm_workbook_get_dxf_count(wb.handle, &count_before), 0);
  std::uint32_t reindex = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &loaded, &reindex), 0);
  EXPECT_EQ(reindex, index);
  std::size_t count_after = 0;
  ASSERT_EQ(fm_workbook_get_dxf_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count_before);
}

TEST(FormulonCApiStyles, DxfAlignmentAndProtectionPreserveIdentityThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  std::string alignment_input = "<alignment horizontal=\"center\" wrapText=\"1\"/>";
  std::string protection_input = "<protection locked=\"0\" hidden=\"1\"/>";
  fm_dxf_t alignment{};
  alignment.alignment_xml = alignment_input.c_str();
  fm_dxf_t protection{};
  protection.protection_xml = protection_input.c_str();

  std::uint32_t alignment_index = 0xFFFFFFFFU;
  std::uint32_t protection_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &alignment, &alignment_index), 0);
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &protection, &protection_index), 0);
  EXPECT_NE(alignment_index, protection_index);
  alignment_input = "caller-owned alignment text replaced";
  protection_input = "caller-owned protection text replaced";

  fm_dxf_t got_alignment{};
  fm_dxf_t got_protection{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, alignment_index, &got_alignment), 0);
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, protection_index, &got_protection), 0);
  ASSERT_NE(got_alignment.alignment_xml, nullptr);
  ASSERT_NE(got_protection.protection_xml, nullptr);
  EXPECT_STREQ(got_alignment.alignment_xml, "<alignment horizontal=\"center\" wrapText=\"1\"/>");
  EXPECT_EQ(got_alignment.protection_xml, nullptr);
  EXPECT_EQ(got_protection.alignment_xml, nullptr);
  EXPECT_STREQ(got_protection.protection_xml, "<protection locked=\"0\" hidden=\"1\"/>");

  std::uint32_t alignment_again = 0xFFFFFFFFU;
  std::uint32_t protection_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &got_alignment, &alignment_again), 0);
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &got_protection, &protection_again), 0);
  EXPECT_EQ(alignment_again, alignment_index);
  EXPECT_EQ(protection_again, protection_index);

  const std::string before_save = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(before_save.find("<alignment horizontal=\"center\" wrapText=\"1\"/>"), std::string::npos);
  EXPECT_NE(before_save.find("<protection locked=\"0\" hidden=\"1\"/>"), std::string::npos);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);

  fm_dxf_t reloaded_alignment{};
  fm_dxf_t reloaded_protection{};
  ASSERT_EQ(fm_workbook_get_dxf_at(reloaded.handle, alignment_index, &reloaded_alignment), 0);
  ASSERT_EQ(fm_workbook_get_dxf_at(reloaded.handle, protection_index, &reloaded_protection), 0);
  EXPECT_STREQ(reloaded_alignment.alignment_xml, "<alignment horizontal=\"center\" wrapText=\"1\"/>");
  EXPECT_EQ(reloaded_alignment.protection_xml, nullptr);
  EXPECT_EQ(reloaded_protection.alignment_xml, nullptr);
  EXPECT_STREQ(reloaded_protection.protection_xml, "<protection locked=\"0\" hidden=\"1\"/>");

  std::uint32_t reloaded_alignment_again = 0xFFFFFFFFU;
  std::uint32_t reloaded_protection_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(reloaded.handle, &reloaded_alignment, &reloaded_alignment_again), 0);
  ASSERT_EQ(fm_workbook_add_dxf(reloaded.handle, &reloaded_protection, &reloaded_protection_again), 0);
  EXPECT_EQ(reloaded_alignment_again, alignment_index);
  EXPECT_EQ(reloaded_protection_again, protection_index);

  const std::string after_load = formulon::io::write_styles(reloaded.handle->workbook().styles());
  EXPECT_NE(after_load.find("<alignment horizontal=\"center\" wrapText=\"1\"/>"), std::string::npos);
  EXPECT_NE(after_load.find("<protection locked=\"0\" hidden=\"1\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, DxfFontDistinguishesAnExplicitBoldOffFromAnAbsentToggle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  formulon::DifferentialFormat stored;
  stored.has_font = true;
  stored.font.has_bold = true;  // `<b val="0"/>`: switch bold off
  stored.font.bold = false;
  wb.handle->workbook().mutable_styles().dxfs.push_back(stored);

  // A rule that leaves bold alone must not fold onto the switch-it-off one.
  fm_dxf_t leave_bold_alone{};
  leave_bold_alone.has_font = 1;
  leave_bold_alone.font.name = "";
  leave_bold_alone.font.size_pt = 11.0;
  leave_bold_alone.font.color = RgbColor(0xFF000000U);
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &leave_bold_alone, &index), 0);
  EXPECT_NE(index, 0U);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<b val=\"0\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, DxfFontWithoutSizeRoundTripsAndCellFontKeepsSize) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_dxf_t dxf{};
  dxf.has_font = 1;
  dxf.font.name = "";
  dxf.font.size_pt = 11.0;
  dxf.font.has_bold = 1;
  dxf.font.bold = 1;
  dxf.font.color = RgbColor(0xFF000000U);
  ASSERT_EQ(dxf.font.has_size_pt, 0);
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &index), 0);

  fm_dxf_t loaded{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, index, &loaded), 0);
  EXPECT_EQ(loaded.font.has_size_pt, 0);

  // The dxf is the only font-bearing entry that may omit the size.
  const std::string dxf_xml = formulon::io::write_styles(wb.handle->workbook().styles());
  const std::size_t dxfs_pos = dxf_xml.find("<dxfs");
  ASSERT_NE(dxfs_pos, std::string::npos);
  EXPECT_EQ(dxf_xml.find("<sz", dxfs_pos), std::string::npos);

  // With the flag set the dxf writes its size.
  dxf.font.has_size_pt = 1;
  dxf.font.size_pt = 13.0;
  std::uint32_t sized = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &sized), 0);
  EXPECT_NE(sized, index);
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, sized, &loaded), 0);
  EXPECT_EQ(loaded.font.has_size_pt, 1);

  // A cell font keeps a stated size.
  fm_font_t cell{};
  cell.name = "Arial";
  cell.has_size_pt = 1;
  cell.size_pt = 14.0;
  cell.color = RgbColor(0xFF000000U);
  std::uint32_t font_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &cell, &font_index), 0);
  fm_font_t font_out{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, font_index, &font_out), 0);
  EXPECT_EQ(font_out.has_size_pt, 1);
  EXPECT_DOUBLE_EQ(font_out.size_pt, 14.0);
  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  const std::size_t fonts_end = xml.find("</fonts>");
  ASSERT_NE(fonts_end, std::string::npos);
  EXPECT_NE(xml.rfind("<sz val=\"14\"/>", fonts_end), std::string::npos);
}

TEST(FormulonCApiStyles, SetDxfOverwritesTheSlotInPlace) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_dxf_t dxf{};
  dxf.has_fill = 1;
  dxf.fill.pattern = 1;
  dxf.fill.foreground = RgbColor(0xFFFFFF00U);
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &index), 0);

  fm_dxf_t replacement{};
  replacement.has_font = 1;
  replacement.font.bold = 1;
  replacement.number_format = "0.00";
  ASSERT_EQ(fm_workbook_set_dxf(wb.handle, index, &replacement), 0) << fm_get_last_error_message();

  fm_dxf_t out{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, index, &out), 0);
  EXPECT_EQ(out.has_font, 1);
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.has_fill, 0);
  ASSERT_NE(out.number_format, nullptr);
  EXPECT_STREQ(out.number_format, "0.00");
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  EXPECT_EQ(fm_workbook_set_dxf(wb.handle, 1U, &replacement), FM_STATUS_INDEX_OUT_OF_RANGE);
}
