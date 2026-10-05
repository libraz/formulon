// Stable C ABI style mutation and round-trip tests.

#include "formulon_c_styles_test_helpers.h"

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

TEST(FormulonCApiStyles, JustifyLastLineRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf xf{};
  xf.has_alignment = 1;
  xf.horizontal_align = 7;  // distributed
  xf.has_horizontal_align = 1;
  xf.vertical_align = 2;  // bottom
  xf.justify_last_line = 1;
  xf.has_justify_last_line = 1;
  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);

  uint32_t duplicate_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &duplicate_idx), 0);
  EXPECT_EQ(duplicate_idx, xf_idx);

  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "distributed"), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  uint32_t reread_idx = 0;
  ASSERT_EQ(fm_cell_get_xf_index(reloaded.handle, 0, 0, 0, &reread_idx), 0);
  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_xf(reloaded.handle, reread_idx, &reread), 0);
  EXPECT_EQ(reread.horizontal_align, 7U);
  EXPECT_EQ(reread.justify_last_line, 1);
}

TEST(FormulonCApiStyles, NamedCellStyleRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf style_xf{};
  style_xf.has_alignment = 1;
  style_xf.horizontal_align = 2;
  style_xf.has_horizontal_align = 1;
  uint32_t xf_id = 0;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, style_xf, &xf_id), 0);
  ASSERT_EQ(fm_styles_set_cell_style(wb.handle, "Highlight", xf_id, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);

  fm_cell_xf cell_xf{};
  cell_xf.has_alignment = 1;
  cell_xf.horizontal_align = 2;
  cell_xf.has_horizontal_align = 1;
  cell_xf.xf_id = xf_id;
  uint32_t cell_xf_id = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, cell_xf, &cell_xf_id), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "styled"), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, cell_xf_id), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  uint32_t style_count = 0;
  ASSERT_EQ(fm_styles_get_cell_style_count(loaded.handle, &style_count), 0);
  // The seeded `Normal` keeps index 0; `Highlight` is appended after it.
  ASSERT_EQ(style_count, 2U);
  fm_cell_style_record_t style{};
  ASSERT_EQ(fm_styles_get_cell_style(loaded.handle, 1, &style), 0);
  EXPECT_STREQ(style.name, "Highlight");
  uint32_t reread_cell_xf = 0;
  ASSERT_EQ(fm_cell_get_xf_index(loaded.handle, 0, 0, 0, &reread_cell_xf), 0);
  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_xf(loaded.handle, reread_cell_xf, &reread), 0);
  EXPECT_EQ(reread.xf_id, style.xf_id);
}

TEST(FormulonCApiStyles, NamedStyleXfRejectsDanglingReferences) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // A cell xf may only inherit from a named-style xf that already exists.
  fm_cell_xf cell_xf{};
  cell_xf.xf_id = 3;
  uint32_t cell_xf_index = 0;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, cell_xf, &cell_xf_index), 0);

  // The named-style table validates its own font / fill / border / numFmt
  // references exactly like the cell-xf table does.
  fm_cell_xf style_xf{};
  style_xf.font_index = 9;
  uint32_t xf_id = 0;
  EXPECT_NE(fm_styles_add_cell_style_xf(wb.handle, style_xf, &xf_id), 0);

  style_xf.font_index = 0;
  style_xf.has_alignment = 1;
  style_xf.justify_last_line = 1;
  style_xf.has_justify_last_line = 1;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, style_xf, &xf_id), 0);
  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_style_xf(wb.handle, xf_id, &reread), 0);
  EXPECT_EQ(reread.justify_last_line, 1);

  // 0..47 is the whole OOXML ordinal space; anything else needs the sentinel.
  EXPECT_NE(fm_styles_set_cell_style(wb.handle, "Custom", xf_id, 48), 0);
  EXPECT_EQ(fm_styles_set_cell_style(wb.handle, "Custom", xf_id, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);
  EXPECT_NE(fm_styles_set_cell_style(wb.handle, "Custom", xf_id + 1U, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);
}

TEST(FormulonCApiStyles, AddBatchDeduplicatesAllStyleTables) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const fm_font_record fonts[] = {MakeArial(), MakeArial()};
  const fm_fill_record fills[] = {MakeRedFill(), MakeRedFill()};
  fm_border_record border{};
  border.left.style = 1;
  border.left.color_argb = 0xFF000000U;
  const fm_border_record borders[] = {border, border};
  uint32_t font_indices[2]{};
  uint32_t fill_indices[2]{};
  uint32_t border_indices[2]{};
  const char* const num_fmt_codes[] = {"0.000%", "0.000%"};
  uint16_t num_fmt_ids[2]{};
  fm_cell_xf xf{};
  xf.font_index = 1U;
  xf.fill_index = 1U;
  xf.border_index = 1U;
  const fm_cell_xf xfs[] = {xf, xf};
  uint32_t xf_indices[2]{};
  const fm_styles_batch batch{fonts, 2U, font_indices, fills,         2U, fill_indices, borders, 2U, border_indices,
                              xfs,   2U, xf_indices,   num_fmt_codes, 2U, num_fmt_ids};

  ASSERT_EQ(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_indices[0], font_indices[1]);
  EXPECT_EQ(fill_indices[0], fill_indices[1]);
  EXPECT_EQ(border_indices[0], border_indices[1]);
  EXPECT_EQ(xf_indices[0], xf_indices[1]);
  EXPECT_EQ(num_fmt_ids[0], num_fmt_ids[1]);
}

TEST(FormulonCApiStyles, AddBatchCommitsNumFmtReferencesTransactionally) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const fm_font_record font = MakeArial();
  uint32_t font_index = 0;
  const char* const codes[] = {"0.000%"};
  uint16_t num_fmt_ids[] = {0xBEEF};
  fm_cell_xf xf{};
  xf.font_index = 1U;  // default root plus the staged font
  xf.num_fmt_id = 164U;
  uint32_t xf_indices[] = {0xDEADBEEFU};
  const fm_styles_batch batch{&font,   1U,  &font_index, nullptr,    0U,    nullptr, nullptr,    0U,
                              nullptr, &xf, 1U,          xf_indices, codes, 1U,      num_fmt_ids};

  ASSERT_EQ(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_index, 1U);
  EXPECT_EQ(num_fmt_ids[0], 164U);
  EXPECT_EQ(xf_indices[0], 1U);
  ASSERT_EQ(wb.handle->workbook().styles().num_fmts.size(), 1U);
  EXPECT_EQ(wb.handle->workbook().styles().num_fmts[0].id, 164U);
  ASSERT_EQ(wb.handle->workbook().styles().cell_xfs.size(), 2U);
  EXPECT_EQ(wb.handle->workbook().styles().cell_xfs[1].num_fmt_id, 164U);
}

TEST(FormulonCApiStyles, AddBatchFailureLeavesTableAndOutputsUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const std::string before_xml = formulon::io::write_styles(wb.handle->workbook().styles());
  const std::size_t before_fonts = wb.handle->workbook().styles().fonts.size();
  const std::size_t before_num_fmts = wb.handle->workbook().styles().num_fmts.size();
  const std::size_t before_cell_xfs = wb.handle->workbook().styles().cell_xfs.size();

  const fm_font_record font = MakeArial();
  uint32_t font_index = 0xA1A2A3A4U;
  const char* const codes[] = {"0.000%"};
  uint16_t num_fmt_ids[] = {0xBEEF};
  fm_cell_xf xf{};
  xf.font_index = 1U;    // valid after staging the font
  xf.num_fmt_id = 999U;  // neither built-in nor registered
  uint32_t xf_indices[] = {0xDEADBEEFU};
  const fm_styles_batch batch{&font,   1U,  &font_index, nullptr,    0U,    nullptr, nullptr,    0U,
                              nullptr, &xf, 1U,          xf_indices, codes, 1U,      num_fmt_ids};

  EXPECT_NE(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_index, 0xA1A2A3A4U);
  EXPECT_EQ(num_fmt_ids[0], 0xBEEFU);
  EXPECT_EQ(xf_indices[0], 0xDEADBEEFU);
  EXPECT_EQ(wb.handle->workbook().styles().fonts.size(), before_fonts);
  EXPECT_EQ(wb.handle->workbook().styles().num_fmts.size(), before_num_fmts);
  EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before_cell_xfs);
  EXPECT_EQ(formulon::io::write_styles(wb.handle->workbook().styles()), before_xml);
}
