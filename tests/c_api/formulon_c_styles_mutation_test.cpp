// Stable C ABI style mutation and round-trip tests.

#include "formulon_c_record_arg.h"
#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, FullLifecycleSurvivesSaveLoad) {
  // Build font -> xf -> stamp on cell -> save -> reload -> the font is
  // resolvable via the read-side accessors.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  std::size_t font_index = 0;
  std::size_t fill_index = 0;
  std::size_t border_index = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &font_index), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &fill_index), 0);
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &border_index), 0);

  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(font_index);
  xf.fill_index = static_cast<uint32_t>(fill_index);
  xf.border_index = static_cast<uint32_t>(border_index);
  xf.num_fmt_id = 0;
  std::uint32_t xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &xf_index), 0);

  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.5), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, xf_index), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  std::uint32_t reread_xf = 0;
  EXPECT_EQ(GetCellXfIndex(wb2.handle, 0, 0, 0, &reread_xf), 0);
  fm_cell_xf_t reloaded{};
  EXPECT_EQ(fm_workbook_get_cell_xf_at(wb2.handle, reread_xf, &reloaded), 0);
  fm_font_t reloaded_font{};
  ASSERT_EQ(fm_workbook_get_font_at(wb2.handle, reloaded.font_index, &reloaded_font), 0);
  EXPECT_STREQ(reloaded_font.name, "Arial");
  EXPECT_DOUBLE_EQ(reloaded_font.size_pt, 12.0);
  EXPECT_EQ(reloaded_font.bold, 1);
}

TEST(FormulonCApiStyles, AddXfWithNonDefaultFontFillNumFmtRoundTripsThroughSaveLoad) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  std::size_t font_index = 0xFFFFFFFFU;
  std::size_t fill_index = 0xFFFFFFFFU;
  uint32_t num_fmt_id = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &font_index), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &fill_index), 0);
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "0.00%", &num_fmt_id), 0);

  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(font_index);
  xf.fill_index = static_cast<uint32_t>(fill_index);
  xf.border_index = 0;
  xf.num_fmt_id = static_cast<uint16_t>(num_fmt_id);
  std::uint32_t xf_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &xf_index), 0);
  ASSERT_NE(xf_index, 0U) << "non-default font/fill/numFmt must not dedup to the "
                             "all-zero placeholder xf at index 0";

  // Stability: re-adding the identical record must return the same index.
  std::uint32_t xf_idx_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &xf_idx_again), 0);
  EXPECT_EQ(xf_index, xf_idx_again);

  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 0.4225), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 1, 1, xf_index), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);

  std::uint32_t reread_xf = 0;
  ASSERT_EQ(GetCellXfIndex(wb2.handle, 0, 1, 1, &reread_xf), 0);
  ASSERT_EQ(reread_xf, xf_index);

  fm_cell_xf_t reloaded{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(wb2.handle, reread_xf, &reloaded), 0);
  EXPECT_EQ(reloaded.font_index, font_index);
  EXPECT_EQ(reloaded.fill_index, fill_index);
  EXPECT_EQ(reloaded.num_fmt_id, num_fmt_id);

  fm_font_t reloaded_font{};
  ASSERT_EQ(fm_workbook_get_font_at(wb2.handle, reloaded.font_index, &reloaded_font), 0);
  EXPECT_STREQ(reloaded_font.name, "Arial");
  EXPECT_DOUBLE_EQ(reloaded_font.size_pt, 12.0);
  EXPECT_EQ(reloaded_font.bold, 1);

  fm_fill_t reloaded_fill{};
  ASSERT_EQ(fm_workbook_get_fill_at(wb2.handle, reloaded.fill_index, &reloaded_fill), 0);
  EXPECT_EQ(reloaded_fill.pattern, 1);
  EXPECT_EQ(reloaded_fill.foreground.argb, 0xFFFF0000U);

  const char* reloaded_fmt = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb2.handle, reloaded.num_fmt_id, &present, &reloaded_fmt), 0);
  ASSERT_NE(reloaded_fmt, nullptr);
  EXPECT_STREQ(reloaded_fmt, "0.00%");
}

TEST(FormulonCApiStyles, JustifyLastLineRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_cell_xf_t xf{};
  xf.has_alignment = 1;
  xf.horizontal = 7;  // distributed
  xf.has_horizontal = 1;
  xf.vertical = 2;  // bottom
  xf.justify_last_line = 1;
  xf.has_justify_last_line = 1;
  std::uint32_t xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &xf_index), 0);

  std::uint32_t duplicate_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &duplicate_index), 0);
  EXPECT_EQ(duplicate_index, xf_index);

  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "distributed"), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, xf_index), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  std::uint32_t reread_index = 0;
  ASSERT_EQ(GetCellXfIndex(reloaded.handle, 0, 0, 0, &reread_index), 0);
  fm_cell_xf_t reread{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(reloaded.handle, reread_index, &reread), 0);
  EXPECT_EQ(reread.horizontal, 7);
  EXPECT_EQ(reread.justify_last_line, 1);
}

TEST(FormulonCApiStyles, NamedCellStyleRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_xf_t style_xf{};
  style_xf.has_alignment = 1;
  style_xf.horizontal = 2;
  style_xf.has_horizontal = 1;
  std::uint32_t xf_id = 0;
  ASSERT_EQ(fm_workbook_add_cell_style_xf(wb.handle, &style_xf, &xf_id), 0);
  fm_cell_style_t highlight{};
  highlight.name = "Highlight";
  highlight.style_xf_index = static_cast<uint32_t>(xf_id);
  highlight.has_builtin_id = 0;
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &highlight), 0);

  fm_cell_xf_t cell_xf{};
  cell_xf.has_alignment = 1;
  cell_xf.horizontal = 2;
  cell_xf.has_horizontal = 1;
  cell_xf.style_xf_index = static_cast<uint32_t>(xf_id);
  std::uint32_t cell_xf_id = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &cell_xf, &cell_xf_id), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "styled"), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, cell_xf_id), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  std::size_t style_count = 0;
  ASSERT_EQ(fm_workbook_get_cell_style_count(loaded.handle, &style_count), 0);
  // The seeded `Normal` keeps index 0; `Highlight` is appended after it.
  ASSERT_EQ(style_count, 2U);
  fm_cell_style_t style{};
  ASSERT_EQ(fm_workbook_get_cell_style_at(loaded.handle, 1, &style), 0);
  EXPECT_STREQ(style.name, "Highlight");
  std::uint32_t reread_cell_xf = 0;
  ASSERT_EQ(GetCellXfIndex(loaded.handle, 0, 0, 0, &reread_cell_xf), 0);
  fm_cell_xf_t reread{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(loaded.handle, reread_cell_xf, &reread), 0);
  EXPECT_EQ(reread.style_xf_index, style.style_xf_index);
}

TEST(FormulonCApiStyles, NamedStyleXfRejectsDanglingReferences) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  // A cell xf may only inherit from a named-style xf that already exists.
  fm_cell_xf_t cell_xf{};
  cell_xf.style_xf_index = 3;
  std::uint32_t cell_xf_index = 0;
  EXPECT_NE(fm_workbook_add_cell_xf(wb.handle, &cell_xf, &cell_xf_index), 0);

  // The named-style table validates its own font / fill / border / numFmt
  // references exactly like the cell-xf table does.
  fm_cell_xf_t style_xf{};
  style_xf.font_index = 9;
  std::uint32_t xf_id = 0;
  EXPECT_NE(fm_workbook_add_cell_style_xf(wb.handle, &style_xf, &xf_id), 0);

  style_xf.font_index = 0;
  style_xf.has_alignment = 1;
  style_xf.justify_last_line = 1;
  style_xf.has_justify_last_line = 1;
  ASSERT_EQ(fm_workbook_add_cell_style_xf(wb.handle, &style_xf, &xf_id), 0);
  fm_cell_xf_t reread{};
  ASSERT_EQ(fm_workbook_get_cell_style_xf_at(wb.handle, xf_id, &reread), 0);
  EXPECT_EQ(reread.justify_last_line, 1);

  // 0..53 is the whole OOXML ordinal space; a custom style states no id at all.
  fm_cell_style_t custom{};
  custom.name = "Custom";
  custom.style_xf_index = static_cast<uint32_t>(xf_id);
  custom.has_builtin_id = 1;
  custom.builtin_id = 54;
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &custom), FM_STATUS_INVALID_ARGUMENT);
  custom.has_builtin_id = 0;
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &custom), 0) << "builtin_id is not read for a custom style";
  custom.style_xf_index = static_cast<uint32_t>(xf_id) + 1U;
  EXPECT_NE(fm_workbook_set_cell_style(wb.handle, &custom), 0);
}

TEST(FormulonCApiStyles, AddBatchDeduplicatesAllStyleTables) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  const fm_font_t fonts[] = {MakeArial(), MakeArial()};
  const fm_fill_t fills[] = {MakeRedFill(), MakeRedFill()};
  fm_border_t border{};
  border.left.style = 1;
  border.left.color = RgbColor(0xFF000000U);
  const fm_border_t borders[] = {border, border};
  uint32_t font_indices[2]{};
  uint32_t fill_indices[2]{};
  uint32_t border_indices[2]{};
  const char* const num_fmt_codes[] = {"0.000%", "0.000%"};
  uint16_t num_fmt_ids[2]{};
  fm_cell_xf_t xf{};
  xf.font_index = 1U;
  xf.fill_index = 1U;
  xf.border_index = 1U;
  const fm_cell_xf_t xfs[] = {xf, xf};
  uint32_t xf_indices[2]{};
  fm_styles_batch_spec_t batch = FM_STRUCT_INIT(fm_styles_batch_spec_t);
  batch.fonts = fonts;
  batch.font_count = 2U;
  batch.font_indices = font_indices;
  batch.fills = fills;
  batch.fill_count = 2U;
  batch.fill_indices = fill_indices;
  batch.borders = borders;
  batch.border_count = 2U;
  batch.border_indices = border_indices;
  batch.cell_xfs = xfs;
  batch.cell_xf_count = 2U;
  batch.cell_xf_indices = xf_indices;
  batch.num_fmt_codes = num_fmt_codes;
  batch.num_fmt_count = 2U;
  batch.num_fmt_ids = num_fmt_ids;

  ASSERT_EQ(fm_workbook_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_indices[0], font_indices[1]);
  EXPECT_EQ(fill_indices[0], fill_indices[1]);
  EXPECT_EQ(border_indices[0], border_indices[1]);
  EXPECT_EQ(xf_indices[0], xf_indices[1]);
  EXPECT_EQ(num_fmt_ids[0], num_fmt_ids[1]);
}

TEST(FormulonCApiStyles, AddBatchCommitsNumFmtReferencesTransactionally) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  const fm_font_t font = MakeArial();
  uint32_t font_index = 0;
  const char* const codes[] = {"0.000%"};
  uint16_t num_fmt_ids[] = {0xBEEF};
  fm_cell_xf_t xf{};
  xf.font_index = 1U;  // default root plus the staged font
  xf.num_fmt_id = 164U;
  uint32_t xf_indices[] = {0xDEADBEEFU};
  fm_styles_batch_spec_t batch = FM_STRUCT_INIT(fm_styles_batch_spec_t);
  batch.fonts = &font;
  batch.font_count = 1U;
  batch.font_indices = &font_index;
  batch.cell_xfs = &xf;
  batch.cell_xf_count = 1U;
  batch.cell_xf_indices = xf_indices;
  batch.num_fmt_codes = codes;
  batch.num_fmt_count = 1U;
  batch.num_fmt_ids = num_fmt_ids;

  ASSERT_EQ(fm_workbook_add_batch(wb.handle, &batch), 0);
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
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  const std::string before_xml = formulon::io::write_styles(wb.handle->workbook().styles());
  const std::size_t before_fonts = wb.handle->workbook().styles().fonts.size();
  const std::size_t before_num_fmts = wb.handle->workbook().styles().num_fmts.size();
  const std::size_t before_cell_xfs = wb.handle->workbook().styles().cell_xfs.size();

  const fm_font_t font = MakeArial();
  uint32_t font_index = 0xA1A2A3A4U;
  const char* const codes[] = {"0.000%"};
  uint16_t num_fmt_ids[] = {0xBEEF};
  fm_cell_xf_t xf{};
  xf.font_index = 1U;    // valid after staging the font
  xf.num_fmt_id = 999U;  // neither built-in nor registered
  uint32_t xf_indices[] = {0xDEADBEEFU};
  fm_styles_batch_spec_t batch = FM_STRUCT_INIT(fm_styles_batch_spec_t);
  batch.fonts = &font;
  batch.font_count = 1U;
  batch.font_indices = &font_index;
  batch.cell_xfs = &xf;
  batch.cell_xf_count = 1U;
  batch.cell_xf_indices = xf_indices;
  batch.num_fmt_codes = codes;
  batch.num_fmt_count = 1U;
  batch.num_fmt_ids = num_fmt_ids;

  EXPECT_NE(fm_workbook_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_index, 0xA1A2A3A4U);
  EXPECT_EQ(num_fmt_ids[0], 0xBEEFU);
  EXPECT_EQ(xf_indices[0], 0xDEADBEEFU);
  EXPECT_EQ(wb.handle->workbook().styles().fonts.size(), before_fonts);
  EXPECT_EQ(wb.handle->workbook().styles().num_fmts.size(), before_num_fmts);
  EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before_cell_xfs);
  EXPECT_EQ(formulon::io::write_styles(wb.handle->workbook().styles()), before_xml);
}
