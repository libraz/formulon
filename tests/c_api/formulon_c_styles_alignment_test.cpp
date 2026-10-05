// Stable C ABI alignment and presence-flag tests.

#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, CellXfAlignmentEnumsValidateRangesBeforeMutation) {
  const auto expect_invalid = [](fm_status_t status, const char* api) {
    EXPECT_EQ(status, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(api, 0), 0U) << message;
  };

  // The upper boundary of each ordinal is accepted.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 7;  // distributed
    record.has_horizontal_align = 1;
    record.vertical_align = 4;  // distributed
    record.has_vertical_align = 1;
    uint32_t index = 0;
    ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);
  }

  // The presence flag decides whether a value is read at all, so an
  // out-of-range ordinal on an omitted attribute is ignored rather than
  // rejected: the record it describes has no such attribute to be invalid.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 8;
    record.vertical_align = 5;
    uint32_t index = 0xAABBCCDDU;
    ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);
    EXPECT_NE(index, 0xAABBCCDDU);
    fm_cell_xf got{};
    ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, index, &got), 0);
    EXPECT_EQ(got.horizontal_align, 0U);
    EXPECT_EQ(got.vertical_align, 2U);
  }

  // Every add shape validates before it creates default roots or changes the
  // output index. Keep each case isolated so the size checks cover all paths.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 8;
    record.has_horizontal_align = 1;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_xfs.size();
    expect_invalid(fm_styles_add_cell_xf(wb.handle, record, &index), "fm_styles_add_cell_xf:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.vertical_align = 5;
    record.has_vertical_align = 1;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_style_xfs.size();
    expect_invalid(fm_styles_add_cell_style_xf(wb.handle, record, &index), "fm_styles_add_cell_style_xf:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_style_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 8;
    record.has_horizontal_align = 1;
    uint32_t indices[] = {0xAABBCCDDU};
    const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr, 0U,      nullptr, nullptr, 0U,
                                nullptr, &record, 1U,      indices, nullptr, 0U,      nullptr};
    const std::size_t before = wb.handle->workbook().styles().cell_xfs.size();
    expect_invalid(fm_styles_add_batch(wb.handle, &batch), "fm_styles_add_batch:");
    EXPECT_EQ(indices[0], 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before);
  }
  // A batch record naming a `<cellStyleXfs>` entry that does not exist is
  // rejected rather than emitted as a dangling `xfId`.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.xf_id = 3U;
    uint32_t indices[] = {0xAABBCCDDU};
    const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr, 0U,      nullptr, nullptr, 0U,
                                nullptr, &record, 1U,      indices, nullptr, 0U,      nullptr};
    expect_invalid(fm_styles_add_batch(wb.handle, &batch), "fm_styles_add_batch:");
    EXPECT_EQ(indices[0], 0xAABBCCDDU);
  }
}

TEST(FormulonCApiStyles, AddBatchEmitsAlignmentOnlyForFlaggedAttributes) {
  const auto batch_xf_alignment = [](const fm_cell_xf& record) {
    WorkbookGuard wb;
    EXPECT_EQ(fm_workbook_create(&wb.handle), 0);
    uint32_t xf_indices[] = {0xDEADBEEFU};
    const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr,    0U,      nullptr, nullptr, 0U,
                                nullptr, &record, 1U,      xf_indices, nullptr, 0U,      nullptr};
    EXPECT_EQ(fm_styles_add_batch(wb.handle, &batch), 0) << fm_last_error_message();
    fm_cell_xf got{};
    EXPECT_EQ(fm_styles_get_cell_xf(wb.handle, xf_indices[0], &got), 0);
    return got;
  };

  // The pre-collapse caller shape: a value with no presence flag.
  fm_cell_xf inferred{};
  inferred.horizontal_align = 3;  // center-continuous
  const fm_cell_xf without_flag = batch_xf_alignment(inferred);
  EXPECT_EQ(without_flag.has_alignment, 0);
  EXPECT_EQ(without_flag.has_horizontal_align, 0);
  EXPECT_EQ(without_flag.horizontal_align, 0U);

  // The same value, now declared present, is the post-collapse spelling.
  fm_cell_xf flagged{};
  flagged.horizontal_align = 3;
  flagged.has_horizontal_align = 1;
  const fm_cell_xf with_flag = batch_xf_alignment(flagged);
  EXPECT_EQ(with_flag.has_alignment, 1);
  EXPECT_EQ(with_flag.has_horizontal_align, 1);
  EXPECT_EQ(with_flag.horizontal_align, 3U);
}

TEST(FormulonCApiStyles, ZeroInitializedCellXfUsesDefaultAlignment) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf record{};
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);
  EXPECT_EQ(index, 0U);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, index, &got), 0);
  EXPECT_EQ(got.horizontal_align, 0U);
  EXPECT_EQ(got.vertical_align, 2U);
  EXPECT_EQ(got.wrap_text, 0);
  EXPECT_EQ(got.justify_last_line, 0);
  EXPECT_EQ(got.has_alignment, 0);
  EXPECT_EQ(got.has_horizontal_align, 0);
  EXPECT_EQ(got.has_vertical_align, 0);
  EXPECT_EQ(got.has_wrap_text, 0);
  EXPECT_EQ(got.has_justify_last_line, 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  const std::size_t cell_xfs_end = styles_xml.find("</cellXfs>", cell_xfs_begin);
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  ASSERT_NE(cell_xfs_end, std::string::npos);
  EXPECT_EQ(styles_xml.substr(cell_xfs_begin, cell_xfs_end - cell_xfs_begin).find("<alignment"), std::string::npos);
}

TEST(FormulonCApiStyles, CellXfIgnoresPoisonForOmittedAlignmentAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf poison{};
  poison.horizontal_align = 0xFFU;
  poison.vertical_align = 0xFFU;
  poison.wrap_text = 7;
  poison.justify_last_line = 9;

  uint32_t omitted_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, poison, &omitted_index), 0);
  EXPECT_EQ(omitted_index, 0U);

  poison.has_alignment = 1;
  uint32_t explicit_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, poison, &explicit_index), 0);

  fm_cell_xf empty{};
  empty.has_alignment = 1;
  uint32_t empty_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, empty, &empty_index), 0);
  EXPECT_EQ(empty_index, explicit_index);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, explicit_index, &got), 0);
  EXPECT_EQ(got.horizontal_align, 0U);
  EXPECT_EQ(got.vertical_align, 2U);
  EXPECT_EQ(got.wrap_text, 0);
  EXPECT_EQ(got.justify_last_line, 0);
  EXPECT_EQ(got.has_alignment, 1);
  EXPECT_EQ(got.has_horizontal_align, 0);
  EXPECT_EQ(got.has_vertical_align, 0);
  EXPECT_EQ(got.has_wrap_text, 0);
  EXPECT_EQ(got.has_justify_last_line, 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  const std::size_t cell_xfs_end = styles_xml.find("</cellXfs>", cell_xfs_begin);
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  ASSERT_NE(cell_xfs_end, std::string::npos);
  const std::string cell_xfs = styles_xml.substr(cell_xfs_begin, cell_xfs_end - cell_xfs_begin);
  EXPECT_NE(cell_xfs.find("<alignment/>"), std::string::npos);
}

TEST(FormulonCApiStyles, CellXfExplicitTopAlignmentRemainsPresent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf record{};
  record.vertical_align = 0;  // top
  record.has_vertical_align = 1;
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, index, &got), 0);
  EXPECT_EQ(got.vertical_align, 0U);
  EXPECT_EQ(got.has_vertical_align, 1);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  const std::size_t cell_xfs_end = styles_xml.find("</cellXfs>", cell_xfs_begin);
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  ASSERT_NE(cell_xfs_end, std::string::npos);
  EXPECT_NE(styles_xml.substr(cell_xfs_begin, cell_xfs_end - cell_xfs_begin).find("<alignment vertical=\"top\"/>"),
            std::string::npos);
}

TEST(FormulonCApiStyles, ZeroInitializedCellStyleXfUsesDefaultAlignment) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf record{};
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, record, &index), 0);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_style_xf(wb.handle, index, &got), 0);
  EXPECT_EQ(got.horizontal_align, 0U);
  EXPECT_EQ(got.vertical_align, 2U);
  EXPECT_EQ(got.wrap_text, 0);
  EXPECT_EQ(got.justify_last_line, 0);
  EXPECT_EQ(got.has_alignment, 0);
  EXPECT_EQ(got.has_vertical_align, 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t style_xfs_begin = styles_xml.find("<cellStyleXfs");
  const std::size_t style_xfs_end = styles_xml.find("</cellStyleXfs>", style_xfs_begin);
  ASSERT_NE(style_xfs_begin, std::string::npos);
  ASSERT_NE(style_xfs_end, std::string::npos);
  EXPECT_EQ(styles_xml.substr(style_xfs_begin, style_xfs_end - style_xfs_begin).find("<alignment"), std::string::npos);
}

TEST(FormulonCApiStyles, CellXfVerticalAlignZeroIsExplicitTopOnlyWithItsPresenceFlag) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Without the presence flag, `vertical_align` is not read at all: the record
  // describes an omitted attribute and canonicalizes to the model default.
  fm_cell_xf omitted{};
  omitted.vertical_align = 0;  // top
  uint32_t omitted_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, omitted, &omitted_index), 0);
  fm_cell_xf got_omitted{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, omitted_index, &got_omitted), 0);
  EXPECT_EQ(got_omitted.vertical_align, 2U);  // bottom
  EXPECT_EQ(got_omitted.has_vertical_align, 0);

  fm_cell_xf explicit_top{};
  explicit_top.has_alignment = 1;
  explicit_top.vertical_align = 0;  // top
  explicit_top.has_vertical_align = 1;
  uint32_t explicit_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, explicit_top, &explicit_index), 0);
  EXPECT_NE(explicit_index, omitted_index);

  fm_cell_xf got_explicit{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, explicit_index, &got_explicit), 0);
  EXPECT_EQ(got_explicit.vertical_align, 0U);
  EXPECT_EQ(got_explicit.has_vertical_align, 1);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  const std::size_t cell_xfs_end = styles_xml.find("</cellXfs>", cell_xfs_begin);
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  ASSERT_NE(cell_xfs_end, std::string::npos);
  EXPECT_NE(styles_xml.substr(cell_xfs_begin, cell_xfs_end - cell_xfs_begin).find("<alignment vertical=\"top\"/>"),
            std::string::npos);
}

TEST(FormulonCApiStyles, CellXfPreservesOptionalAlignmentAndPresence) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf xf{};
  xf.vertical_align = 2;
  xf.justify_last_line = 1;
  xf.xf_id = 0;
  xf.has_alignment = 1;
  xf.has_text_rotation = 1;
  xf.text_rotation = 255;
  xf.has_indent = 1;
  xf.indent = 7;
  xf.has_relative_indent = 1;
  xf.relative_indent = -3;
  xf.has_shrink_to_fit = 1;
  xf.shrink_to_fit = 0;
  xf.has_reading_order = 1;
  xf.reading_order = 2;
  xf.has_horizontal_align = 1;
  xf.has_vertical_align = 1;
  xf.has_wrap_text = 1;
  xf.has_justify_last_line = 1;

  uint32_t index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);
  uint32_t duplicate = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &duplicate), 0);
  EXPECT_EQ(duplicate, index);

  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, index, &reread), 0);
  EXPECT_EQ(reread.justify_last_line, 1);
  EXPECT_EQ(reread.xf_id, 0U);
  EXPECT_EQ(reread.has_alignment, 1);
  EXPECT_EQ(reread.has_text_rotation, 1);
  EXPECT_EQ(reread.text_rotation, 255U);
  EXPECT_EQ(reread.has_indent, 1);
  EXPECT_EQ(reread.indent, 7U);
  EXPECT_EQ(reread.has_relative_indent, 1);
  EXPECT_EQ(reread.relative_indent, -3);
  EXPECT_EQ(reread.has_shrink_to_fit, 1);
  EXPECT_EQ(reread.shrink_to_fit, 0);
  EXPECT_EQ(reread.has_reading_order, 1);
  EXPECT_EQ(reread.reading_order, 2U);
  EXPECT_EQ(reread.has_horizontal_align, 1);
  EXPECT_EQ(reread.has_vertical_align, 1);
  EXPECT_EQ(reread.has_wrap_text, 1);
  EXPECT_EQ(reread.has_justify_last_line, 1);

  // Presence is part of the dedup key: explicit zero / false differs from
  // an omitted attribute even when the effective value is the same.
  fm_cell_xf absent = xf;
  absent.has_text_rotation = 0;
  absent.has_indent = 0;
  absent.has_relative_indent = 0;
  absent.has_shrink_to_fit = 0;
  absent.has_reading_order = 0;
  uint32_t absent_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, absent, &absent_index), 0);
  EXPECT_NE(absent_index, index);
}

TEST(FormulonCApiStyles, CellXfRejectsInvalidExcelAlignmentRanges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  uint32_t index = 0;

  xf.has_text_rotation = 1;
  xf.text_rotation = 181;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);
  xf.text_rotation = 255;
  EXPECT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);

  xf.has_text_rotation = 0;
  xf.has_indent = 1;
  xf.indent = 256;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);
  xf.indent = 255;
  EXPECT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);

  xf.has_indent = 0;
  xf.has_reading_order = 1;
  xf.reading_order = 3;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);
}

TEST(FormulonCApiStyles, CellXfPresenceFlagsDistinguishExplicitDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf explicit_defaults{};
  explicit_defaults.vertical_align = 2;
  explicit_defaults.has_alignment = 1;
  explicit_defaults.has_horizontal_align = 1;
  explicit_defaults.has_vertical_align = 1;
  explicit_defaults.has_wrap_text = 1;
  explicit_defaults.has_justify_last_line = 1;
  uint32_t explicit_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, explicit_defaults, &explicit_index), 0);

  fm_cell_xf omitted_defaults = explicit_defaults;
  omitted_defaults.has_horizontal_align = 0;
  omitted_defaults.has_vertical_align = 0;
  omitted_defaults.has_wrap_text = 0;
  omitted_defaults.has_justify_last_line = 0;
  uint32_t omitted_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, omitted_defaults, &omitted_index), 0);
  EXPECT_NE(explicit_index, omitted_index);

  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, explicit_index, &reread), 0);
  EXPECT_EQ(reread.has_horizontal_align, 1);
  EXPECT_EQ(reread.has_vertical_align, 1);
  EXPECT_EQ(reread.has_wrap_text, 1);
  EXPECT_EQ(reread.has_justify_last_line, 1);
}

TEST(FormulonCApiStyles, CellStyleXfRoundTripsOptionalAlignment) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf style{};
  style.has_text_rotation = 1;
  style.text_rotation = 0;
  style.has_relative_indent = 1;
  style.relative_indent = -2;
  style.has_shrink_to_fit = 1;
  style.shrink_to_fit = false;
  uint32_t style_index = 0;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, style, &style_index), 0);

  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_style_xf(wb.handle, style_index, &reread), 0);
  EXPECT_EQ(reread.has_text_rotation, 1);
  EXPECT_EQ(reread.text_rotation, 0U);
  EXPECT_EQ(reread.has_relative_indent, 1);
  EXPECT_EQ(reread.relative_indent, -2);
  EXPECT_EQ(reread.has_shrink_to_fit, 1);
  EXPECT_EQ(reread.shrink_to_fit, 0);
}
