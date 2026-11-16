// Stable C ABI font tests.

#include <array>
#include <limits>

#include "font_completion.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_styles_test_helpers.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr std::array<double, 8> kInvalidFontSizes = {std::numeric_limits<double>::quiet_NaN(),
                                                     std::numeric_limits<double>::infinity(),
                                                     -std::numeric_limits<double>::infinity(),
                                                     -1.0,
                                                     0.0,
                                                     -0.0,
                                                     0.5,
                                                     410.0};

std::string StylesXml(fm_workbook_t* wb) {
  return formulon::io::write_styles(wb->workbook().styles());
}

}  // namespace

TEST(FormulonCApiStyles, AddFontDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t a = 0xFFFFFFFFU;
  std::uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &a), 0);
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddFontDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t a = 0;
  std::uint32_t b = 0;
  fm_font_t arial = MakeArial();
  fm_font_t calibri = MakeArial();
  calibri.name = "Calibri";
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &arial, &a), 0);
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &calibri, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddFontGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t before = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &before), 0);
  std::uint32_t idx = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &idx), 0);
  std::size_t after = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &after), 0);
  // A new workbook seeds the default font at index 0, so the caller's
  // first font lands at index 1.
  EXPECT_EQ(before, 1U);
  EXPECT_EQ(after, 2U);
  EXPECT_EQ(idx, 1U);
  // Round-trip: the freshly added font should read back equal.
  fm_font_t out{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, idx, &out), 0);
  EXPECT_STREQ(out.name, "Arial");
  EXPECT_DOUBLE_EQ(out.size_pt, 12.0);
  EXPECT_EQ(out.color.argb, 0xFF112233U);
  EXPECT_EQ(out.bold, 1);
}

TEST(FormulonCApiStyles, GetFontRoundTripIsIdentityForSuperscript) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_font_t superscript = MakeArial();
  superscript.vert_align = 1;
  std::uint32_t idx = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &superscript, &idx), 0);

  fm_font_t loaded{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, idx, &loaded), 0);
  EXPECT_STREQ(loaded.name, "Arial");
  EXPECT_EQ(loaded.vert_align, 1);

  std::size_t count_before = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &count_before), 0);
  std::uint32_t reindex = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &loaded, &reindex), 0);
  EXPECT_EQ(reindex, idx);
  std::size_t count_after = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count_before);
}

TEST(FormulonCApiStyles, GetEditAddFontPreservesTheUneditedFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_font_t superscript = MakeArial();
  superscript.vert_align = 1;
  superscript.has_family = 1;
  superscript.family = 2;
  superscript.has_charset = 1;
  superscript.charset = 128;
  std::uint32_t idx = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &superscript, &idx), 0);

  fm_font_t edited{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, idx, &edited), 0);
  edited.color = RgbColor(0xFF00FF00U);
  std::uint32_t recolored = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &edited, &recolored), 0);
  EXPECT_NE(recolored, idx);

  fm_font_t reread{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, recolored, &reread), 0);
  EXPECT_EQ(reread.color.argb, 0xFF00FF00U);
  EXPECT_EQ(reread.vert_align, 1);
  EXPECT_EQ(reread.has_family, 1);
  EXPECT_EQ(reread.family, 2);
  EXPECT_EQ(reread.has_charset, 1);
  EXPECT_EQ(reread.charset, 128U);
}

TEST(FormulonCApiStyles, AddFontIsIdentityAgainstAFileLoadedTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));
  // The table is read directly, so complete its fonts the way a file load does.
  for (formulon::FontRecord& font : wb.handle->workbook().mutable_styles().fonts) {
    formulon::complete_font_from_face(font, wb.handle->workbook().excel_profile().locale);
  }

  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &count), 0);
  ASSERT_EQ(count, 3U);
  for (uint32_t i = 0; i < count; ++i) {
    fm_font_t loaded{};
    ASSERT_EQ(fm_workbook_get_font_at(wb.handle, i, &loaded), 0) << "font " << i;
    std::uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_workbook_add_font(wb.handle, &loaded, &reindex), 0) << "font " << i;
    EXPECT_EQ(reindex, i);
  }
  std::size_t count_after = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count);
}

TEST(FormulonCApiStyles, AddFontDoesNotAliasAThemeColourOntoLiteralRgb) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  // Font 0 is `<color theme="1"/>`; folding it together with a caller's
  // literal RGB record would lose the selector and change the template's
  // body text semantics.
  fm_font_t themed{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, 0U, &themed), 0);
  ASSERT_EQ(themed.color.kind, static_cast<uint8_t>(FM_COLOR_KIND_THEME));

  fm_font_t literal_rgb = themed;
  literal_rgb.color = RgbColor(0xFF112233U);
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &literal_rgb, &index), 0);
  EXPECT_NE(index, 0U);

  // The theme record is untouched and still serialises as a theme colour.
  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<color theme=\"1\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<color theme=\"0\" tint=\"-0.25\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, AddFontDistinguishesAnExplicitOffToggleFromAnAbsentOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_font_t absent = MakeArial();
  absent.bold = 0;
  absent.has_bold = 0;
  fm_font_t explicit_off = absent;
  explicit_off.has_bold = 1;

  std::uint32_t absent_index = 0;
  std::uint32_t explicit_index = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &absent, &absent_index), 0);
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &explicit_off, &explicit_index), 0);
  EXPECT_NE(absent_index, explicit_index);
}

TEST(FormulonCApiStyles, AddFontRejectsInvalidSizesWithoutMutation) {
  for (const double size : kInvalidFontSizes) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const std::string before_xml = StylesXml(wb.handle);
    const std::size_t before_count = wb.handle->workbook().styles().fonts.size();

    fm_font_t font = MakeArial();
    font.size_pt = size;
    font.has_size_pt = 1;
    std::uint32_t out_index = 0xA1A2A3A4U;
    EXPECT_EQ(fm_workbook_add_font(wb.handle, &font, &out_index), kInvalidArgument);
    EXPECT_EQ(out_index, 0xA1A2A3A4U);
    EXPECT_EQ(wb.handle->workbook().styles().fonts.size(), before_count);
    EXPECT_EQ(StylesXml(wb.handle), before_xml);
  }
}

TEST(FormulonCApiStyles, SetFontRejectsInvalidSizesWithoutMutationAndKeepsIndexPrecedence) {
  for (const double size : kInvalidFontSizes) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const std::string before_xml = StylesXml(wb.handle);
    const std::size_t before_count = wb.handle->workbook().styles().fonts.size();

    fm_font_t font = MakeArial();
    font.size_pt = size;
    font.has_size_pt = 1;
    EXPECT_EQ(fm_workbook_set_font(wb.handle, 0U, &font), kInvalidArgument);
    EXPECT_EQ(wb.handle->workbook().styles().fonts.size(), before_count);
    EXPECT_EQ(StylesXml(wb.handle), before_xml);
  }

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_font_t invalid = MakeArial();
  invalid.size_pt = std::numeric_limits<double>::quiet_NaN();
  EXPECT_EQ(fm_workbook_set_font(wb.handle, std::numeric_limits<uint32_t>::max(), &invalid),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_NE(std::string(fm_get_last_error_message()).find("font_index"), std::string::npos);
  EXPECT_EQ(fm_workbook_set_font(nullptr, 0U, &invalid),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiStyles, SetDefaultFontRejectsInvalidSizesWithoutMutation) {
  for (const double size : kInvalidFontSizes) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const std::string before_xml = StylesXml(wb.handle);
    const std::size_t before_count = wb.handle->workbook().styles().fonts.size();

    fm_font_t font = MakeArial();
    font.size_pt = size;
    font.has_size_pt = 1;
    EXPECT_EQ(fm_workbook_set_default_font(wb.handle, &font), kInvalidArgument);
    EXPECT_EQ(wb.handle->workbook().styles().fonts.size(), before_count);
    EXPECT_EQ(StylesXml(wb.handle), before_xml);
  }
}

TEST(FormulonCApiStyles, AddBatchRejectsInvalidFontsTransactionally) {
  for (const double size : kInvalidFontSizes) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const std::string before_xml = StylesXml(wb.handle);
    const std::size_t before_count = wb.handle->workbook().styles().fonts.size();

    fm_font_t fonts[] = {MakeArial(), MakeArial()};
    fonts[1].size_pt = size;
    fonts[1].has_size_pt = 1;
    uint32_t font_indices[] = {0xA1A2A3A4U, 0xB1B2B3B4U};
    fm_styles_batch_spec_t batch = FM_STRUCT_INIT(fm_styles_batch_spec_t);
    batch.fonts = fonts;
    batch.font_count = 2U;
    batch.font_indices = font_indices;

    EXPECT_EQ(fm_workbook_add_batch(wb.handle, &batch), kInvalidArgument);
    EXPECT_EQ(font_indices[0], 0xA1A2A3A4U);
    EXPECT_EQ(font_indices[1], 0xB1B2B3B4U);
    EXPECT_EQ(wb.handle->workbook().styles().fonts.size(), before_count);
    EXPECT_EQ(StylesXml(wb.handle), before_xml);
  }
}

TEST(FormulonCApiStyles, AddDxfRejectsInvalidPresentFontSizesWithoutMutation) {
  for (const double size : kInvalidFontSizes) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const std::string before_xml = StylesXml(wb.handle);
    const std::size_t before_count = wb.handle->workbook().styles().dxfs.size();

    fm_dxf_t dxf{};
    dxf.has_font = 1;
    dxf.font = MakeArial();
    dxf.font.size_pt = size;
    dxf.font.has_size_pt = 1;
    std::uint32_t out_index = 0xA1A2A3A4U;
    EXPECT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &out_index), kInvalidArgument);
    EXPECT_EQ(out_index, 0xA1A2A3A4U);
    EXPECT_EQ(wb.handle->workbook().styles().dxfs.size(), before_count);
    EXPECT_EQ(StylesXml(wb.handle), before_xml);
  }
}

TEST(FormulonCApiStyles, AddDxfIgnoresAbsentFontSizePayloads) {
  const std::array<fm_dxf_t, 2> dxfs = [] {
    std::array<fm_dxf_t, 2> records{};
    records[0].has_font = 0;
    records[0].font.has_size_pt = 1;
    records[0].font.size_pt = std::numeric_limits<double>::quiet_NaN();
    records[1].has_font = 1;
    records[1].font.has_size_pt = 0;
    records[1].font.size_pt = -1.0;
    return records;
  }();

  for (const fm_dxf_t& dxf : dxfs) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    std::uint32_t out_index = 0xA1A2A3A4U;
    EXPECT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &out_index), 0) << fm_get_last_error_message();
    EXPECT_EQ(out_index, 0U);
    EXPECT_EQ(wb.handle->workbook().styles().dxfs.size(), 1U);
  }
}

TEST(FormulonCApiStyles, CellAndDifferentialFontsAcceptTheBoundarySizes) {
  for (const double size : std::array<double, 2>{1.0, 409.0}) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

    fm_font_t font = MakeArial();
    font.size_pt = size;
    font.has_size_pt = 1;
    std::uint32_t index = 0xFFFFFFFFU;
    EXPECT_EQ(fm_workbook_add_font(wb.handle, &font, &index), 0) << fm_get_last_error_message();
    EXPECT_EQ(fm_workbook_set_font(wb.handle, 0U, &font), 0) << fm_get_last_error_message();
    EXPECT_EQ(fm_workbook_set_default_font(wb.handle, &font), 0) << fm_get_last_error_message();

    const uint32_t batch_sentinel = 0xB1B2B3B4U;
    uint32_t batch_index = batch_sentinel;
    fm_styles_batch_spec_t batch = FM_STRUCT_INIT(fm_styles_batch_spec_t);
    batch.fonts = &font;
    batch.font_count = 1U;
    batch.font_indices = &batch_index;
    EXPECT_EQ(fm_workbook_add_batch(wb.handle, &batch), 0) << fm_get_last_error_message();
    EXPECT_NE(batch_index, batch_sentinel);

    fm_dxf_t dxf{};
    dxf.has_font = 1;
    dxf.font = font;
    dxf.font.has_size_pt = 1;
    std::uint32_t dxf_index = 0xFFFFFFFFU;
    EXPECT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &dxf_index), 0) << fm_get_last_error_message();
    EXPECT_EQ(dxf_index, 0U);
  }
}

TEST(FormulonCApiStyles, CellFontWithoutAStatedSizeTakesTheNormalFontSize) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_font_t normal = MakeArial();
  normal.size_pt = 14.0;
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, &normal), 0) << fm_get_last_error_message();

  fm_font_t unsized = MakeArial();
  unsized.name = "Courier New";
  unsized.has_size_pt = 0;
  unsized.size_pt = 0.0;  // Ignored without the flag.
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &unsized, &index), 0) << fm_get_last_error_message();
  fm_font_t read{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, index, &read), 0);
  EXPECT_DOUBLE_EQ(read.size_pt, 14.0);
  EXPECT_NE(read.has_size_pt, 0);

  fm_font_t stated = unsized;
  stated.has_size_pt = 1;
  stated.size_pt = 0.0;
  EXPECT_EQ(fm_workbook_add_font(wb.handle, &stated, &index), kInvalidArgument);
}

TEST(FormulonCApiStyles, SetDefaultFontWithoutAStatedSizeKeepsTheCurrentSize) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_font_t normal = MakeArial();
  normal.size_pt = 14.0;
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, &normal), 0);

  fm_font_t renamed = MakeArial();
  renamed.name = "Courier New";
  renamed.has_size_pt = 0;
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, &renamed), 0) << fm_get_last_error_message();
  fm_font_t read{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, 0, &read), 0);
  EXPECT_DOUBLE_EQ(read.size_pt, 14.0);
  EXPECT_STREQ(read.name, "Courier New");
}

TEST(FormulonCApiStyles, CellFontWithANullNameTakesTheNormalFontName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_font_t normal = MakeArial();
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, &normal), 0);

  fm_font_t unnamed = MakeArial();
  unnamed.name = nullptr;
  unnamed.bold = 0;
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &unnamed, &index), 0) << fm_get_last_error_message();
  fm_font_t read{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, index, &read), 0);
  EXPECT_STREQ(read.name, "Arial");

  // The Normal font keeps its own name when the default font is set without one.
  fm_font_t resized = MakeArial();
  resized.name = nullptr;
  resized.size_pt = 20.0;
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, &resized), 0);
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, 0, &read), 0);
  EXPECT_STREQ(read.name, "Arial");
  EXPECT_DOUBLE_EQ(read.size_pt, 20.0);
}

TEST(FormulonCApiStyles, CellFontWithNoColourReadsBackWithoutOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_font_t plain = MakeArial();
  plain.color = fm_color_t{};
  std::uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, &plain, &index), 0);
  fm_font_t read{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, index, &read), 0);
  EXPECT_EQ(read.color.kind, FM_COLOR_KIND_NONE);
}

namespace {

struct SavedFace {
  const char* face;
  int family;   // 0: no <family>
  int charset;  // 0: no <charset>
};

// Mac Excel 16.115, a bare rename of the Normal font from the ja-JP default (nf_ja_default.json).
constexpr std::array<SavedFace, 21> kRenamedFaces = {{
    {"Calibri", 2, 0},         {"Aptos", 0, 0},           {"Aptos Narrow", 0, 0},
    {"游明朝", 1, 128},        {"等线", 4, 134},          {"맑은 고딕", 2, 129},
    {"新細明體", 1, 136},      {"Tahoma", 2, 0},          {"ＭＳ Ｐゴシック", 2, 128},
    {"ＭＳ ゴシック", 2, 128}, {"ＭＳ 明朝", 1, 128},     {"Meiryo", 2, 128},
    {"Meiryo UI", 2, 128},     {"メイリオ", 2, 128},      {"Arial", 2, 0},
    {"Cambria", 1, 0},         {"Times New Roman", 1, 0}, {"Verdana", 2, 0},
    {"Helvetica", 2, 0},       {"Helvetica Neue", 2, 0},  {"Courier New", 1, 0},
}};

/// The first `<font>` of a saved package's `<fonts>` list.
std::string SavedNormalFont(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  const std::string xml = ExtractStylesXml(saved);
  const std::size_t fonts = xml.find("<fonts");
  const std::size_t begin = xml.find("<font>", fonts);
  const std::size_t end = xml.find("</font>", begin);
  if (fonts == std::string::npos || begin == std::string::npos || end == std::string::npos) {
    ADD_FAILURE() << "no font in " << xml;
    return {};
  }
  return xml.substr(begin, end - begin);
}

std::string Element(const char* tag, int value) {
  return "<" + std::string(tag) + " val=\"" + std::to_string(value) + "\"/>";
}

void ExpectNormalFont(fm_workbook_t* wb, const std::string& face, int family, int charset) {
  const std::string font = SavedNormalFont(wb);
  EXPECT_NE(font.find("<name val=\"" + face + "\"/>"), std::string::npos) << font;
  EXPECT_EQ(font.find("<family") != std::string::npos, family != 0) << font;
  EXPECT_EQ(font.find("<charset") != std::string::npos, charset != 0) << font;
  if (family != 0) {
    EXPECT_NE(font.find(Element("family", family)), std::string::npos) << font;
  }
  if (charset != 0) {
    EXPECT_NE(font.find(Element("charset", charset)), std::string::npos) << font;
  }
  EXPECT_EQ(font.find("<scheme"), std::string::npos) << font;
}

fm_status_t RenameDefaultFont(fm_workbook_t* wb, const char* face) {
  fm_font_t record{};
  record.name = face;
  return fm_workbook_set_default_font(wb, &record);
}

}  // namespace

TEST(FormulonCApiStyles, BareRenameOfTheDefaultFontSavesWhatExcelCompletes) {
  for (const SavedFace& row : kRenamedFaces) {
    SCOPED_TRACE(row.face);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
    ASSERT_EQ(RenameDefaultFont(wb.handle, row.face), 0) << fm_get_last_error_message();
    ExpectNormalFont(wb.handle, row.face, row.family, row.charset);
  }
}

TEST(FormulonCApiStyles, BareRenameToYuGothicIsFamilyThreeButTheStockDefaultKeepsItsOwn) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  ASSERT_EQ(RenameDefaultFont(wb.handle, "Arial"), 0);
  ASSERT_EQ(RenameDefaultFont(wb.handle, "游ゴシック"), 0) << fm_get_last_error_message();
  ExpectNormalFont(wb.handle, "游ゴシック", 3, 128);

  WorkbookGuard stock;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &stock.handle), 0);
  const std::string font = SavedNormalFont(stock.handle);
  EXPECT_NE(font.find("<name val=\"游ゴシック\"/>"), std::string::npos) << font;
  EXPECT_NE(font.find(Element("family", 2)), std::string::npos) << font;
  EXPECT_NE(font.find(Element("charset", 128)), std::string::npos) << font;
  EXPECT_NE(font.find("<scheme val=\"minor\"/>"), std::string::npos) << font;
}

TEST(FormulonCApiStyles, StatedCharsetOrFamilyOnTheDefaultFontIsNotCompleted) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  fm_font_t record{};
  record.name = "Courier New";
  record.has_charset = 1;
  record.charset = 128;
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, &record), 0);
  ExpectNormalFont(wb.handle, "Courier New", 0, 128);

  record = fm_font_t{};
  record.name = "Arial";
  record.has_family = 1;
  record.family = 1;
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, &record), 0);
  ExpectNormalFont(wb.handle, "Arial", 1, 0);
}

TEST(FormulonCApiStyles, FaceOutsideTheCompletionTableIsSavedWithNeither) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  ASSERT_EQ(RenameDefaultFont(wb.handle, "NoSuchFontXyz"), 0);
  ExpectNormalFont(wb.handle, "NoSuchFontXyz", 0, 0);
}

TEST(FormulonCApiStyles, ReadingTheDefaultFontBackAndRenamingItKeepsTheStatedCharset) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  fm_font_t read{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, 0, &read), 0);
  ASSERT_NE(read.has_charset, 0);
  read.name = "Tahoma";
  read.scheme = 0;
  ASSERT_EQ(fm_workbook_set_default_font(wb.handle, &read), 0);
  ExpectNormalFont(wb.handle, "Tahoma", read.has_family != 0 ? static_cast<int>(read.family) : 0,
                   static_cast<int>(read.charset));
}

TEST(FormulonCApiStyles, RenamingTheDefaultFontToAFaceOutsideTheTableStatesNeitherFamilyNorCharset) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  ASSERT_EQ(RenameDefaultFont(wb.handle, "NoSuchFontXyz"), 0);
  fm_font_t read{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_family, 0);
  EXPECT_EQ(read.has_charset, 0);
}

TEST(FormulonCApiStyles, RenamingTheDefaultFontStatesTheFontZeroFamilyAtTheSetter) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  ASSERT_EQ(RenameDefaultFont(wb.handle, "Arial"), 0);
  ASSERT_EQ(RenameDefaultFont(wb.handle, "游ゴシック"), 0);
  fm_font_t read{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.has_family, 1);
  EXPECT_EQ(read.family, 3);
  EXPECT_EQ(read.has_charset, 1);
  EXPECT_EQ(read.charset, 128);
  EXPECT_EQ(read.scheme, 0);
}

namespace {

struct LoadedFace {
  const char* face;
  int family;
  int charset;
};

/// A workbook of `profile` whose font table holds one literal font per face with no family, charset or scheme (the
/// styles table is edited directly: the API completes such a font when it is added), saved in `format`.
void SaveFamilyLessFaces(const char* profile, int32_t format, const LoadedFace* faces, size_t count, BufferGuard* out) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile(profile, &wb.handle), 0);
  for (size_t i = 0; i < count; ++i) {
    formulon::FontRecord font;
    font.name = faces[i].face;
    wb.handle->workbook().mutable_styles().fonts.push_back(font);
  }
  ASSERT_EQ(fm_workbook_save(wb.handle, format, &out->data, &out->len), 0) << fm_get_last_error_message();
}

/// The font records appended after the Normal font, in file order, of a package loaded under `profile`.
void ExpectLoadedFaces(const BufferGuard& saved, const char* profile, const LoadedFace* faces, size_t count) {
  fm_load_options_spec_t options = FM_STRUCT_INIT(fm_load_options_spec_t);
  options.profile_id = profile;
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &options, &loaded.handle), 0) << fm_get_last_error_message();
  size_t fonts = 0;
  ASSERT_EQ(fm_workbook_get_font_count(loaded.handle, &fonts), 0);
  ASSERT_GE(fonts, count + 1);
  for (size_t i = 0; i < count; ++i) {
    SCOPED_TRACE(std::string(profile) + " " + faces[i].face);
    fm_font_t font{};
    ASSERT_EQ(fm_workbook_get_font_at(loaded.handle, static_cast<uint32_t>(fonts - count + i), &font), 0);
    EXPECT_STREQ(font.name, faces[i].face);
    EXPECT_EQ(font.has_family != 0, faces[i].family != 0);
    EXPECT_EQ(font.has_charset != 0, faces[i].charset != 0);
    if (faces[i].family != 0) {
      EXPECT_EQ(font.family, faces[i].family);
    }
    if (faces[i].charset != 0) {
      EXPECT_EQ(font.charset, faces[i].charset);
    }
    EXPECT_EQ(font.scheme, 0);
  }
}

}  // namespace

TEST(FormulonCApiStyles, LoadCompletesFamilyAndCharsetOfAFamilyLessFaceInBothPackageFormats) {
  // The measured table of a ja-profile load; 游ゴシック is family 3 here and has none under en.
  constexpr std::array<LoadedFace, 15> kJa = {{{"Arial", 2, 0},
                                               {"Verdana", 2, 0},
                                               {"Helvetica Neue", 2, 0},
                                               {"Cambria", 1, 0},
                                               {"游ゴシック", 3, 128},
                                               {"游明朝", 1, 128},
                                               {"等线", 4, 134},
                                               {"맑은 고딕", 2, 129},
                                               {"新細明體", 1, 136},
                                               {"ＭＳ ゴシック", 2, 128},
                                               {"ＭＳ Ｐゴシック", 2, 128},
                                               {"ＭＳ 明朝", 1, 128},
                                               {"Meiryo", 2, 128},
                                               {"Meiryo UI", 2, 128},
                                               {"Aptos Narrow", 0, 0}}};
  for (const int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    SCOPED_TRACE(format);
    BufferGuard saved;
    SaveFamilyLessFaces("mac-365-ja_JP", format, kJa.data(), kJa.size(), &saved);
    ExpectLoadedFaces(saved, "mac-365-ja_JP", kJa.data(), kJa.size());
  }
}

TEST(FormulonCApiStyles, LoadGivesYuGothicNoFamilyUnderAnEnglishProfileAndNormalFaceNothingExtra) {
  constexpr std::array<LoadedFace, 3> kEn = {{{"游ゴシック", 0, 128},
                                              {"Aptos Narrow", 0, 0},  // the en Normal face: no family copied from it
                                              {"Calibri", 2, 0}}};
  for (const int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    SCOPED_TRACE(format);
    BufferGuard saved;
    SaveFamilyLessFaces("mac-365-en_US", format, kEn.data(), kEn.size(), &saved);
    ExpectLoadedFaces(saved, "mac-365-en_US", kEn.data(), kEn.size());
  }
}

TEST(FormulonCApiStyles, LoadLeavesAPartialOrThemeLinkedFontAlone) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-en_US", &wb.handle), 0);
  formulon::StylesTable& styles = wb.handle->workbook().mutable_styles();
  formulon::FontRecord charset_only;
  charset_only.name = "Arial";
  charset_only.has_charset = true;
  charset_only.charset = 1;
  formulon::FontRecord themed;
  themed.name = "Arial";
  themed.scheme = FM_FONT_SCHEME_MINOR;
  styles.fonts.push_back(charset_only);
  styles.fonts.push_back(themed);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &loaded.handle), 0);
  size_t fonts = 0;
  ASSERT_EQ(fm_workbook_get_font_count(loaded.handle, &fonts), 0);
  fm_font_t partial{};
  ASSERT_EQ(fm_workbook_get_font_at(loaded.handle, static_cast<uint32_t>(fonts - 2), &partial), 0);
  EXPECT_EQ(partial.has_family, 0);
  EXPECT_EQ(partial.charset, 1);
  fm_font_t linked{};
  ASSERT_EQ(fm_workbook_get_font_at(loaded.handle, static_cast<uint32_t>(fonts - 1), &linked), 0);
  EXPECT_EQ(linked.has_family, 0);
  EXPECT_EQ(linked.has_charset, 0);
}

TEST(FormulonCApiStyles, AddFontWithoutFamilyOrCharsetIsCompletedAtCreation) {
  for (const auto& [profile, family] : {std::pair<const char*, int>{"mac-365-ja_JP", 3}, {"mac-365-en_US", 0}}) {
    SCOPED_TRACE(profile);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_for_profile(profile, &wb.handle), 0);
    fm_font_t record{};
    record.name = "游ゴシック";
    std::uint32_t index = 0;
    ASSERT_EQ(fm_workbook_add_font(wb.handle, &record, &index), 0);
    fm_font_t read{};
    ASSERT_EQ(fm_workbook_get_font_at(wb.handle, index, &read), 0);
    EXPECT_EQ(read.has_family != 0, family != 0);
    EXPECT_EQ(read.has_charset, 1);
    EXPECT_EQ(read.charset, 128);

    fm_font_t stated = record;
    stated.has_charset = 1;
    stated.charset = 5;
    ASSERT_EQ(fm_workbook_add_font(wb.handle, &stated, &index), 0);
    ASSERT_EQ(fm_workbook_get_font_at(wb.handle, index, &read), 0);
    EXPECT_EQ(read.has_family, 0);  // a stated charset suppresses the completion
    EXPECT_EQ(read.charset, 5);
  }
}
