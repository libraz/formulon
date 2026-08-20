//
// Unit tests for `formulon::io::write_styles`. The writer is the
// symmetric counterpart of `read_styles`; the integration round-trip
// suite exercises the full read/write/read pipeline. These tests pin
// the byte-level shape of the emitted document.

#include "io/styles_writer.h"

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/styles_reader.h"

namespace formulon {
namespace io {
namespace {

TEST(StylesWriter, EmitsMinimalStyleSheetForEmptyTable) {
  StylesTable table;
  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<styleSheet"), std::string::npos);
  EXPECT_NE(xml.find("<fonts count=\"1\""), std::string::npos);
  EXPECT_NE(xml.find("<fills count=\"1\""), std::string::npos);
  EXPECT_NE(xml.find("<borders count=\"1\""), std::string::npos);
  EXPECT_NE(xml.find("<cellXfs count=\"1\""), std::string::npos);
}

TEST(StylesWriter, NormalizesEmptyNamedStyleTablesWithoutMutatingModel) {
  StylesTable table;
  const std::string xml = write_styles(table);

  EXPECT_NE(xml.find("<cellStyleXfs count=\"1\">"), std::string::npos);
  EXPECT_NE(xml.find("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<cellStyles count=\"1\">"), std::string::npos);
  EXPECT_NE(xml.find("<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"), std::string::npos);
  EXPECT_TRUE(table.cell_style_xfs.empty());
  EXPECT_TRUE(table.cell_styles.empty());

  const std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto round_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(round_or)) << round_or.error().message;
  ASSERT_EQ(round_or.value().cell_style_xfs.size(), 1U);
  EXPECT_EQ(round_or.value().cell_style_xfs[0].font_index, 0U);
  ASSERT_EQ(round_or.value().cell_styles.size(), 1U);
  EXPECT_EQ(round_or.value().cell_styles[0].name, "Normal");
  EXPECT_EQ(round_or.value().cell_styles[0].xf_id, 0U);
  EXPECT_EQ(round_or.value().cell_styles[0].builtin_id, 0U);
}

TEST(StylesWriter, RetainsUnmodelledTopLevelStyleSections) {
  const std::string source =
      "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:x14=\"urn:test:x14\" mc:Ignorable=\"x14\" "
      "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
      "<colors><indexedColors><rgbColor rgb=\"FF112233\"/></indexedColors></colors>"
      "<tableStyles count=\"1\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>"
      "<futureStyles vendor=\"example\"><futureStyle id=\"7\"/></futureStyles>"
      "<extLst><ext uri=\"urn:test\"><x14:future/></ext></extLst>"
      "</styleSheet>";
  auto read_or = read_styles(std::vector<std::uint8_t>(source.begin(), source.end()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;

  const std::string xml = write_styles(read_or.value());
  EXPECT_NE(xml.find("xmlns:x14=\"urn:test:x14\""), std::string::npos);
  EXPECT_NE(xml.find("mc:Ignorable=\"x14\""), std::string::npos);
  EXPECT_NE(xml.find("<colors><indexedColors><rgbColor rgb=\"FF112233\"/></indexedColors></colors>"),
            std::string::npos);
  EXPECT_NE(xml.find("<tableStyles count=\"1\""), std::string::npos);
  EXPECT_NE(xml.find("<futureStyles vendor=\"example\"><futureStyle id=\"7\"/></futureStyles>"), std::string::npos);
  EXPECT_NE(xml.find("<extLst><ext uri=\"urn:test\"><x14:future/></ext></extLst>"), std::string::npos);

  auto round_or = read_styles(std::vector<std::uint8_t>(xml.begin(), xml.end()));
  ASSERT_TRUE(static_cast<bool>(round_or)) << round_or.error().message;
  EXPECT_EQ(round_or.value().colors_xml, read_or.value().colors_xml);
  EXPECT_EQ(round_or.value().table_styles_xml, read_or.value().table_styles_xml);
  EXPECT_EQ(round_or.value().ext_lst_xml, read_or.value().ext_lst_xml);
  EXPECT_EQ(round_or.value().unknown_top_level_xml, read_or.value().unknown_top_level_xml);
}

TEST(StylesWriter, EmitsExplicitFalseFontTogglesForDifferentialFormat) {
  StylesTable table;
  DifferentialFormat dxf;
  dxf.has_font = true;
  dxf.font.has_bold = true;
  dxf.font.has_italic = true;
  dxf.font.has_strike = true;
  table.dxfs.push_back(dxf);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<b val=\"0\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<i val=\"0\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<strike val=\"0\"/>"), std::string::npos);

  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto read_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(read_or));
  const FontRecord& font = read_or.value().dxfs[0].font;
  EXPECT_TRUE(font.has_bold);
  EXPECT_FALSE(font.bold);
  EXPECT_TRUE(font.has_italic);
  EXPECT_FALSE(font.italic);
  EXPECT_TRUE(font.has_strike);
  EXPECT_FALSE(font.strike);
}

TEST(StylesWriter, SkipsBuiltinNumFmts) {
  // Built-in ids 0..163 must NOT appear in `<numFmts>`. Excel rejects
  // packages that redeclare built-in ids.
  StylesTable table;
  NumFmtRecord builtin;
  builtin.id = 14;  // mm-dd-yy
  builtin.format_string_index = 0;
  table.num_fmt_strings.emplace_back("mm-dd-yy");
  table.num_fmts.push_back(builtin);

  const std::string xml = write_styles(table);
  EXPECT_EQ(xml.find("<numFmts"), std::string::npos);
}

TEST(StylesWriter, EmitsCustomNumFmts) {
  StylesTable table;
  NumFmtRecord rec;
  rec.id = 164;
  rec.format_string_index = 0;
  table.num_fmt_strings.emplace_back("0.0000");
  table.num_fmts.push_back(rec);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<numFmts count=\"1\""), std::string::npos);
  EXPECT_NE(xml.find("numFmtId=\"164\""), std::string::npos);
  EXPECT_NE(xml.find("formatCode=\"0.0000\""), std::string::npos);
}

TEST(StylesWriter, NumFmtCodeWithNewlineIsAttributeEscaped) {
  // `formatCode` is an attribute value; a literal embedded newline would
  // be normalised away by any conforming XML parser on reload. The
  // writer must emit a character reference instead of the raw byte.
  StylesTable table;
  NumFmtRecord rec;
  rec.id = 164;
  rec.format_string_index = 0;
  table.num_fmt_strings.emplace_back("0.00;\n-0.00");
  table.num_fmts.push_back(rec);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("formatCode=\"0.00;&#10;-0.00\""), std::string::npos) << xml;
}

TEST(StylesWriter, FontNameWithTabIsAttributeEscaped) {
  StylesTable table;
  FontRecord f;
  f.name = "Meiryo\tUI";
  table.fonts.push_back(f);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<name val=\"Meiryo&#9;UI\"/>"), std::string::npos) << xml;
}

TEST(StylesWriter, EmitsFontFields) {
  StylesTable table;
  FontRecord f;
  f.name = "Meiryo";
  f.size = 14.0;
  f.bold = true;
  f.italic = true;
  f.underline = 2;  // double
  f.color_argb = 0xFFFF0000U;
  table.fonts.push_back(f);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<b/>"), std::string::npos);
  EXPECT_NE(xml.find("<i/>"), std::string::npos);
  EXPECT_NE(xml.find("<u val=\"double\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<sz val=\"14\""), std::string::npos);
  EXPECT_NE(xml.find("rgb=\"FFFF0000\""), std::string::npos);
  EXPECT_NE(xml.find("<name val=\"Meiryo\""), std::string::npos);
}

TEST(StylesWriter, FontSizeAndTintUseTheShortestRoundTripSpelling) {
  // Format metrics share the cell path's number spelling. Neither 10.5 nor
  // -0.25 exposes the difference (both are exact binary fractions), so this
  // uses the values that do: a fixed 17-significant-digit format would spell
  // them 8.0999999999999996 and 0.34999999999999998.
  StylesTable table;
  FontRecord f;
  f.size = 8.1;
  f.color.kind = ColorSpec::Kind::kTheme;
  f.color.theme = 3U;
  f.color.tint = 0.35;
  table.fonts.push_back(f);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<sz val=\"8.1\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("tint=\"0.35\""), std::string::npos) << xml;
}

TEST(StylesWriter, EmitsFillsAndBorders) {
  StylesTable table;
  FillRecord fill;
  fill.pattern = 1;  // solid
  fill.fg_argb = 0xFF00FF00U;
  table.fills.push_back(fill);
  BorderRecord b;
  b.left.style = 1;
  b.left.color_argb = 0xFF000000U;
  b.right.style = 2;
  table.borders.push_back(b);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("patternType=\"solid\""), std::string::npos);
  EXPECT_NE(xml.find("<fgColor rgb=\"FF00FF00\""), std::string::npos);
  EXPECT_NE(xml.find("<left style=\"thin\""), std::string::npos);
  EXPECT_NE(xml.find("<right style=\"medium\""), std::string::npos);
}

TEST(StylesWriter, EmitsCellXfAlignment) {
  StylesTable table;
  // Two fonts so `fontId="1"` names a record the same part emits; the
  // writer replaces any id it cannot resolve with the default.
  table.fonts.resize(2);
  CellXf xf;
  xf.num_fmt_id = 164;
  xf.font_index = 1;
  xf.horizontal_align = 2;  // center
  xf.vertical_align = 1;    // center
  xf.wrap_text = true;
  table.cell_xfs.push_back(xf);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("numFmtId=\"164\""), std::string::npos);
  EXPECT_NE(xml.find("fontId=\"1\""), std::string::npos);
  EXPECT_NE(xml.find("horizontal=\"center\""), std::string::npos);
  EXPECT_NE(xml.find("vertical=\"center\""), std::string::npos);
  EXPECT_NE(xml.find("wrapText=\"1\""), std::string::npos);
}

TEST(StylesWriter, EmitsExplicitEmptyAlignment) {
  StylesTable table;
  table.cell_xfs.emplace_back();
  table.cell_xfs[0].has_alignment = true;
  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<alignment/>"), std::string::npos);
  auto round_or = read_styles(std::vector<std::uint8_t>(xml.begin(), xml.end()));
  ASSERT_TRUE(static_cast<bool>(round_or)) << round_or.error().message;
  ASSERT_EQ(round_or.value().cell_xfs.size(), 1U);
  EXPECT_TRUE(round_or.value().cell_xfs[0].has_alignment);
}

TEST(StylesWriter, EmitsExplicitAlignmentDefaults) {
  StylesTable table;
  CellXf xf;
  xf.has_horizontal_align = true;
  xf.has_vertical_align = true;
  xf.has_wrap_text = true;
  xf.has_justify_last_line = true;
  table.cell_xfs.push_back(xf);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("horizontal=\"general\""), std::string::npos);
  EXPECT_NE(xml.find("vertical=\"bottom\""), std::string::npos);
  EXPECT_NE(xml.find("wrapText=\"0\""), std::string::npos);
  EXPECT_NE(xml.find("justifyLastLine=\"0\""), std::string::npos);
  auto round_or = read_styles(std::vector<std::uint8_t>(xml.begin(), xml.end()));
  ASSERT_TRUE(static_cast<bool>(round_or)) << round_or.error().message;
  const CellXf& round = round_or.value().cell_xfs[0];
  EXPECT_TRUE(round.has_horizontal_align);
  EXPECT_TRUE(round.has_vertical_align);
  EXPECT_TRUE(round.has_wrap_text);
  EXPECT_TRUE(round.has_justify_last_line);
}

TEST(StylesWriter, NormalizesDanglingXfIdsWithoutMutatingModel) {
  StylesTable table;
  table.cell_style_xfs.emplace_back();
  CellXf cell_xf;
  cell_xf.xf_id = 99U;
  table.cell_xfs.push_back(cell_xf);
  CellStyleRecord style;
  style.name = "Custom";
  style.xf_id = 99U;
  table.cell_styles.push_back(style);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<cellStyle name=\"Custom\" xfId=\"0\""), std::string::npos);
  EXPECT_EQ(table.cell_xfs[0].xf_id, 99U);
  EXPECT_EQ(table.cell_styles[0].xf_id, 99U);
}

TEST(StylesWriter, RoundTripsOptionalAlignmentAttributesIncludingExplicitDefaults) {
  StylesTable original;
  CellXf absent;
  original.cell_xfs.push_back(absent);
  CellXf xf;
  xf.has_text_rotation = true;
  xf.text_rotation = 0;
  xf.has_indent = true;
  xf.indent = 0;
  xf.has_relative_indent = true;
  xf.relative_indent = -4;
  xf.has_shrink_to_fit = true;
  xf.shrink_to_fit = false;
  xf.has_reading_order = true;
  xf.reading_order = 0;
  original.cell_xfs.push_back(xf);

  const std::string xml = write_styles(original);
  EXPECT_NE(xml.find("textRotation=\"0\""), std::string::npos);
  EXPECT_NE(xml.find("indent=\"0\""), std::string::npos);
  EXPECT_NE(xml.find("relativeIndent=\"-4\""), std::string::npos);
  EXPECT_NE(xml.find("shrinkToFit=\"0\""), std::string::npos);
  EXPECT_NE(xml.find("readingOrder=\"0\""), std::string::npos);

  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto round_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(round_or)) << round_or.error().message;
  ASSERT_EQ(round_or.value().cell_xfs.size(), 2U);
  EXPECT_FALSE(round_or.value().cell_xfs[0].has_text_rotation);
  const CellXf& round = round_or.value().cell_xfs[1];
  EXPECT_TRUE(round.has_text_rotation);
  EXPECT_EQ(round.text_rotation, 0U);
  EXPECT_TRUE(round.has_indent);
  EXPECT_EQ(round.indent, 0U);
  EXPECT_TRUE(round.has_relative_indent);
  EXPECT_EQ(round.relative_indent, -4);
  EXPECT_TRUE(round.has_shrink_to_fit);
  EXPECT_FALSE(round.shrink_to_fit);
  EXPECT_TRUE(round.has_reading_order);
  EXPECT_EQ(round.reading_order, 0U);
}

TEST(StylesWriter, RoundTripsThroughReader) {
  StylesTable original;
  FontRecord font;
  font.name = "Arial";
  font.size = 12.0;
  font.bold = true;
  font.color_argb = 0xFF112233U;
  original.fonts.push_back(font);

  FillRecord fill;
  fill.pattern = 1;
  fill.fg_argb = 0xFFAABBCCU;
  original.fills.push_back(fill);

  original.borders.emplace_back();
  original.num_fmt_strings.emplace_back("#,##0.00");
  NumFmtRecord nf;
  nf.id = 200;
  nf.format_string_index = 0;
  original.num_fmts.push_back(nf);

  CellXf xf;
  xf.num_fmt_id = 200;
  xf.font_index = 0;
  xf.fill_index = 0;
  xf.horizontal_align = 1;  // left
  original.cell_xfs.push_back(xf);

  DifferentialFormat dxf;
  dxf.has_font = true;
  dxf.font.bold = true;
  dxf.font.color_argb = 0xFFFF0000U;
  dxf.has_fill = true;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFFFFFF00U;
  dxf.has_num_fmt = true;
  dxf.num_fmt_id = 201;
  dxf.num_fmt_code = "0.0";
  original.dxfs.push_back(dxf);

  const std::string xml = write_styles(original);
  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto round_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(round_or)) << "read failed: " << round_or.error().message;
  const StylesTable& rt = round_or.value();
  ASSERT_EQ(rt.fonts.size(), 1U);
  EXPECT_EQ(rt.fonts[0].name, "Arial");
  EXPECT_TRUE(rt.fonts[0].bold);
  EXPECT_EQ(rt.fonts[0].color_argb, 0xFF112233U);
  ASSERT_EQ(rt.fills.size(), 1U);
  EXPECT_EQ(rt.fills[0].pattern, 1U);
  EXPECT_EQ(rt.fills[0].fg_argb, 0xFFAABBCCU);
  ASSERT_EQ(rt.num_fmts.size(), 1U);
  EXPECT_EQ(rt.num_fmts[0].id, 200U);
  EXPECT_EQ(rt.num_fmt_strings[rt.num_fmts[0].format_string_index], "#,##0.00");
  ASSERT_EQ(rt.cell_xfs.size(), 1U);
  EXPECT_EQ(rt.cell_xfs[0].num_fmt_id, 200U);
  EXPECT_EQ(rt.cell_xfs[0].horizontal_align, 1U);
  ASSERT_EQ(rt.dxfs.size(), 1U);
  EXPECT_TRUE(rt.dxfs[0].has_font);
  EXPECT_TRUE(rt.dxfs[0].font.bold);
  EXPECT_EQ(rt.dxfs[0].font.color_argb, 0xFFFF0000U);
  EXPECT_TRUE(rt.dxfs[0].has_fill);
  EXPECT_EQ(rt.dxfs[0].fill.fg_argb, 0xFFFFFF00U);
  EXPECT_TRUE(rt.dxfs[0].has_num_fmt);
  EXPECT_EQ(rt.dxfs[0].num_fmt_code, "0.0");
}

TEST(StylesWriter, EmitsCellXfProtection) {
  StylesTable table;
  CellXf unlocked;
  unlocked.has_protection = true;
  unlocked.locked = false;
  table.cell_xfs.push_back(unlocked);

  const std::string xml = write_styles(table);
  EXPECT_NE(xml.find("<protection locked=\"0\" hidden=\"0\"/>"), std::string::npos);
}

TEST(StylesWriter, OmitsProtectionWhenAbsent) {
  // An xf that never carried a <protection> element must not gain one on
  // write; otherwise a plain cell would round-trip to explicitly locked
  // and defeat inheritance from the cell style.
  StylesTable table;
  table.cell_xfs.emplace_back();

  const std::string xml = write_styles(table);
  EXPECT_EQ(xml.find("<protection"), std::string::npos);
}

TEST(StylesWriter, RoundTripsProtectionState) {
  StylesTable original;
  CellXf locked_default;  // no protection element
  original.cell_xfs.push_back(locked_default);
  CellXf unlocked;
  unlocked.has_protection = true;
  unlocked.locked = false;
  unlocked.hidden = true;
  original.cell_xfs.push_back(unlocked);

  const std::string xml = write_styles(original);
  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto round_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(round_or)) << "read failed: " << round_or.error().message;
  const StylesTable& rt = round_or.value();
  ASSERT_EQ(rt.cell_xfs.size(), 2U);
  EXPECT_FALSE(rt.cell_xfs[0].has_protection);
  EXPECT_TRUE(rt.cell_xfs[0].locked);  // schema default
  EXPECT_TRUE(rt.cell_xfs[1].has_protection);
  EXPECT_FALSE(rt.cell_xfs[1].locked);
  EXPECT_TRUE(rt.cell_xfs[1].hidden);
}

TEST(StylesWriter, RoundTripsThemeAndIndexedColors) {
  StylesTable original;
  FontRecord font;
  font.name = "Calibri";
  font.color.kind = ColorSpec::Kind::kTheme;
  font.color.theme = 1;
  font.color.tint = -0.25;
  original.fonts.push_back(font);

  FillRecord fill;
  fill.pattern = 1;
  fill.fg.kind = ColorSpec::Kind::kIndexed;
  fill.fg.indexed = 64;
  fill.bg.kind = ColorSpec::Kind::kAuto;
  original.fills.push_back(fill);

  BorderRecord border;
  border.left.style = 1;
  border.left.color.kind = ColorSpec::Kind::kTheme;
  border.left.color.theme = 4;
  original.borders.push_back(border);

  const std::string xml = write_styles(original);
  EXPECT_NE(xml.find("<color theme=\"1\" tint=\"-0.25\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<fgColor indexed=\"64\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<bgColor auto=\"1\"/>"), std::string::npos);

  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto round_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(round_or)) << "read failed: " << round_or.error().message;
  const StylesTable& rt = round_or.value();
  ASSERT_EQ(rt.fonts.size(), 1U);
  EXPECT_EQ(rt.fonts[0].color.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(rt.fonts[0].color.theme, 1U);
  EXPECT_DOUBLE_EQ(rt.fonts[0].color.tint, -0.25);
  ASSERT_EQ(rt.fills.size(), 1U);
  EXPECT_EQ(rt.fills[0].fg.kind, ColorSpec::Kind::kIndexed);
  EXPECT_EQ(rt.fills[0].fg.indexed, 64U);
  EXPECT_EQ(rt.fills[0].bg.kind, ColorSpec::Kind::kAuto);
  ASSERT_EQ(rt.borders.size(), 1U);
  EXPECT_EQ(rt.borders[0].left.color.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(rt.borders[0].left.color.theme, 4U);
}

TEST(StylesWriter, RoundTripsFontVertAlignFamilyCharset) {
  StylesTable original;
  FontRecord font;
  font.name = "\xEF\xBC\xAD\xEF\xBC\xB3 \xE6\x98\x8E\xE6\x9C\x9D";  // "ＭＳ 明朝"
  font.vert_align = 1;                                              // superscript
  font.has_family = true;
  font.family = 1;
  font.has_charset = true;
  font.charset = 128;  // Shift_JIS
  original.fonts.push_back(font);

  const std::string xml = write_styles(original);
  EXPECT_NE(xml.find("<vertAlign val=\"superscript\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<family val=\"1\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<charset val=\"128\"/>"), std::string::npos);

  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto round_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(round_or)) << "read failed: " << round_or.error().message;
  const StylesTable& rt = round_or.value();
  ASSERT_EQ(rt.fonts.size(), 1U);
  EXPECT_EQ(rt.fonts[0].vert_align, 1U);
  EXPECT_TRUE(rt.fonts[0].has_family);
  EXPECT_EQ(rt.fonts[0].family, 1U);
  EXPECT_TRUE(rt.fonts[0].has_charset);
  EXPECT_EQ(rt.fonts[0].charset, 128U);
}

TEST(StylesWriter, RoundTripsDxfThemeColor) {
  StylesTable original;
  DifferentialFormat dxf;
  dxf.has_font = true;
  dxf.font.color.kind = ColorSpec::Kind::kTheme;
  dxf.font.color.theme = 5;
  original.dxfs.push_back(dxf);

  const std::string xml = write_styles(original);
  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto round_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(round_or)) << "read failed: " << round_or.error().message;
  const StylesTable& rt = round_or.value();
  ASSERT_EQ(rt.dxfs.size(), 1U);
  ASSERT_TRUE(rt.dxfs[0].has_font);
  EXPECT_EQ(rt.dxfs[0].font.color.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(rt.dxfs[0].font.color.theme, 5U);
}

TEST(StylesWriter, RoundTripsDxfAlignmentAndProtection) {
  // A dxf carrying <alignment> and <protection> (rare, but valid) must
  // survive verbatim, coexisting with a theme-coloured font.
  StylesTable original;
  DifferentialFormat dxf;
  dxf.has_font = true;
  dxf.font.color.kind = ColorSpec::Kind::kTheme;
  dxf.font.color.theme = 4;
  dxf.alignment_xml = "<alignment horizontal=\"center\" wrapText=\"1\"/>";
  dxf.protection_xml = "<protection locked=\"0\"/>";
  original.dxfs.push_back(dxf);

  const std::string xml = write_styles(original);
  // CT_Dxf order: alignment after fill (none here), protection after border.
  const std::size_t p_align = xml.find("<alignment");
  const std::size_t p_prot = xml.find("<protection");
  ASSERT_NE(p_align, std::string::npos) << xml;
  ASSERT_NE(p_prot, std::string::npos) << xml;
  EXPECT_LT(p_align, p_prot) << xml;

  std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  auto round_or = read_styles(bytes);
  ASSERT_TRUE(static_cast<bool>(round_or)) << "read failed: " << round_or.error().message;
  const StylesTable& rt = round_or.value();
  ASSERT_EQ(rt.dxfs.size(), 1U);
  EXPECT_NE(rt.dxfs[0].alignment_xml.find("horizontal=\"center\""), std::string::npos);
  EXPECT_NE(rt.dxfs[0].alignment_xml.find("wrapText=\"1\""), std::string::npos);
  EXPECT_NE(rt.dxfs[0].protection_xml.find("locked=\"0\""), std::string::npos);
  // Coexists with the ColorSpec theme-colour capture.
  EXPECT_EQ(rt.dxfs[0].font.color.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(rt.dxfs[0].font.color.theme, 4U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
