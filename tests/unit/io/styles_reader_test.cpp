//
// Unit tests for `formulon::io::read_styles`. The reader builds a flat
// `StylesTable` carrying every record kind needed for round-trip
// (`fonts`, `fills`, `borders`, `num_fmts`, `cell_xfs`) plus an
// interned vector of custom number-format strings.

#include "io/styles_reader.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "utils/error.h"

namespace formulon {
namespace io {
namespace {

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

TEST(StylesReader, EmptyStyleSheetSelfClosing) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read failed: " << result_or.error().message;
  // Default tables guarantee `xf_index = 0` always resolves: each list
  // carries at least one default entry even when the source document
  // has no children.
  EXPECT_EQ(result_or.value().fonts.size(), 1U);
  EXPECT_EQ(result_or.value().fills.size(), 1U);
  EXPECT_EQ(result_or.value().borders.size(), 1U);
  EXPECT_EQ(result_or.value().cell_xfs.size(), 1U);
  EXPECT_TRUE(result_or.value().num_fmts.empty());
}

TEST(StylesReader, EmptyStyleSheetWithChildren) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>\n");
  xml.append("  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>\n");
  xml.append("  <borders count=\"1\"><border/></borders>\n");
  xml.append(
      "  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n");
  xml.append("  <cellXfs count=\"0\"></cellXfs>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  // Empty `cellXfs` falls back to one default `<xf>` so the index
  // space stays self-consistent.
  EXPECT_EQ(result_or.value().cell_xfs.size(), 1U);
  EXPECT_TRUE(result_or.value().num_fmts.empty());
}

TEST(StylesReader, ReadsCellXfFields) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <cellXfs count=\"3\">\n");
  xml.append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
  xml.append("    <xf numFmtId=\"14\" fontId=\"1\" fillId=\"2\" borderId=\"0\"/>\n");
  xml.append("    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\">");
  xml.append("<alignment horizontal=\"center\" vertical=\"top\" wrapText=\"1\"/></xf>\n");
  xml.append("  </cellXfs>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().cell_xfs.size(), 3U);
  EXPECT_EQ(result_or.value().cell_xfs[0].num_fmt_id, 0U);
  EXPECT_EQ(result_or.value().cell_xfs[1].num_fmt_id, 14U);
  EXPECT_EQ(result_or.value().cell_xfs[1].font_index, 1U);
  EXPECT_EQ(result_or.value().cell_xfs[1].fill_index, 2U);
  EXPECT_EQ(result_or.value().cell_xfs[2].horizontal_align, 2U);  // center
  EXPECT_EQ(result_or.value().cell_xfs[2].vertical_align, 0U);    // top
  EXPECT_TRUE(result_or.value().cell_xfs[2].wrap_text);
  EXPECT_TRUE(result_or.value().cell_xfs[2].has_horizontal_align);
  EXPECT_TRUE(result_or.value().cell_xfs[2].has_vertical_align);
  EXPECT_TRUE(result_or.value().cell_xfs[2].has_wrap_text);
}

TEST(StylesReader, CollapsesXmlWhitespaceAndAcceptsSignedLexicalForms) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<cellXfs count=\"1\"><xf fontId=\" +1 \" fillId=\"0\" borderId=\"0\" numFmtId=\" +14 \" ");
  xml.append("applyFont=\" true \" quotePrefix=\" 0 \">");
  xml.append("<alignment horizontal=\"center\" vertical=\"bottom\" wrapText=\" 0 \" ");
  xml.append("justifyLastLine=\" false \" textRotation=\" +255 \" indent=\" +7 \" ");
  xml.append("relativeIndent=\" +3 \" shrinkToFit=\" 0 \" readingOrder=\" +2 \"/></xf>");
  xml.append("</cellXfs></styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const CellXf& xf = result_or.value().cell_xfs[0];
  EXPECT_EQ(xf.font_index, 1U);
  EXPECT_EQ(xf.num_fmt_id, 14U);
  EXPECT_TRUE(xf.apply_font);
  EXPECT_FALSE(xf.quote_prefix);
  EXPECT_EQ(xf.horizontal_align, 2U);
  EXPECT_EQ(xf.vertical_align, 2U);
  EXPECT_FALSE(xf.wrap_text);
  EXPECT_FALSE(xf.justify_last_line);
  EXPECT_TRUE(xf.has_horizontal_align);
  EXPECT_TRUE(xf.has_vertical_align);
  EXPECT_TRUE(xf.has_wrap_text);
  EXPECT_TRUE(xf.has_justify_last_line);
  EXPECT_EQ(xf.text_rotation, 255U);
  EXPECT_EQ(xf.indent, 7U);
  EXPECT_EQ(xf.relative_indent, 3);
  EXPECT_FALSE(xf.shrink_to_fit);
  EXPECT_EQ(xf.reading_order, 2U);
}

TEST(StylesReader, ReadsCellXfProtection) {
  // Excel-shaped input: default xf (no protection) plus an unlocked
  // input-cell xf. The reader must distinguish "no element" from
  // "explicit locked=0".
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <cellXfs count=\"2\">\n");
  xml.append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
  xml.append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyProtection=\"1\">");
  xml.append("<protection locked=\"0\" hidden=\"1\"/></xf>\n");
  xml.append("  </cellXfs>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read failed: " << result_or.error().message;
  const StylesTable& table = result_or.value();
  ASSERT_EQ(table.cell_xfs.size(), 2U);
  EXPECT_FALSE(table.cell_xfs[0].has_protection);
  EXPECT_TRUE(table.cell_xfs[0].locked);  // schema default
  EXPECT_FALSE(table.cell_xfs[0].hidden);
  EXPECT_TRUE(table.cell_xfs[1].has_protection);
  EXPECT_FALSE(table.cell_xfs[1].locked);
  EXPECT_TRUE(table.cell_xfs[1].hidden);
}

TEST(StylesReader, ReadsOptionalAlignmentAttributesWithPresence) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <cellXfs count=\"3\">\n");
  xml.append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
  xml.append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\">");
  xml.append("<alignment textRotation=\"255\" indent=\"7\" relativeIndent=\"-3\" ");
  xml.append("shrinkToFit=\"0\" readingOrder=\"2\"/></xf>\n");
  xml.append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\">");
  xml.append("<alignment textRotation=\"0\" indent=\"0\" relativeIndent=\"0\" ");
  xml.append("shrinkToFit=\"1\" readingOrder=\"0\"/></xf>\n");
  xml.append("  </cellXfs>\n</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const auto& xfs = result_or.value().cell_xfs;
  ASSERT_EQ(xfs.size(), 3U);
  EXPECT_FALSE(xfs[0].has_text_rotation);
  EXPECT_FALSE(xfs[0].has_indent);
  EXPECT_FALSE(xfs[0].has_relative_indent);
  EXPECT_FALSE(xfs[0].has_shrink_to_fit);
  EXPECT_FALSE(xfs[0].has_reading_order);
  EXPECT_TRUE(xfs[1].has_text_rotation);
  EXPECT_EQ(xfs[1].text_rotation, 255U);
  EXPECT_TRUE(xfs[1].has_indent);
  EXPECT_EQ(xfs[1].indent, 7U);
  EXPECT_TRUE(xfs[1].has_relative_indent);
  EXPECT_EQ(xfs[1].relative_indent, -3);
  EXPECT_TRUE(xfs[1].has_shrink_to_fit);
  EXPECT_FALSE(xfs[1].shrink_to_fit);
  EXPECT_TRUE(xfs[1].has_reading_order);
  EXPECT_EQ(xfs[1].reading_order, 2U);
  EXPECT_TRUE(xfs[2].has_text_rotation);
  EXPECT_EQ(xfs[2].text_rotation, 0U);
  EXPECT_TRUE(xfs[2].has_shrink_to_fit);
  EXPECT_TRUE(xfs[2].shrink_to_fit);
}

TEST(StylesReader, PreservesEmptyAlignmentPresence) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<cellXfs count=\"2\"><xf/><xf><alignment/></xf></cellXfs></styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  ASSERT_EQ(result_or.value().cell_xfs.size(), 2U);
  EXPECT_FALSE(result_or.value().cell_xfs[0].has_alignment);
  EXPECT_TRUE(result_or.value().cell_xfs[1].has_alignment);
  EXPECT_EQ(result_or.value().cell_xfs[1].vertical_align, 2U);
}

TEST(StylesReader, RejectsMalformedAlignmentAttributesWithContext) {
  const std::vector<std::string> alignments = {
      "<alignment textRotation=\"181\"/>",
      "<alignment textRotation=\"wat\"/>",
      "<alignment indent=\"256\"/>",
      "<alignment readingOrder=\"3\"/>",
      "<alignment relativeIndent=\"2147483648\"/>",
      "<alignment wrapText=\"yes\"/>",
      "<alignment wrapText=\"tr ue\"/>",
      "<alignment horizontal=\"sideways\"/>",
      "<alignment horizontal=\" center \"/>",
      "<alignment vertical=\" bottom \"/>",
      "<alignment horizontal=\" center nope \"/>",
      "<alignment textRotation=\"-1\"/>",
      "<alignment indent=\"-1\"/>",
      "<alignment readingOrder=\"4294967296\"/>",
      "<alignment readingOrder=\"+\"/>",
  };
  for (const std::string& alignment : alignments) {
    std::string xml(kXmlDecl);
    xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
    xml.append("<cellXfs count=\"1\"><xf>");
    xml.append(alignment);
    xml.append("</xf></cellXfs></styleSheet>");
    auto result_or = read_styles(Bytes(xml));
    ASSERT_FALSE(static_cast<bool>(result_or)) << alignment;
    EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoSheetCorrupt) << alignment;
    EXPECT_NE(result_or.error().context.find("section=cellXfs"), std::string::npos) << alignment;
    EXPECT_NE(result_or.error().context.find("index=0"), std::string::npos) << alignment;
  }
}

TEST(StylesReader, ReadsThemeIndexedAndAutoColors) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <fonts count=\"1\"><font><sz val=\"11\"/><color theme=\"3\" tint=\"0.5\"/>");
  xml.append("<name val=\"Calibri\"/></font></fonts>\n");
  xml.append("  <fills count=\"1\"><fill><patternFill patternType=\"solid\">");
  xml.append("<fgColor indexed=\"9\"/><bgColor auto=\"1\"/></patternFill></fill></fills>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read failed: " << result_or.error().message;
  const StylesTable& table = result_or.value();
  ASSERT_EQ(table.fonts.size(), 1U);
  EXPECT_EQ(table.fonts[0].color.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(table.fonts[0].color.theme, 3U);
  EXPECT_DOUBLE_EQ(table.fonts[0].color.tint, 0.5);
  ASSERT_EQ(table.fills.size(), 1U);
  EXPECT_EQ(table.fills[0].fg.kind, ColorSpec::Kind::kIndexed);
  EXPECT_EQ(table.fills[0].fg.indexed, 9U);
  EXPECT_EQ(table.fills[0].bg.kind, ColorSpec::Kind::kAuto);
}

TEST(StylesReader, ReadsFontVertAlignFamilyCharset) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <fonts count=\"1\"><font><vertAlign val=\"subscript\"/><sz val=\"11\"/>");
  xml.append("<name val=\"MS Gothic\"/><family val=\"3\"/><charset val=\"128\"/></font></fonts>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read failed: " << result_or.error().message;
  const StylesTable& table = result_or.value();
  ASSERT_EQ(table.fonts.size(), 1U);
  EXPECT_EQ(table.fonts[0].vert_align, 2U);  // subscript
  EXPECT_TRUE(table.fonts[0].has_family);
  EXPECT_EQ(table.fonts[0].family, 3U);
  EXPECT_TRUE(table.fonts[0].has_charset);
  EXPECT_EQ(table.fonts[0].charset, 128U);
}

TEST(StylesReader, InternsCustomNumFmts) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <numFmts count=\"2\">\n");
  xml.append("    <numFmt numFmtId=\"164\" formatCode=\"0.0000\"/>\n");
  xml.append("    <numFmt numFmtId=\"165\" formatCode=\"yyyy/mm/dd\"/>\n");
  xml.append("  </numFmts>\n");
  xml.append("  <cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().num_fmts.size(), 2U);
  EXPECT_EQ(result_or.value().num_fmts[0].id, 164U);
  EXPECT_EQ(result_or.value().num_fmts[1].id, 165U);
  ASSERT_EQ(result_or.value().num_fmt_strings.size(), 2U);
  EXPECT_EQ(result_or.value().num_fmt_strings[result_or.value().num_fmts[0].format_string_index], "0.0000");
  EXPECT_EQ(result_or.value().num_fmt_strings[result_or.value().num_fmts[1].format_string_index], "yyyy/mm/dd");
}

TEST(StylesReader, ReadsFontFields) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <fonts count=\"2\">\n");
  xml.append("    <font><sz val=\"11\"/><name val=\"Calibri\"/></font>\n");
  xml.append("    <font><b/><i/><u val=\"double\"/><sz val=\"14\"/>");
  xml.append("<color rgb=\"FFFF0000\"/><name val=\"Meiryo\"/></font>\n");
  xml.append("  </fonts>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().fonts.size(), 2U);
  EXPECT_EQ(result_or.value().fonts[0].name, "Calibri");
  EXPECT_FLOAT_EQ(static_cast<float>(result_or.value().fonts[0].size), 11.0F);
  EXPECT_FALSE(result_or.value().fonts[0].bold);
  EXPECT_TRUE(result_or.value().fonts[1].bold);
  EXPECT_TRUE(result_or.value().fonts[1].italic);
  EXPECT_EQ(result_or.value().fonts[1].underline, 2U);
  EXPECT_EQ(result_or.value().fonts[1].name, "Meiryo");
  EXPECT_EQ(result_or.value().fonts[1].color_argb, 0xFFFF0000U);
}

TEST(StylesReader, ReadsFillsAndBorders) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <fills count=\"2\">\n");
  xml.append("    <fill><patternFill patternType=\"none\"/></fill>\n");
  xml.append("    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF00FF00\"/></patternFill></fill>\n");
  xml.append("  </fills>\n");
  xml.append("  <borders count=\"1\">\n");
  xml.append("    <border>");
  xml.append("<left style=\"thin\"><color rgb=\"FF000000\"/></left>");
  xml.append("<right style=\"medium\"/>");
  xml.append("<top/><bottom/><diagonal/>");
  xml.append("</border>\n");
  xml.append("  </borders>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().fills.size(), 2U);
  EXPECT_EQ(result_or.value().fills[1].pattern, 1U);  // solid
  EXPECT_EQ(result_or.value().fills[1].fg_argb, 0xFF00FF00U);
  ASSERT_EQ(result_or.value().borders.size(), 1U);
  EXPECT_EQ(result_or.value().borders[0].left.style, 1U);   // thin
  EXPECT_EQ(result_or.value().borders[0].right.style, 2U);  // medium
  EXPECT_EQ(result_or.value().borders[0].left.color_argb, 0xFF000000U);
}

TEST(StylesReader, RejectsMissingStyleSheetRoot) {
  std::string xml(kXmlDecl);
  xml.append("<notAStyleSheet/>");
  auto result_or = read_styles(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(StylesReader, RejectsMalformedXml) {
  std::string xml = "<styleSheet><cellXfs>";  // unterminated
  auto result_or = read_styles(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(StylesReader, IgnoresUnknownChildren) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs>\n");
  xml.append("  <dxfs count=\"0\"/>\n");
  xml.append("  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\"/>\n");
  xml.append("  <extLst><ext uri=\"{...}\"/></extLst>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().cell_xfs.size(), 1U);
  EXPECT_TRUE(result_or.value().num_fmts.empty());
}

TEST(StylesReader, ReadsDifferentialFormats) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <dxfs count=\"1\"><dxf>");
  xml.append("<font><b/><color rgb=\"FFFF0000\"/></font>");
  xml.append("<numFmt numFmtId=\"164\" formatCode=\"0.00\"/>");
  xml.append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFF00\"/></patternFill></fill>");
  xml.append("<border><left style=\"thin\"><color rgb=\"FF000000\"/></left></border>");
  xml.append("</dxf></dxfs>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().dxfs.size(), 1U);
  const DifferentialFormat& dxf = result_or.value().dxfs[0];
  ASSERT_TRUE(dxf.has_font);
  EXPECT_TRUE(dxf.font.bold);
  EXPECT_EQ(dxf.font.color_argb, 0xFFFF0000U);
  ASSERT_TRUE(dxf.has_num_fmt);
  EXPECT_EQ(dxf.num_fmt_id, 164U);
  EXPECT_EQ(dxf.num_fmt_code, "0.00");
  ASSERT_TRUE(dxf.has_fill);
  EXPECT_EQ(dxf.fill.fg_argb, 0xFFFFFF00U);
  ASSERT_TRUE(dxf.has_border);
  EXPECT_EQ(dxf.border.left.style, 1U);
}

TEST(StylesReader, PreservesExplicitFalseFontTogglesInDifferentialFormat) {
  const std::string xml =
      "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<dxfs count=\"1\"><dxf><font><b val=\"0\"/><i val=\"0\"/><strike val=\"0\"/></font></dxf></dxfs>"
      "</styleSheet>";
  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().dxfs.size(), 1U);
  const FontRecord& font = result_or.value().dxfs[0].font;
  EXPECT_TRUE(font.has_bold);
  EXPECT_FALSE(font.bold);
  EXPECT_TRUE(font.has_italic);
  EXPECT_FALSE(font.italic);
  EXPECT_TRUE(font.has_strike);
  EXPECT_FALSE(font.strike);
}

TEST(BuiltinNumFmt, ResolvesKnownIds) {
  // Built-in 0 = "General"; 14 = "mm-dd-yy"; 49 = "@".
  EXPECT_STREQ(builtin_num_fmt(0), "General");
  EXPECT_STREQ(builtin_num_fmt(14), "mm-dd-yy");
  EXPECT_STREQ(builtin_num_fmt(49), "@");
  // Reserved-but-undocumented slots return the empty string.
  EXPECT_STREQ(builtin_num_fmt(5), "");
  // Above the documented range, returns empty string.
  EXPECT_STREQ(builtin_num_fmt(200), "");
}

}  // namespace
}  // namespace io
}  // namespace formulon
