// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
