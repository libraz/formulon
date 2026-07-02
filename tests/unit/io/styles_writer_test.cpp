// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
