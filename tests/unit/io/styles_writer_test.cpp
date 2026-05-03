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
}

}  // namespace
}  // namespace io
}  // namespace formulon
