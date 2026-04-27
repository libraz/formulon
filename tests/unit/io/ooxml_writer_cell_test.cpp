// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the OOXML cell/row/sheetData builder. These exercise the
// pure helper surface (`EncodeA1`, `BuildSheetDataXml`) directly and pin
// the substring-level shape of the emitted markup. Higher-level integration
// of the builder with the zip orchestration is covered by
// tests/integration/ooxml_minimal_writer_test.cpp.

#include "io/ooxml_writer_cell.h"

#include <cmath>
#include <cstdint>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"

namespace formulon {
namespace io {
namespace {

// ---------------------------------------------------------------------------
// EncodeA1
// ---------------------------------------------------------------------------

TEST(EncodeA1, FirstCell) {
  EXPECT_EQ(EncodeA1(0U, 0U), "A1");
}

TEST(EncodeA1, ColumnZ) {
  EXPECT_EQ(EncodeA1(0U, 25U), "Z1");
}

TEST(EncodeA1, ColumnAA) {
  EXPECT_EQ(EncodeA1(0U, 26U), "AA1");
}

TEST(EncodeA1, ColumnXFD) {
  EXPECT_EQ(EncodeA1(0U, 16383U), "XFD1");
}

TEST(EncodeA1, MaxCell) {
  EXPECT_EQ(EncodeA1(1048575U, 16383U), "XFD1048576");
}

// ---------------------------------------------------------------------------
// BuildSheetDataXml: scalar cell shapes
// ---------------------------------------------------------------------------

// An empty sheet collapses to the self-closing form. Documented choice;
// integration tests pin this so a future `<sheetData></sheetData>` switch
// would be a deliberate, intentional change.
TEST(BuildSheetDataXml, EmptySheet) {
  Sheet s("Sheet1");
  EXPECT_EQ(BuildSheetDataXml(s), "<sheetData/>");
}

TEST(BuildSheetDataXml, NumberCell) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(42.5));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("<c r=\"A1\"><v>42.5</v></c>"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, IntegerNumberCell) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(42.0));
  const std::string xml = BuildSheetDataXml(s);
  // %.17g elides the trailing ".0" for whole-number doubles, so the wire
  // form is "42" not "42.0".
  EXPECT_NE(xml.find("<v>42</v>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<v>42.0</v>"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, BoolTrueCell) {
  Sheet s("Sheet1");
  s.set_cell_value(1U, 1U, Value::boolean(true));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("<c r=\"B2\" t=\"b\"><v>1</v></c>"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, BoolFalseCell) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::boolean(false));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("<c r=\"A1\" t=\"b\"><v>0</v></c>"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, TextCell) {
  Sheet s("Sheet1");
  s.set_cell_value(2U, 2U, Value::text("hello"));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("<c r=\"C3\" t=\"inlineStr\"><is><t>hello</t></is></c>"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, EscapedText) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::text("a&b<c>"));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("a&amp;b&lt;c&gt;"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, UnicodeText) {
  Sheet s("Sheet1");
  // "日本" in UTF-8: E6 97 A5 E6 9C AC.
  s.set_cell_value(0U, 0U, Value::text("\xE6\x97\xA5\xE6\x9C\xAC"));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("\xE6\x97\xA5\xE6\x9C\xAC"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, ErrorCell) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::error(ErrorCode::Div0));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("<c r=\"A1\" t=\"e\"><v>#DIV/0!</v></c>"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, BlankCellOmitted) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::blank());
  // The cell exists in storage but holds Blank; no <c> should be emitted.
  EXPECT_EQ(BuildSheetDataXml(s), "<sheetData/>");
}

// ---------------------------------------------------------------------------
// Number normalisations
// ---------------------------------------------------------------------------

TEST(BuildSheetDataXml, NaNDowngradesToNum) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(std::numeric_limits<double>::quiet_NaN()));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("t=\"e\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("#NUM!"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, InfDowngradesToNum) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(std::numeric_limits<double>::infinity()));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("t=\"e\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("#NUM!"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, NegZeroNormalized) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(-0.0));
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("<v>0</v>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<v>-0</v>"), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// Row ordering
// ---------------------------------------------------------------------------

TEST(BuildSheetDataXml, RowsSortedAscending) {
  Sheet s("Sheet1");
  // Insert in reverse order; output must still be ascending.
  s.set_cell_value(4U, 0U, Value::number(50.0));
  s.set_cell_value(0U, 0U, Value::number(10.0));
  s.set_cell_value(2U, 0U, Value::number(30.0));
  const std::string xml = BuildSheetDataXml(s);

  const std::size_t r1 = xml.find("r=\"1\"");
  const std::size_t r3 = xml.find("r=\"3\"");
  const std::size_t r5 = xml.find("r=\"5\"");
  ASSERT_NE(r1, std::string::npos);
  ASSERT_NE(r3, std::string::npos);
  ASSERT_NE(r5, std::string::npos);
  EXPECT_LT(r1, r3);
  EXPECT_LT(r3, r5);
}

// ---------------------------------------------------------------------------
// Formula cell shapes
// ---------------------------------------------------------------------------

TEST(BuildSheetDataXml, FormulaCell) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=SUM(A1:B2)");
  const std::string xml = BuildSheetDataXml(s);
  // The leading '=' must be stripped; the formula body follows. The Sheet
  // API only exposes formula cells with blank cached_value, so the
  // <v>-shape variant is covered by the spill-anchor tests below (where
  // commit_spill sets the anchor's cached_value to cells[0]).
  EXPECT_NE(xml.find("<f>SUM(A1:B2)</f>"), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, FormulaCellBlankCachedValueOmitsV) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=1+2");
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("<f>1+2</f>"), std::string::npos) << xml;
  // The cell tag must NOT contain <v> when cached_value is blank.
  const std::size_t cell_start = xml.find("<c r=\"A1\"");
  ASSERT_NE(cell_start, std::string::npos);
  const std::size_t cell_end = xml.find("</c>", cell_start);
  ASSERT_NE(cell_end, std::string::npos);
  const std::string cell_xml = xml.substr(cell_start, cell_end - cell_start);
  EXPECT_EQ(cell_xml.find("<v>"), std::string::npos) << cell_xml;
}

TEST(BuildSheetDataXml, FormulaXmlEscaped) {
  Sheet s("Sheet1");
  // Both '&' and '<' must be XML-escaped inside <f>.
  s.set_cell_formula(0U, 0U, "=SUM(A1) & \"<x>\"");
  const std::string xml = BuildSheetDataXml(s);
  EXPECT_NE(xml.find("&amp;"), std::string::npos) << xml;
  EXPECT_NE(xml.find("&lt;x&gt;"), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// Spill anchor / phantom suppression
// ---------------------------------------------------------------------------

TEST(BuildSheetDataXml, SpillAnchorHasArrayType) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=SEQUENCE(3)");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  const std::string xml = BuildSheetDataXml(s);
  // Anchor at A1 carries t="array" on its <f>.
  EXPECT_NE(xml.find("<f t=\"array\">SEQUENCE(3)</f>"), std::string::npos) << xml;
  // No ref= attribute is emitted (legacy CSE marker we deliberately omit).
  const std::size_t f_start = xml.find("<f t=\"array\"");
  ASSERT_NE(f_start, std::string::npos);
  const std::size_t f_end = xml.find('>', f_start);
  ASSERT_NE(f_end, std::string::npos);
  const std::string f_open = xml.substr(f_start, f_end - f_start);
  EXPECT_EQ(f_open.find("ref="), std::string::npos) << f_open;
}

TEST(BuildSheetDataXml, SpillPhantomsSuppressed) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=SEQUENCE(2,2)");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 2U, std::move(cells)));

  const std::string xml = BuildSheetDataXml(s);
  // The anchor must be present.
  EXPECT_NE(xml.find("r=\"A1\""), std::string::npos) << xml;
  // Phantoms B1, A2, B2 must be entirely absent.
  EXPECT_EQ(xml.find("r=\"B1\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("r=\"A2\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("r=\"B2\""), std::string::npos) << xml;
}

TEST(BuildSheetDataXml, SpillCollisionEmitsSpillError) {
  Sheet s("Sheet1");
  // Pre-populate B1 so the 2x2 spill collides.
  s.set_cell_value(0U, 1U, Value::number(99.0));
  s.set_cell_formula(0U, 0U, "=SEQUENCE(2,2)");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0)};
  ASSERT_FALSE(s.commit_spill(0U, 0U, 2U, 2U, std::move(cells)));

  // The anchor's cached value is now #SPILL!; the existing B1 literal is
  // preserved.
  const std::string xml = BuildSheetDataXml(s);
  // A1 has a formula and an error cached value; the cell must surface
  // t="e" with #SPILL!.
  const std::size_t a1_start = xml.find("<c r=\"A1\"");
  ASSERT_NE(a1_start, std::string::npos);
  const std::size_t a1_end = xml.find("</c>", a1_start);
  ASSERT_NE(a1_end, std::string::npos);
  const std::string a1_xml = xml.substr(a1_start, a1_end - a1_start);
  EXPECT_NE(a1_xml.find("t=\"e\""), std::string::npos) << a1_xml;
  EXPECT_NE(a1_xml.find("#SPILL!"), std::string::npos) << a1_xml;

  // B1 should still be present and carry its literal 99.
  EXPECT_NE(xml.find("r=\"B1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("<v>99</v>"), std::string::npos) << xml;
}

}  // namespace
}  // namespace io
}  // namespace formulon
