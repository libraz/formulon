// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for shared-formula resolution on the streaming SAX read
// path (`io::read_sheet_data_sax`). The SAX scanner now surfaces the
// `<f>` element's `t` / `si` / `ref` attributes, so a shared-formula
// group's followers (`<f t="shared" si="N"/>`, empty body) recover the
// master body shifted to their own cell — matching the DOM path. Before
// this, followers lost their formula entirely on the SAX path.

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "io/sheet_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

io::ByteSpan SpanOf(std::string_view s) {
  return io::ByteSpan{reinterpret_cast<const std::uint8_t*>(s.data()), s.size()};
}

// Drives the SAX cell reader over `sheet_xml` into sheet 0 of a fresh
// workbook and returns it.
Workbook LoadSax(std::string_view sheet_xml) {
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  auto rs = read_sheet_data_sax(SpanOf(sheet_xml), 0U, wb, ctx, text_storage);
  EXPECT_TRUE(static_cast<bool>(rs)) << (rs ? "" : rs.error().message);
  return wb;
}

TEST(SaxSharedFormula, FollowersRecoverShiftedMasterBody) {
  // Master at B1 = A1*2 (si=0); followers at B2/B3 have empty bodies and
  // must resolve to the master body shifted down one / two rows.
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><v>10</v></c>"
      "<c r=\"B1\"><f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*2</f><v>20</v></c></row>\n"
      "<row r=\"2\"><c r=\"A2\"><v>11</v></c><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>22</v></c></row>\n"
      "<row r=\"3\"><c r=\"A3\"><v>12</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>24</v></c></row>\n"
      "</sheetData></worksheet>\n";
  const Workbook wb = LoadSax(kXml);
  const Sheet& sheet = wb.sheet(0);

  const Cell* b1 = sheet.cell_at(0U, 1U);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(b1->formula_text, "=A1*2");

  const Cell* b2 = sheet.cell_at(1U, 1U);
  ASSERT_NE(b2, nullptr);
  EXPECT_EQ(b2->formula_text, "=A2*2") << "shared follower must shift the master's relative refs";

  const Cell* b3 = sheet.cell_at(2U, 1U);
  ASSERT_NE(b3, nullptr);
  EXPECT_EQ(b3->formula_text, "=A3*2");
}

TEST(SaxSharedFormula, PlainFormulaUnaffected) {
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><f>1+2</f><v>3</v></c></row>\n"
      "</sheetData></worksheet>\n";
  const Workbook wb = LoadSax(kXml);
  const Cell* a1 = wb.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  EXPECT_EQ(a1->formula_text, "=1+2");
}

TEST(SaxSharedFormula, ArrayFormulaTreatedAsPlainBody) {
  // Array (CSE) formulas are read as plain formulas (body verbatim),
  // matching the DOM path.
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1\">SUM(B1:B3)</f><v>6</v></c></row>\n"
      "</sheetData></worksheet>\n";
  const Workbook wb = LoadSax(kXml);
  const Cell* a1 = wb.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  EXPECT_EQ(a1->formula_text, "=SUM(B1:B3)");
}

TEST(SaxSharedFormula, UnknownSharedSiIsRejected) {
  // A follower referencing an unregistered `si` is corrupt, mirroring the
  // DOM path's diagnostic.
  constexpr std::string_view kXml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "<sheetData>\n"
      "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"7\"/></c></row>\n"
      "</sheetData></worksheet>\n";
  Workbook wb = Workbook::create();
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  SheetReadContext ctx;
  auto rs = read_sheet_data_sax(SpanOf(kXml), 0U, wb, ctx, text_storage);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

}  // namespace
}  // namespace io
}  // namespace formulon
