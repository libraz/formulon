//
// Unit tests for `formulon::io::read_sheet_data`. Each test feeds the
// reader a tiny synthetic `<worksheet>` document and asserts the
// resulting `Workbook` state. Recalc is invoked where the test wants to
// verify formulas evaluate end-to-end.

#include "io/sheet_reader.h"

#include <cstdint>
#include <deque>
#include <string>
#include <tuple>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Returns the cached value stored at (row, col), or Blank when no cell
/// is present. Mirrors the helper used by workbook_recalc_test.cpp.
Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

/// Returns the formula text stored at (row, col), or empty when no cell.
std::string StoredFormula(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->formula_text;
  }
  return {};
}

TEST(SheetReader, EmptySheetDataNoCells) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<worksheet><sheetData/></worksheet>"));
  Workbook wb = Workbook::create();
  ASSERT_EQ(wb.sheet_count(), 1U);

  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage);
  ASSERT_TRUE(static_cast<bool>(rs));
  EXPECT_EQ(wb.sheet(0).cell_count(), 0U);
  EXPECT_EQ(ctx.pending_sst_cells.size(), 0U);
}

TEST(SheetReader, SimpleLiteralsAndFormulaRecalc) {
  // A1=1, A2=2, A3==A1+A2 ; after recalc, A3=3.
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
      "<row r=\"2\"><c r=\"A2\"><v>2</v></c></row>"
      "<row r=\"3\"><c r=\"A3\"><f>A1+A2</f></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 0U).as_number(), 1.0);
  ASSERT_TRUE(StoredValue(wb, 0U, 1U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 1U, 0U).as_number(), 2.0);
  // Before recalc: formula stored, cached value is blank.
  EXPECT_EQ(StoredFormula(wb, 0U, 2U, 0U), "=A1+A2");
  EXPECT_TRUE(StoredValue(wb, 0U, 2U, 0U).is_blank());

  // After recalc: A3 = 3.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(StoredValue(wb, 0U, 2U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 2U, 0U).as_number(), 3.0);
}

TEST(SheetReader, EmptyNumericValueElementIsBlank) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v/></c></row></sheetData></worksheet>"));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  EXPECT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_blank());
}

TEST(SheetReader, FormulaCachedValueSurvivesLoadBeforeRecalc) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(
      "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><f>UNIMPLEMENTED(1)</f><v>42</v></c>"
      "<c r=\"B1\" t=\"str\"><f>UNIMPLEMENTED_TEXT()</f><v>Excel cache</v></c></row></sheetData></worksheet>"));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  EXPECT_EQ(StoredFormula(wb, 0U, 0U, 0U), "=UNIMPLEMENTED(1)");
  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 0U).as_number(), 42.0);
  EXPECT_EQ(StoredFormula(wb, 0U, 0U, 1U), "=UNIMPLEMENTED_TEXT()");
  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 1U).is_text());
  EXPECT_EQ(StoredValue(wb, 0U, 0U, 1U).as_text(), "Excel cache");
}

TEST(SheetReader, InlineStringCell) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Hello</t></is></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  Value v = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Hello");
}

TEST(SheetReader, ErrorCellRoundTrips) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#DIV/0!</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  Value v = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(SheetReader, SharedFormulaSlaveShiftsRelativeReferences) {
  // si=0 master at E3 with formula "C3*D3"; slave at E4 with no body.
  // Excel applies the slave's row/column offset to relative references,
  // so E4 must become "C4*D4" and recalc against row 4 inputs.
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"3\"><c r=\"C3\"><v>10</v></c><c r=\"D3\"><v>20</v></c>"
      "<c r=\"E3\"><f t=\"shared\" si=\"0\" ref=\"E3:E4\">C3*D3</f></c></row>"
      "<row r=\"4\"><c r=\"C4\"><v>100</v></c><c r=\"D4\"><v>500</v></c>"
      "<c r=\"E4\"><f t=\"shared\" si=\"0\"/></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));
  EXPECT_EQ(StoredFormula(wb, 0U, 2U, 4U), "=C3*D3");
  EXPECT_EQ(StoredFormula(wb, 0U, 3U, 4U), "=C4*D4");

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(StoredValue(wb, 0U, 3U, 4U).is_number());
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 3U, 4U).as_number(), 50000.0);
}

TEST(SheetReader, SharedFormulaSlaveWithoutMasterErrors) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"7\"/></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SheetReader, IgnoresArrayFormulaWhoseRefDoesNotStartAtItsAnchor) {
  // The ref's last column lies before C1. Previously this was recorded as a
  // spill beginning at C1, so `last_col - anchor_col + 1` underflowed during
  // registration and attempted a huge allocation.
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"C1\"><f t=\"array\" ref=\"A1:B2\">1</f></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  EXPECT_TRUE(ctx.array_anchors.empty());
  EXPECT_EQ(StoredFormula(wb, 0U, 0U, 2U), "=1");
}

TEST(SheetReader, PendingSstCellsCollected) {
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>2</v></c></row>"
      "<row r=\"2\"><c r=\"A2\"><v>42</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  ASSERT_TRUE(static_cast<bool>(read_sheet_data(doc, 0U, wb, ctx, text_storage)));

  ASSERT_EQ(ctx.pending_sst_cells.size(), 2U);
  EXPECT_EQ(std::get<0>(ctx.pending_sst_cells[0]), 0U);
  EXPECT_EQ(std::get<1>(ctx.pending_sst_cells[0]), 0U);
  EXPECT_EQ(std::get<2>(ctx.pending_sst_cells[0]), 0U);
  EXPECT_EQ(std::get<0>(ctx.pending_sst_cells[1]), 0U);
  EXPECT_EQ(std::get<1>(ctx.pending_sst_cells[1]), 1U);
  EXPECT_EQ(std::get<2>(ctx.pending_sst_cells[1]), 2U);

  // Placeholders are written as Text("").
  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_text());
  EXPECT_EQ(a1.as_text(), "");
  Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(b1.is_text());
  EXPECT_EQ(b1.as_text(), "");
  // The non-shared-string cell is unaffected.
  Value a2 = StoredValue(wb, 0U, 1U, 0U);
  ASSERT_TRUE(a2.is_number());
  EXPECT_DOUBLE_EQ(a2.as_number(), 42.0);
}

TEST(SheetReader, RejectsOutOfRangeSheetIndex) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<worksheet><sheetData/></worksheet>"));
  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, /*sheet_index=*/5U, wb, ctx, text_storage);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(SheetReader, MalformedCellPropagates) {
  // <c r="A0"> — row 0 is invalid.
  const char* xml =
      "<worksheet><sheetData>"
      "<row r=\"1\"><c r=\"A0\"><v>1</v></c></row>"
      "</sheetData></worksheet>";
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml));

  Workbook wb = Workbook::create();
  SheetReadContext ctx;
  std::deque<std::string> text_storage;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, text_storage);
  ASSERT_FALSE(static_cast<bool>(rs));
  EXPECT_EQ(rs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

}  // namespace
}  // namespace io
}  // namespace formulon
