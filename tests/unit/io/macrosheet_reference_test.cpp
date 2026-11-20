//
// Worksheet formulas that refer to an Excel 4.0 macro sheet. Excel evaluates such a reference like any cell
// reference, and leaves a macro sheet out of a 3-D span. The engine keeps a macro sheet as an opaque part saved
// verbatim, loading only its stored cell values (not its formulas), so a reference reads the value Excel cached.
// The expectations are Excel's cached values (Mac 16.115).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFile(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::vector<std::uint8_t> FixtureBytes() {
  return ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/macrosheet_reference/reference_into_macro_sheet.xlsm");
}

// The macro-sheet reference book, loaded and fully recalculated. Sheet 0 holds the formulas in C1:C12 over the macro
// sheet Macro1 (A1 = 5, A2 = "txt", A3 = #N/A, A4 = 1+1, A6 = A1*2, B1:B3 = 1..3) and B1 = 7.
Workbook LoadAndRecalc() {
  const std::vector<std::uint8_t> bytes = FixtureBytes();
  auto loaded = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  if (!loaded) {
    return Workbook::create_empty();
  }
  Workbook wb = std::move(loaded.value().workbook);
  auto stats = wb.recalc_full(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(stats)) << (stats ? "" : stats.error().message);
  return wb;
}

std::vector<Value> Column(const Workbook& wb, const std::vector<std::uint32_t>& rows) {
  std::vector<Value> out;
  for (const std::uint32_t row : rows) {
    out.push_back(wb.sheet(0).resolve_cell_value(row, 2U));
  }
  return out;
}

TEST(MacroSheetReference, TheMacroSheetIsKeptAsAnOpaqueSheet) {
  const Workbook wb = LoadAndRecalc();
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(1).kind(), SheetKind::kMacrosheet);
}

// C5 =Macro1!A5 (blank), C8 =SUM(W1), C9 =Macro1!M1 (blank), C10 =SUM(Sheet1:Macro1!B1), C12 =ROWS(Macro1!A1:A6):
// blank-as-0, the shape of a range, or a sum that leaves the macro sheet out.
TEST(MacroSheetReference, ReferencesThatDoNotDependOnMacroSheetCellsMatchExcel) {
  const Workbook wb = LoadAndRecalc();
  const std::vector<Value> expected = {Value::number(0.0), Value::number(0.0), Value::number(0.0), Value::number(7.0),
                                       Value::number(6.0)};
  EXPECT_EQ(Column(wb, {4U, 7U, 8U, 9U, 11U}), expected);
}

// C1 =Macro1!A1, C2 =Macro1!A2, C3 =Macro1!A3, C4 =Macro1!A4, C6 =Macro1!A6, C7 =SUM(Macro1!B1:B3): Excel's cached
// values, read from the macro sheet's stored cell values.
TEST(MacroSheetReference, AReferenceIntoMacroSheetCellsReadsTheirCachedValues) {
  const Workbook wb = LoadAndRecalc();
  const std::vector<Value> excel = {Value::number(5.0), Value::text("txt"),  Value::error(ErrorCode::NA),
                                    Value::number(2.0), Value::number(10.0), Value::number(6.0)};
  EXPECT_EQ(Column(wb, {0U, 1U, 2U, 3U, 5U, 6U}), excel);
}

// Excel returns TRUE for C11 =ISFORMULA(Macro1!A4); the engine does not load macro-sheet formulas, so it reads FALSE.
TEST(MacroSheetReference, IsFormulaOverAMacroSheetFormulaCellIsFalseBecauseMacroSheetFormulasAreNotLoaded) {
  const Workbook wb = LoadAndRecalc();
  EXPECT_EQ(Column(wb, {10U}), (std::vector<Value>{Value::boolean(false)}));
}

// The macro sheet's cells keep their stored values through a full recalculation and carry no formula text.
TEST(MacroSheetReference, MacroSheetCellsKeepTheirCachedValuesAndHoldNoFormula) {
  const Workbook wb = LoadAndRecalc();
  EXPECT_EQ(wb.sheet(1).resolve_cell_value(5U, 0U), Value::number(10.0));
  EXPECT_EQ(wb.sheet(1).resolve_cell_value(3U, 0U), Value::number(2.0));
  for (const std::uint32_t row : {0U, 1U, 2U, 3U, 5U}) {
    const Cell* cell = wb.sheet(1).cell_at(row, 0U);
    ASSERT_NE(cell, nullptr) << "row " << row;
    EXPECT_TRUE(cell->formula_text.empty()) << "row " << row;
  }
}

// The macro sheet part is written back byte for byte.
TEST(MacroSheetReference, SavingKeepsTheMacroSheetPartByteIdentical) {
  const std::vector<std::uint8_t> input = FixtureBytes();
  auto loaded = io::read_ooxml(io::ByteSpan{input.data(), input.size()});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto saved = io::write_ooxml(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  std::string before;
  std::string after;
  ASSERT_TRUE(test::extract_part(io::ByteSpan{input.data(), input.size()}, "xl/macrosheets/sheet1.xml", &before));
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/macrosheets/sheet1.xml", &after));
  EXPECT_EQ(before, after);
}

// An .xlsb save writes a macro sheet read from .xlsx as a worksheet holding its stored values.
TEST(MacroSheetReference, AnXlsbSaveWritesTheMacroSheetAsAWorksheetOfItsStoredValues) {
  const std::vector<std::uint8_t> input = FixtureBytes();
  auto loaded = io::read_ooxml(io::ByteSpan{input.data(), input.size()});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto saved = io::xlsb::write_xlsb(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto back = io::xlsb::read_xlsb(io::ByteSpan{saved.value().data(), saved.value().size()});
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  const Workbook& wb = back.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(1).kind(), SheetKind::kWorksheet);
  EXPECT_EQ(wb.sheet(1).resolve_cell_value(0U, 0U), Value::number(5.0));
  EXPECT_EQ(wb.sheet(1).resolve_cell_value(1U, 0U), Value::text("txt"));
  EXPECT_EQ(wb.sheet(1).resolve_cell_value(5U, 0U), Value::number(10.0));
}

}  // namespace
}  // namespace formulon
