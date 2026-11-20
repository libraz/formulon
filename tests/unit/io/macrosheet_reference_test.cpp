//
// Worksheet formulas that refer to an Excel 4.0 macro sheet. Excel evaluates such a reference like any cell
// reference, and leaves a macro sheet out of a 3-D span. The engine keeps a macro sheet as an opaque part with no
// cells, so a reference into it reads blank. The expectations are Excel's cached values (Mac 16.115).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest-spi.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "sheet.h"
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

// The macro-sheet reference book, loaded and fully recalculated. Sheet 0 holds the formulas in C1:C12 over the macro
// sheet Macro1 (A1 = 5, A2 = "txt", A3 = #N/A, A4 = 1+1, A6 = A1*2, B1:B3 = 1..3) and B1 = 7.
Workbook LoadAndRecalc() {
  const std::string path =
      std::string(FORMULON_FIXTURES_DIR) + "/excel/macrosheet_reference/reference_into_macro_sheet.xlsm";
  const std::vector<std::uint8_t> bytes = ReadFile(path);
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

// C1 =Macro1!A1, C2 =Macro1!A2, C3 =Macro1!A3, C4 =Macro1!A4, C6 =Macro1!A6, C7 =SUM(Macro1!B1:B3), C11
// =ISFORMULA(Macro1!A4) against Excel's cached values.
void ExpectMacroSheetCellsAsExcel() {
  const Workbook wb = LoadAndRecalc();
  const std::vector<Value> excel = {Value::number(5.0),  Value::text("txt"),  Value::error(ErrorCode::NA),
                                    Value::number(2.0),  Value::number(10.0), Value::number(6.0),
                                    Value::boolean(true)};
  EXPECT_EQ(Column(wb, {0U, 1U, 2U, 3U, 5U, 6U, 10U}), excel);
}

// A macro sheet holds no cells in the model, so a reference into it reads blank where Excel reads the cell's value.
// Fails once a macro sheet is a reference target.
TEST(MacroSheetReference, AReferenceIntoMacroSheetCellsReadsBlankNotTheirValue) {
  const Workbook wb = LoadAndRecalc();
  EXPECT_EQ(Column(wb, {0U, 6U}), (std::vector<Value>{Value::number(0.0), Value::number(0.0)}));
  EXPECT_NONFATAL_FAILURE(ExpectMacroSheetCellsAsExcel(), "");
}

}  // namespace
}  // namespace formulon
