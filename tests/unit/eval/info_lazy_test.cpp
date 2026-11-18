//
// Unit tests for the context-aware information predicates:
// ISFORMULA, ISREF, SHEET, SHEETS. Oracle coverage pins the
// argument-shape matrix for these functions, but workbook-layout
// dependent cases (SHEET() with no args, SHEETS() with no args,
// SHEET(qualified_ref)) cannot be asserted against the xlwings
// oracle harness because it writes every test case to a fresh sheet
// and the sheet index therefore depends on generation order. Those
// shapes are covered here with a deterministic in-memory Workbook.

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Builds a workbook with three sheets (Sheet1/Sheet2/Data) so that
// SHEET / SHEETS have a non-trivial layout to probe.
Workbook MakeThreeSheetWorkbook() {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.add_sheet("Data");
  return wb;
}

// Parses `src` and evaluates against `ctx`.
Value EvalWith(std::string_view src, const EvalContext& ctx) {
  static thread_local Arena arena;
  arena.reset();
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// ISFORMULA
// ---------------------------------------------------------------------------

TEST(BuiltinsIsFormula, FormulaCellIsTrue) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=1+2");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(A1)", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsFormula, LiteralCellIsFalse) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(A1)", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsFormula, EmptyCellIsFalse) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(A1)", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsFormula, LiteralNumberArgIsValue) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(42)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIsFormula, LiteralTextArgIsValue) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(\"A1\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIsFormula, QualifiedRefToOtherSheet) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(1).set_cell_formula(0, 0, "=100");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(Sheet2!A1)", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsFormula, RangeArgAnswersEachCell) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(1, 0, "=1+1");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(A1:A2)", ctx);
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array()->rows, 2U);
  EXPECT_FALSE(v.as_array()->cells[0].as_boolean());
  EXPECT_TRUE(v.as_array()->cells[1].as_boolean());
}

TEST(BuiltinsFormulaText, RangeArgAnswersEachCell) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=1+1");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=FORMULATEXT(A1:A2)", ctx);
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array()->rows, 2U);
  EXPECT_EQ(v.as_array()->cells[0].as_text(), "=1+1");
  ASSERT_TRUE(v.as_array()->cells[1].is_error());
  EXPECT_EQ(v.as_array()->cells[1].as_error(), ErrorCode::NA);
}

// Mac Excel 365: `=ISFORMULA(A:A)` and `=FORMULATEXT(1:1)` lift over the
// whole axis, so mid-sheet they surface #SPILL! (off the grid).
TEST(BuiltinsIsFormula, WholeColumnAnswersEachRow) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(1, 0, "=1+1");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(A:A)", ctx);
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array()->rows, Sheet::kMaxRows);
  ASSERT_EQ(v.as_array()->cols, 1U);
  EXPECT_FALSE(v.as_array()->cells[0].as_boolean());
  EXPECT_TRUE(v.as_array()->cells[1].as_boolean());
  EXPECT_FALSE(v.as_array()->cells[Sheet::kMaxRows - 1U].as_boolean());
}

TEST(BuiltinsFormulaText, WholeRowAnswersEachColumn) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 2, "=1+1");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=FORMULATEXT(1:1)", ctx);
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array()->rows, 1U);
  ASSERT_EQ(v.as_array()->cols, Sheet::kMaxCols);
  EXPECT_EQ(v.as_array()->cells[2].as_text(), "=1+1");
  ASSERT_TRUE(v.as_array()->cells[0].is_error());
  EXPECT_EQ(v.as_array()->cells[0].as_error(), ErrorCode::NA);
}

TEST(BuiltinsIsFormula, LetBoundRefReportsTargetFormulaCell) {
  // `=LET(c, A1, ISFORMULA(c))` must look through the LET binding to A1 and
  // report whether A1 holds a formula. A1 is a formula cell -> TRUE.
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=1+2");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(c, A1, ISFORMULA(c))", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsFormula, LetBoundRefToLiteralCellIsFalse) {
  // A1 holds a literal, so the resolved binding reports FALSE.
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(c, A1, ISFORMULA(c))", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsFormula, LetScalarBindingIsValue) {
  // `=LET(x, 5, ISFORMULA(x))` binds a scalar; the name carries no AST, so
  // ISFORMULA sees a non-reference argument and surfaces #VALUE!.
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(x, 5, ISFORMULA(x))", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIsFormula, OffsetReferenceReportsTargetFormulaCell) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=1+2");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(OFFSET(A1,0,0))", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// FORMULATEXT
// ---------------------------------------------------------------------------

TEST(BuiltinsFormulaText, OffsetReferenceReturnsTargetFormulaText) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=1+2");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=FORMULATEXT(OFFSET(A1,0,0))", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "=1+2");
}

TEST(BuiltinsFormulaText, LetBoundReferenceReturnsTargetFormulaText) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_formula(0, 0, "=1+2");
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(r,A1,FORMULATEXT(r))", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "=1+2");
}

TEST(BuiltinsFormulaText, EnteredFormulaShowsRespelledLiterals) {
  Workbook wb = MakeThreeSheetWorkbook();
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 0, "=+1.50+ -1E21*SUM(-0,1E-5)-+2"));
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=FORMULATEXT(A1)", ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "=1.5+ -1E+21*SUM(0,0.00001)-2");
}

// tblSales covers A1:C6; F2 (outside) and C2 (inside) hold stored references.
Workbook MakeTableWorkbook() {
  Workbook wb = MakeThreeSheetWorkbook();
  TableMetadata table;
  table.id = 1;
  table.name = "tblSales";
  table.display_name = "tblSales";
  table.ref = "A1:C6";
  table.columns.emplace_back(1U, "Item", "", "", "");
  table.columns.emplace_back(2U, "Qty", "", "", "");
  table.columns.emplace_back(3U, "Calc", "", "", "");
  wb.mutable_tables().push_back(table);
  wb.sheet(0).set_cell_formula(1, 2, "=tblSales[[#This Row],[Qty]]");
  wb.sheet(0).set_cell_formula(1, 7, "=tblSales[[#This Row],[Qty]]");
  wb.sheet(0).set_cell_formula(1, 8, "=COUNTA(tblSales[[#Headers],[Qty]])");
  wb.sheet(0).set_cell_formula(1, 9, "=tblSales[[#This Row],[Qty]:[Calc]]");
  return wb;
}

std::string FormulaTextOf(const Workbook& wb, std::string_view ref) {
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith(std::string("=FORMULATEXT(") + std::string(ref) + ")", ctx);
  EXPECT_TRUE(v.is_text());
  return v.is_text() ? std::string(v.as_text()) : std::string();
}

TEST(BuiltinsFormulaText, StructuredRefInsideTableOmitsTableName) {
  const Workbook wb = MakeTableWorkbook();
  EXPECT_EQ(FormulaTextOf(wb, "C2"), "=[@Qty]");
}

TEST(BuiltinsFormulaText, StructuredRefOutsideTableKeepsTableName) {
  const Workbook wb = MakeTableWorkbook();
  EXPECT_EQ(FormulaTextOf(wb, "H2"), "=tblSales[@Qty]");
}

TEST(BuiltinsFormulaText, StructuredRefMultiColumnThisRowOutside) {
  const Workbook wb = MakeTableWorkbook();
  EXPECT_EQ(FormulaTextOf(wb, "J2"), "=tblSales[@[Qty]:[Calc]]");
}

TEST(BuiltinsFormulaText, StructuredRefSpecifierIsLocalizedInJapaneseProfile) {
  Workbook wb = MakeTableWorkbook();
  wb.set_excel_profile(ExcelProfile{ExcelHost::kMac365, ExcelLocale::kJaJP});
  EXPECT_EQ(FormulaTextOf(wb, "I2"), "=COUNTA(tblSales[[#見出し],[Qty]])");
}

TEST(BuiltinsFormulaText, StructuredRefWholeRowShowsBareAt) {
  Workbook wb = MakeTableWorkbook();
  wb.sheet(0).set_cell_formula(1, 10, "=tblSales[[#This Row]]");
  EXPECT_EQ(FormulaTextOf(wb, "K2"), "=tblSales[@]");
}

// ---------------------------------------------------------------------------
// ISREF
// ---------------------------------------------------------------------------

TEST(BuiltinsIsRef, StaticRefIsTrue) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISREF(A1)", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsRef, LiteralNumberIsFalse) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISREF(42)", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsRef, ArithmeticIsFalse) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISREF(1+2)", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsRef, LetScalarBindingIsFalse) {
  // `=LET(x, 5, ISREF(x))` binds a scalar value; x is not a reference, so
  // ISREF returns FALSE. Before the fix, the NameRef AST shape made this
  // wrongly report TRUE.
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(x, 5, ISREF(x))", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIsRef, LetReferenceBindingIsTrue) {
  // `=LET(r, A1, ISREF(r))` binds a single-cell reference; resolving the
  // name surfaces the underlying Ref shape, so ISREF returns TRUE.
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(r, A1, ISREF(r))", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsRef, LetRangeBindingIsTrue) {
  // A range binding resolves to a RangeOp shape, which is also a reference.
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(r, A1:A3, ISREF(r))", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsIsRef, ChooseScalarIsFalse) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISREF(CHOOSE(1,42))", ctx);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// SHEET — workbook-layout-dependent; oracle harness cannot cover this.
// ---------------------------------------------------------------------------

TEST(BuiltinsSheet, NoArgsReturnsCurrentSheetIndex) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  {
    const EvalContext ctx(wb, wb.sheet(0), state);
    const Value v = EvalWith("=SHEET()", ctx);
    ASSERT_TRUE(v.is_number());
    EXPECT_EQ(v.as_number(), 1.0);
  }
  {
    const EvalContext ctx(wb, wb.sheet(1), state);
    const Value v = EvalWith("=SHEET()", ctx);
    ASSERT_TRUE(v.is_number());
    EXPECT_EQ(v.as_number(), 2.0);
  }
  {
    const EvalContext ctx(wb, wb.sheet(2), state);
    const Value v = EvalWith("=SHEET()", ctx);
    ASSERT_TRUE(v.is_number());
    EXPECT_EQ(v.as_number(), 3.0);
  }
}

TEST(BuiltinsSheet, LocalRefIsCurrentSheet) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(1), state);
  const Value v = EvalWith("=SHEET(A1)", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsSheet, QualifiedRefResolvesToTargetSheet) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=SHEET(Data!A1)", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

// A table reference names the table's sheet, whatever sheet the formula is on.
TEST(BuiltinsSheet, TableReferenceIsTheTablesSheet) {
  Workbook wb = MakeThreeSheetWorkbook();
  TableMetadata table;
  table.sheet_index = 2;
  table.name = "Tbl";
  table.display_name = "Tbl";
  table.ref = "A1:B3";
  for (const char* name : {"Jan", "Feb"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  for (const char* formula : {"=SHEET(Tbl[])", "=SHEET(Tbl[Feb])", "=SHEET(Tbl[#All])"}) {
    const Value v = EvalWith(formula, ctx);
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 3.0) << formula;
  }
}

TEST(BuiltinsSheets, TableReferenceIsOneSheet) {
  Workbook wb = MakeThreeSheetWorkbook();
  TableMetadata table;
  table.sheet_index = 2;
  table.name = "Tbl";
  table.display_name = "Tbl";
  table.ref = "A1:B3";
  for (const char* name : {"Jan", "Feb"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  for (const char* formula : {"=SHEETS(Tbl[])", "=SHEETS(Tbl[Feb])"}) {
    const Value v = EvalWith(formula, ctx);
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 1.0) << formula;
  }
}

// An Excel-saved book's `SHEET(TableN[])` cells, one table per sheet, recalculate to the sheet numbers Excel cached.
TEST(BuiltinsSheet, ExcelSavedTableReferencesGiveTheirSheets) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(
      FORMULON_FIXTURES_DIR "/../oracle/external/ironcalc/fixtures/calc_test_no_export/tables.xlsx");
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::read_ooxml(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_full(default_registry())));
  const Sheet& s = wb.sheet(wb.sheet_index_by_name("Sheet1"));
  for (std::uint32_t row = 4U; row < 8U; ++row) {
    SCOPED_TRACE(row);
    const Value v = s.resolve_cell_value(row, 0U);
    ASSERT_TRUE(v.is_number());
    EXPECT_EQ(v.as_number(), static_cast<double>(row - 3U));
  }
}

TEST(BuiltinsSheet, TextLookupResolvesByName) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=SHEET(\"Sheet2\")", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsSheet, UnknownNameIsNA) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=SHEET(\"Missing\")", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// Mac Excel 365: an array argument is one #N/A, even when every element
// names a sheet, and even for one element or a computed array.
TEST(BuiltinsSheet, ArrayArgumentIsOneNA) {
  Workbook wb = MakeThreeSheetWorkbook();
  wb.sheet(0).set_cell_value(0, 5, Value::text("Sheet2"));
  wb.sheet(0).set_cell_value(1, 5, Value::text("Sheet3"));
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  for (const char* src : {"=SHEET({\"Sheet2\",\"Sheet3\"})", "=SHEET({\"Sheet3\"})", "=SHEET({1,\"Sheet2\"})",
                          "=SHEET(F1:F2&\"\")", "=SHEET(TRANSPOSE(F1:F2))"}) {
    const Value v = EvalWith(src, ctx);
    ASSERT_TRUE(v.is_error()) << src << " " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), ErrorCode::NA) << src;
  }
  const Value ref = EvalWith("=SHEET(INDEX(F1:F2,0,1))", ctx);
  ASSERT_TRUE(ref.is_number()) << ref.debug_to_string();
  EXPECT_EQ(ref.as_number(), 1.0);
}

TEST(BuiltinsSheet, UnboundContextReturnsValue) {
  // Default-constructed context: no sheet, no workbook.
  const EvalContext ctx;
  const Value v = EvalWith("=SHEET()", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSheet, LetBoundReferenceRetainsQualifiedSheet) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(r,Sheet2!A1,SHEET(r))", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// SHEETS
// ---------------------------------------------------------------------------

TEST(BuiltinsSheets, NoArgsReturnsWorkbookCount) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=SHEETS()", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsSheets, SingleRefIsOne) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=SHEETS(A1)", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsSheets, RangeIsOne) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=SHEETS(A1:B2)", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsSheets, NonReferenceIsNA) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=SHEETS(42)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsSheets, LetBoundScalarIsNA) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=LET(x,42,SHEETS(x))", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsSheets, NoWorkbookReturnsOne) {
  // Sheet-only context (no workbook).
  Sheet sheet("LoneSheet");
  EvalState state;
  const EvalContext ctx(sheet, state);
  const Value v = EvalWith("=SHEETS()", ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
