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

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
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

TEST(BuiltinsIsFormula, RangeArgIsValue) {
  Workbook wb = MakeThreeSheetWorkbook();
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = EvalWith("=ISFORMULA(A1:A2)", ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
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
