// Copyright 2026 libraz. Licensed under the MIT License.
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
