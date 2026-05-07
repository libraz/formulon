// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the lazy `UNIQUE(array, [by_col], [exactly_once])` builtin.
//
// UNIQUE shares its TU with FILTER (`eval/dynamic_array_lazy.cpp`) and uses
// the same `eval_node_as_array` seam to keep `array`'s 2D shape intact, then
// performs O(n^2) cellwise lane comparisons under Excel-canonical equality
// (case-insensitive text). Default mode returns first-occurrence distinct
// lanes; `exactly_once = TRUE` mode returns only lanes that appear exactly
// once. Empty / all-duplicated inputs surface #CALC!.

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "test_eval_helpers.h"
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

Value EvalUnder(std::string_view src, Arena* parse_arena, Arena* eval_arena, const EvalContext& ctx) {
  parser::Parser p(src, *parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, *eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Default mode (by_col=FALSE, exactly_once=FALSE) -- distinct rows
// ---------------------------------------------------------------------------

TEST(BuiltinsUnique, OneColumnNumbersDedupedInOrder) {
  // A1:A6 = {1,2,1,3,2,1}. Expected: 3x1 column {1,2,3} (first-occurrence
  // order preserved; Excel does NOT sort).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::number(2));
  sheet.set_cell_value(2, 0, Value::number(1));
  sheet.set_cell_value(3, 0, Value::number(3));
  sheet.set_cell_value(4, 0, Value::number(2));
  sheet.set_cell_value(5, 0, Value::number(1));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A6)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 3.0);
}

TEST(BuiltinsUnique, MultiColumnRowDedupRequiresFullRowMatch) {
  // A1:B4 = {{1,10},{2,20},{1,11},{1,10}}. Row 0 and row 3 are identical;
  // row 2 differs in column B from row 0. Expected: 3x2 array
  // {{1,10},{2,20},{1,11}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(10));
  sheet.set_cell_value(1, 0, Value::number(2));
  sheet.set_cell_value(1, 1, Value::number(20));
  sheet.set_cell_value(2, 0, Value::number(1));
  sheet.set_cell_value(2, 1, Value::number(11));
  sheet.set_cell_value(3, 0, Value::number(1));
  sheet.set_cell_value(3, 1, Value::number(10));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:B4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 11.0);
}

// ---------------------------------------------------------------------------
// by_col = TRUE -- distinct columns
// ---------------------------------------------------------------------------

TEST(BuiltinsUnique, ByColTrueDedupsColumns) {
  // A1:D2 = {{1,2,1,3},{10,20,10,30}}. Columns 0 and 2 are identical pairs.
  // With by_col=TRUE expected output is 2x3 = {{1,2,3},{10,20,30}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(0, 2, Value::number(1));
  sheet.set_cell_value(0, 3, Value::number(3));
  sheet.set_cell_value(1, 0, Value::number(10));
  sheet.set_cell_value(1, 1, Value::number(20));
  sheet.set_cell_value(1, 2, Value::number(10));
  sheet.set_cell_value(1, 3, Value::number(30));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:D2, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 30.0);
}

TEST(BuiltinsUnique, ByColFalseExplicitMatchesDefault) {
  // The two-arg form with by_col=FALSE should be indistinguishable from the
  // single-arg form. Guards against accidental boolean inversion.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(7));
  sheet.set_cell_value(1, 0, Value::number(7));
  sheet.set_cell_value(2, 0, Value::number(8));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A3, FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 8.0);
}

// ---------------------------------------------------------------------------
// exactly_once = TRUE -- only lanes that appear exactly once survive
// ---------------------------------------------------------------------------

TEST(BuiltinsUnique, ExactlyOnceKeepsSingletonsOnly) {
  // A1:A6 = {1,2,1,3,2,1}. Counts: 1->3, 2->2, 3->1. Only 3 appears exactly
  // once. Expected: 1x1 array {3}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::number(2));
  sheet.set_cell_value(2, 0, Value::number(1));
  sheet.set_cell_value(3, 0, Value::number(3));
  sheet.set_cell_value(4, 0, Value::number(2));
  sheet.set_cell_value(5, 0, Value::number(1));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A6, FALSE, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 3.0);
}

TEST(BuiltinsUnique, ExactlyOnceAllDuplicatedYieldsCalc) {
  // No singletons -> #CALC! (matches Mac Excel surface for empty UNIQUE).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(5));
  sheet.set_cell_value(1, 0, Value::number(5));
  sheet.set_cell_value(2, 0, Value::number(7));
  sheet.set_cell_value(3, 0, Value::number(7));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A4, FALSE, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

// ---------------------------------------------------------------------------
// Mixed-kind value semantics -- different kinds never match
// ---------------------------------------------------------------------------

TEST(BuiltinsUnique, NumberAndTextOfSameDigitsAreDistinct) {
  // 1 (Number) and "1" (Text) are different kinds -- both kept. Excel
  // matches: cross-kind cells never collapse.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::text("1"));
  sheet.set_cell_value(2, 0, Value::number(1));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_TRUE(v.as_array_cells()[1].is_text());
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "1");
}

TEST(BuiltinsUnique, BoolAndOneAreDistinct) {
  // TRUE and 1 are different kinds -- both kept.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::boolean(true));
  sheet.set_cell_value(1, 0, Value::number(1));
  sheet.set_cell_value(2, 0, Value::boolean(true));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_TRUE(v.as_array_cells()[0].is_boolean());
  EXPECT_TRUE(v.as_array_cells()[1].is_number());
}

TEST(BuiltinsUnique, TextEqualityIsCaseInsensitive) {
  // "abc" / "ABC" / "AbC" must collapse to a single distinct text. The
  // first occurrence wins -- "abc" is preserved verbatim.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("abc"));
  sheet.set_cell_value(1, 0, Value::text("ABC"));
  sheet.set_cell_value(2, 0, Value::text("AbC"));
  sheet.set_cell_value(3, 0, Value::text("xyz"));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "abc");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "xyz");
}

TEST(BuiltinsUnique, ErrorsWithSameCodeCollapse) {
  // Two #N/A cells must collapse to one (error-code equality). #DIV/0!
  // remains distinct.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::error(ErrorCode::NA));
  sheet.set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  sheet.set_cell_value(2, 0, Value::error(ErrorCode::NA));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A3)", &parse_arena, &eval_arena, ctx);
  // UNIQUE preserves error cells verbatim in the output -- it does not
  // propagate them out as the call-level error.
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_TRUE(v.as_array_cells()[0].is_error());
  EXPECT_EQ(v.as_array_cells()[0].as_error(), ErrorCode::NA);
  EXPECT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Arity / shape gates
// ---------------------------------------------------------------------------

TEST(BuiltinsUnique, ScalarInputReturnsScalarRoundTrip) {
  // A 1x1 input is a degenerate array: it has one distinct lane, so the
  // output mirrors the input shape (1x1 array containing the same value).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(42));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 42.0);
}

TEST(BuiltinsUnique, ZeroArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsUnique, FourArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(A1:A1, FALSE, FALSE, FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsUnique, ScalarErrorArrayArgPreservedAsCell) {
  // `1/0` evaluates to a 1x1 array containing #DIV/0! (the array-context
  // seam wraps scalars). UNIQUE treats it as a single lane and returns a
  // 1x1 array carrying that error cell verbatim. When this 1x1 array
  // spills to a worksheet cell, the user sees #DIV/0!. Consistent with
  // `ErrorsWithSameCodeCollapse`: error cells are values, not propagated
  // call-level errors.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=UNIQUE(1/0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_TRUE(v.as_array_cells()[0].is_error());
  EXPECT_EQ(v.as_array_cells()[0].as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
