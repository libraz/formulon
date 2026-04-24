// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the shape / geometry-inspection built-ins:
// ROWS, COLUMNS, ROW, COLUMN, and SUMPRODUCT.
//
// All five are routed through the lazy dispatch table
// (`eval_rows_lazy`, ..., `eval_sumproduct_lazy` in
// `eval/shape_ops_lazy.cpp`) because they must introspect each
// argument's AST shape. The tests here pin the three supported
// argument kinds — `Ref`, `RangeOp(Ref, Ref)`, and `ArrayLiteral` —
// plus the scalar / non-reference fallback and the error-propagation
// rules for each function.

#include <cstdint>
#include <string_view>

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

// Parses `src` and evaluates it with no bound sheet. References inside
// `src` therefore fall back to `#NAME?`; use `EvalSourceIn` for anything
// that needs live cells.
Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// Parses `src` and evaluates against a bound workbook + current sheet so
// A1-style references resolve through `EvalContext::expand_range`.
Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Registry pins: lazy dispatch, never eager.
// ---------------------------------------------------------------------------

TEST(BuiltinsShapeRegistry, ShapeBuiltinsAreLazyOnly) {
  // None of the five impls live in the eager registry - they are routed
  // through `kLazyDispatch` in tree_walker.cpp. Pin that explicitly so a
  // future accidental `register_function(..., "ROWS", ...)` is caught.
  EXPECT_EQ(default_registry().lookup("ROWS"), nullptr);
  EXPECT_EQ(default_registry().lookup("COLUMNS"), nullptr);
  EXPECT_EQ(default_registry().lookup("ROW"), nullptr);
  EXPECT_EQ(default_registry().lookup("COLUMN"), nullptr);
  EXPECT_EQ(default_registry().lookup("SUMPRODUCT"), nullptr);
}

// ---------------------------------------------------------------------------
// ROWS / COLUMNS
// ---------------------------------------------------------------------------

TEST(BuiltinsRows, RangeRectangleReturnsRowCount) {
  Workbook wb = Workbook::create();
  // 3x2 rectangle across A1:B3.
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=ROWS(A1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsColumns, RangeRectangleReturnsColumnCount) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=COLUMNS(A1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsRows, SingleCellRefIsOne) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=ROWS(A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsColumns, SingleCellRefIsOne) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=COLUMNS(A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsRows, ArrayLiteralColumnShape) {
  // `{1;2;3}` is a 3x1 column literal (semicolons separate rows).
  const Value v = EvalSource("=ROWS({1;2;3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsRows, ArrayLiteralRowIsOne) {
  // `{1,2,3}` is a 1x3 row literal (commas separate columns).
  const Value v = EvalSource("=ROWS({1,2,3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsColumns, ArrayLiteralRowShape) {
  const Value v = EvalSource("=COLUMNS({1,2,3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsRows, ScalarFallbackIsOne) {
  // Bare number expression: treated as 1x1 per Excel's scalar treatment.
  const Value v = EvalSource("=ROWS(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsRows, ErrorSubtreePropagates) {
  // Scalar argument that evaluates to an error must propagate that error
  // unchanged instead of silently becoming 1.
  const Value v = EvalSource("=ROWS(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsRows, ZeroArgIsArityViolation) {
  const Value v = EvalSource("=ROWS()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// ROW / COLUMN (1-arg form only; zero-arity would need current-cell ctx)
// ---------------------------------------------------------------------------

TEST(BuiltinsRow, SingleRefIsOneBasedRow) {
  // C5 -> 0-based (row=4, col=2) -> 1-based row=5.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(C5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsColumn, SingleRefIsOneBasedCol) {
  // C5 -> col=2 -> 1-based col=3.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(C5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsRow, RangeOpReturnsFirstRow) {
  // B2:D4 -> row range [1, 3]. First row (1-based) = 2.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=ROW(B2:D4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsColumn, RangeOpReturnsFirstCol) {
  // B2:D4 -> col range [1, 3]. First col (1-based) = 2.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=COLUMN(B2:D4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsRow, ArrayLiteralIsValue) {
  // Array literal is not a reference; Excel's `=ROW({1,2,3})` is #VALUE!.
  const Value v = EvalSource("=ROW({1,2,3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsRow, ScalarLiteralIsValue) {
  // A bare number is not a reference.
  const Value v = EvalSource("=ROW(42)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsRow, ErrorSubtreePropagates) {
  const Value v = EvalSource("=ROW(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsRow, ZeroArgWithoutAnchorIsValue) {
  // Without a formula-cell anchor the 0-arg form has no meaningful answer,
  // so surface #VALUE! rather than inventing a row number.
  const Value v = EvalSource("=ROW()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsRow, ZeroArgReturnsAnchorRowOneBased) {
  // The oracle harness anchors each formula at its own cell so ROW() / COLUMN()
  // can report the containing cell's coordinates. Mirror that here by
  // constructing an EvalContext with the anchor explicitly.
  Workbook wb = Workbook::create();
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=ROW()", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EvalState state;
  const EvalContext ctx =
      EvalContext(wb, wb.sheet(0), state).with_formula_cell(4U, 2U);  // C5
  const Value v = evaluate(*root, eval_arena, default_registry(), ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsColumn, ZeroArgReturnsAnchorColumnOneBased) {
  Workbook wb = Workbook::create();
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=COLUMN()", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EvalState state;
  const EvalContext ctx =
      EvalContext(wb, wb.sheet(0), state).with_formula_cell(4U, 2U);  // C5
  const Value v = evaluate(*root, eval_arena, default_registry(), ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// SUMPRODUCT
// ---------------------------------------------------------------------------

TEST(BuiltinsSumproduct, TwoParallelRanges) {
  // [1,2,3] * [4,5,6] => 1*4 + 2*5 + 3*6 = 32.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3,B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 32.0);
}

TEST(BuiltinsSumproduct, SingleRangeIsPlainSum) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsSumproduct, ShapeMismatchIsValue) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(5.0));
  // A1:A2 is 2x1; B1:B3 is 3x1 -> mismatched rectangles.
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A2,B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSumproduct, TextCellTreatedAsZero) {
  // [1, "x", 3] * [10, 20, 30] = 1*10 + 0*20 + 3*30 = 100. SUMPRODUCT
  // specifically does NOT coerce non-numeric cells - they contribute 0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3,B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

TEST(BuiltinsSumproduct, ErrorInRangePropagates) {
  // An error cell in the first (leftmost) range wins.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_formula(1, 0, "=1/0");
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3,B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSumproduct, ArrayLiteralColumns) {
  // {1;2;3} * {4;5;6} = 32, identical to the range variant.
  const Value v = EvalSource("=SUMPRODUCT({1;2;3},{4;5;6})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 32.0);
}

TEST(BuiltinsSumproduct, ThreeArrays) {
  // {1;2} * {3;4} * {5;6} = 1*3*5 + 2*4*6 = 15 + 48 = 63.
  const Value v = EvalSource("=SUMPRODUCT({1;2},{3;4},{5;6})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 63.0);
}

TEST(BuiltinsSumproduct, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=SUMPRODUCT()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSumproduct, ScalarArgsMultiply) {
  // All 1x1 arguments: 2 * 3 * 4 = 24.
  const Value v = EvalSource("=SUMPRODUCT(2,3,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 24.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
