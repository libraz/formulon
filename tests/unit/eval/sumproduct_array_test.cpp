// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for `=SUMPRODUCT(...)` once `eval_sumproduct_lazy`
// consumes BinaryOp / UnaryOp arguments through `eval_node_as_array`
// (the array-context broadcaster declared in `eval/shape_ops_lazy.h`).
//
// The cases here pin the cellwise array semantics that distinguish
// SUMPRODUCT from a naive scalar collapse:
//   * `(A1:A5>2)*1` produces a 5x1 boolean rectangle that coerces to
//     0/1 cells; sum across the rectangle.
//   * `(A>2)*(B<10)*C` is the canonical SUMPRODUCT-as-implicit-AND
//     idiom; both comparisons must broadcast.
//   * 1x1 BinaryOp results broadcast against any range shape.
//   * Mixed call-returns-range operands (e.g. `ROW(A1:A5)>2`) flow
//     through `eval_node_as_array`'s range-producing-call branch.
//   * Per-cell errors short-circuit in row-major scan order.
//   * Shape-mismatched operands inside a BinaryOp surface `#VALUE!`
//     verbatim (the broadcaster does not spill an array of errors).
//   * Pre-existing non-BinaryOp paths (Ref, RangeOp, ArrayLiteral)
//     continue to behave as before.

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

// Convenience for sources that need no live workbook (literals only).
Value EvalSource(std::string_view src) {
  Workbook wb = Workbook::create();
  return EvalSourceIn(src, wb, wb.sheet(0));
}

// Populates A1..A5 in the first sheet with the supplied values.
void SetColumnA(Workbook* wb, std::initializer_list<double> values) {
  std::uint32_t row = 0;
  for (double v : values) {
    wb->sheet(0).set_cell_value(row, 0, Value::number(v));
    ++row;
  }
}

// ---------------------------------------------------------------------------
// Boolean array * scalar: the canonical "count-cells-passing-predicate" form.
// ---------------------------------------------------------------------------

TEST(SumproductArray, BooleanArrayTimesScalar) {
  // A=[1,2,3,4,5]; (A>2)*1 = {0,0,1,1,1}; sum = 3.
  Workbook wb = Workbook::create();
  SetColumnA(&wb, {1.0, 2.0, 3.0, 4.0, 5.0});
  const Value v = EvalSourceIn("=SUMPRODUCT((A1:A5>2)*1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// Two boolean arrays + value range: implicit-AND idiom.
// ---------------------------------------------------------------------------

TEST(SumproductArray, TwoBooleanArraysWithValueRange) {
  // A=[1,2,3], B=[1,2,3], C=[100,200,300]; (A>1)={F,T,T}; (B<3)={T,T,F};
  // product = {0,1,0}; * C = {0,200,0}; sum = 200.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::number(2.0));
  s.set_cell_value(2, 0, Value::number(3.0));
  s.set_cell_value(0, 1, Value::number(1.0));
  s.set_cell_value(1, 1, Value::number(2.0));
  s.set_cell_value(2, 1, Value::number(3.0));
  s.set_cell_value(0, 2, Value::number(100.0));
  s.set_cell_value(1, 2, Value::number(200.0));
  s.set_cell_value(2, 2, Value::number(300.0));
  const Value v = EvalSourceIn("=SUMPRODUCT((A1:A3>1)*(B1:B3<3),C1:C3)", wb, s);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 200.0);
}

// ---------------------------------------------------------------------------
// 1x1 BinaryOp result paired with another 1x1 BinaryOp.
// ---------------------------------------------------------------------------

TEST(SumproductArray, OneByOneBinaryOpsCompose) {
  // (1>0)*5 evaluates to a 1x1 array containing 5.0; (2+3)*2 evaluates to
  // a 1x1 array containing 10.0. SUMPRODUCT collapses both to the product
  // 50.0. Verifies the BinaryOp branch produces a properly-shaped 1x1 arg
  // that still passes SUMPRODUCT's shape check against a parallel 1x1
  // BinaryOp arg (both must be exactly 1x1 - SUMPRODUCT does NOT broadcast
  // 1x1 against a non-1x1 range, matching Mac Excel).
  const Value v = EvalSource("=SUMPRODUCT((1>0)*5,(2+3)*2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

// ---------------------------------------------------------------------------
// Pure BinaryOp with no range argument: degenerate 1x1 case.
// ---------------------------------------------------------------------------

TEST(SumproductArray, PureBinaryOpNoRange) {
  // 2+3 evaluates to a 1x1 array containing 5.0; sum = 5.0. Verifies the
  // BinaryOp branch doesn't blow up when no range participates.
  const Value v = EvalSource("=SUMPRODUCT(2+3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

// ---------------------------------------------------------------------------
// BinaryOp where one operand is a range-producing function call.
// ---------------------------------------------------------------------------

TEST(SumproductArray, BinaryOpWithRangeProducingCall) {
  // ROW(A1:A5) -> {1;2;3;4;5}; (>2) -> {F;F;T;T;T}; *1 -> {0;0;1;1;1};
  // sum = 3. Exercises the range-producing-call branch inside
  // `eval_node_as_array`.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=SUMPRODUCT((ROW(A1:A5)>2)*1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// Error in a cellwise input short-circuits the whole call.
// ---------------------------------------------------------------------------

TEST(SumproductArray, CellwiseErrorPropagation) {
  // A2 = 1/0 -> #DIV/0!; the comparison rectangle inherits that error
  // cell, and SUMPRODUCT's leftmost-error rule surfaces #DIV/0!.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  s.set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=SUMPRODUCT((A1:A3>2)*1)", wb, s);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Shape mismatch inside a BinaryOp surfaces #VALUE! verbatim (scalar).
// ---------------------------------------------------------------------------

TEST(SumproductArray, ShapeMismatchInsideBinaryOp) {
  // (A1:A3>2) is 3x1; (B1:B5<10) is 5x1; the broadcaster cannot reconcile
  // them and short-circuits to #VALUE! before SUMPRODUCT itself runs.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (std::uint32_t r = 0; r < 3; ++r) {
    s.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  for (std::uint32_t r = 0; r < 5; ++r) {
    s.set_cell_value(r, 1, Value::number(static_cast<double>(r + 1)));
  }
  const Value v = EvalSourceIn("=SUMPRODUCT((A1:A3>2)*(B1:B5<10))", wb, s);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Regression: pre-existing non-BinaryOp paths are unchanged.
// ---------------------------------------------------------------------------

TEST(SumproductArray, BareRangeStillSumsLikeSum) {
  // SUMPRODUCT(A1:A3) over a single range is just SUM(A1:A3).
  Workbook wb = Workbook::create();
  SetColumnA(&wb, {1.0, 2.0, 3.0});
  const Value v = EvalSourceIn("=SUMPRODUCT(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(SumproductArray, ArrayLiteralStillWorks) {
  // SUMPRODUCT({1,2,3}) over a horizontal array literal sums to 6.
  const Value v = EvalSource("=SUMPRODUCT({1,2,3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
