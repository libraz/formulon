// Copyright 2026 libraz. Licensed under the MIT License.
//
// Aggregator-context tests for ROW / COLUMN. The standalone scalar form
// (`=ROW(A1:A5)` -> 1) is pinned in `builtins_shape_test.cpp`; this file
// exercises the seam where ROW / COLUMN appear as range-shaped arguments
// to range-aware aggregators (SUM / SUMPRODUCT / PRODUCT). Mac Excel 365
// spills the call to `{1;2;3;4;5}` / `{1,2,3,4,5}` and the surrounding
// aggregator iterates the spill; Formulon mirrors that without a
// `Value::Array` runtime by routing through `expand_row_call` /
// `expand_column_call` in the aggregator dispatch path.

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

// Parse + evaluate against an empty workbook. References resolve through
// the bound sheet so RangeOps inside ROW / COLUMN parse correctly.
Value EvalIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
// SUM / SUMPRODUCT / PRODUCT over ROW(range) / COLUMN(range)
// ---------------------------------------------------------------------------

TEST(BuiltinsRowColumnArray, SumOfRowOverVerticalRange) {
  // Mac Excel 365 spills ROW(A1:A5) to {1;2;3;4;5}; SUM consumes it as 15.
  // The two oracle divergences this fixes: shape_sum_of_row_over_vertical_range.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(ROW(A1:A5))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsRowColumnArray, SumOfColumnOverHorizontalRange) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(COLUMN(A1:E1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsRowColumnArray, SumproductOfColumnOverHorizontalRange) {
  // Matches oracle case shape_sumproduct_of_column_over_horizontal_range.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUMPRODUCT(COLUMN(A1:E1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsRowColumnArray, SumproductOfRowOverVerticalRange) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUMPRODUCT(ROW(A1:A3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsRowColumnArray, SumOfRowOverOffsetRange) {
  // Verify the offset-row case (top != 0): ROW(B5:B7) -> {5;6;7} -> 18.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(ROW(B5:B7))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 18.0);
}

TEST(BuiltinsRowColumnArray, ProductOfRowAcceptsRange) {
  // PRODUCT also has accepts_ranges = true. ROW(A2:A4) -> {2;3;4} -> 24.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=PRODUCT(ROW(A2:A4))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 24.0);
}

// ---------------------------------------------------------------------------
// Non-regression: scalar / non-aggregator forms must not be array-expanded.
// ---------------------------------------------------------------------------

TEST(BuiltinsRowColumnArray, BinaryOpOverRowStillCollapses) {
  // `ROW(A1:A5)*2` is a BinaryOp whose lhs is the ROW call. Without a
  // Value::Array runtime that propagates through arithmetic, we must NOT
  // try to broadcast — the lhs collapses to scalar 1, so SUM(ROW(...)*2)
  // sees the single value 2. This pins that we did not over-reach into
  // the BinaryOp path.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(ROW(A1:A5)*2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsRowColumnArray, StandaloneRowOverRangeStillScalarOne) {
  // `eval_row_lazy` is unchanged: `=ROW(A1:A5)` outside an aggregator
  // still collapses to the rectangle's first row (1) because no
  // Value::Array exists to spill into.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=ROW(A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsRowColumnArray, StandaloneColumnOverRangeStillScalarOne) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=COLUMN(A1:E1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
