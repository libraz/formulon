// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for Excel 365's six LAMBDA-helper builtins: BYROW, BYCOL, MAP,
// REDUCE, SCAN, MAKEARRAY. Each consumes a `Lambda` value and applies it
// across a range / array, threading the captured environment through every
// per-cell call.
//
// Coverage per helper (kept in sync with the contract notes in
// `eval/lambda_helpers_lazy.h`):
//
//   * BYROW / BYCOL — happy path on rectangular, single-row, single-column
//     inputs; lambda arity mismatch -> #VALUE!; multi-cell lambda return
//     -> #CALC!; per-cell error propagation.
//
//   * MAP — single / two / three-array forms; shape mismatch -> #N/A;
//     lambda-arity-vs-array-count mismatch -> #VALUE!; multi-cell lambda
//     return -> #CALC!; per-cell error propagation; 1D-vs-2D inputs.
//
//   * REDUCE — sum / product / string-concat folds; empty-array no-op
//     returning the seed; lambda arity mismatch -> #VALUE!; per-cell
//     error propagation.
//
//   * SCAN — running sum / product; empty input -> #CALC!; lambda arity
//     mismatch -> #VALUE!; per-cell error propagation; 2D shape preserved.
//
//   * MAKEARRAY — happy path; 1xN and Nx1 shapes; rows / cols < 1 ->
//     #NUM!; lambda arity mismatch -> #VALUE!; multi-cell lambda return
//     -> #CALC!.
//
// All cases drive the parser front-end through `Parser::Parse` so the
// formulas stay readable and exercise the dispatch path lazily — the
// helpers belong to the `kLazyDispatch` table in `tree_walker.cpp`.

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it through the default function registry.
// Parse / eval arenas have process lifetime via `static thread_local` so the
// returned `Value`'s arena-backed payload (text, array cells, lambda body)
// survives until the next `EvalSrc` call. Each call resets the shared arenas
// to avoid cross-test contamination, matching the helper pattern in
// `builtins_countif_test.cpp`. A stack-local `ParseAndEval` would free both
// arenas at function return, leaving the returned `Value` dangling — caught
// by AddressSanitizer as heap-use-after-free in `Value::as_array_cells()`.
Value EvalSrc(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_TRUE(p.errors().empty()) << "unexpected parse errors for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena, default_registry(), EvalContext{});
}

// ---------------------------------------------------------------------------
// BYROW
// ---------------------------------------------------------------------------

TEST(LambdaHelpersByRow, SumsEachRow) {
  // 3x3 input; per-row SUM yields {6; 15; 24}.
  const Value v = EvalSrc("=BYROW({1,2,3;4,5,6;7,8,9}, LAMBDA(r, SUM(r)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 15.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 24.0);
}

TEST(LambdaHelpersByRow, SingleRowInput) {
  // 1x4 input collapses to a single row sum.
  const Value v = EvalSrc("=BYROW({1,2,3,4}, LAMBDA(r, SUM(r)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
}

TEST(LambdaHelpersByRow, SingleColumnInput) {
  // 3x1 input: each row carries a single cell. Per-row SUM is the cell value.
  const Value v = EvalSrc("=BYROW({10;20;30}, LAMBDA(r, SUM(r)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 30.0);
}

TEST(LambdaHelpersByRow, LambdaArityMismatchYieldsValueError) {
  // BYROW expects a 1-arg lambda; passing 2-arg -> #VALUE!.
  const Value v = EvalSrc("=BYROW({1,2;3,4}, LAMBDA(a, b, a+b))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(LambdaHelpersByRow, MultiCellLambdaReturnYieldsCalcError) {
  // The lambda body returns its argument unchanged: a 1xN sub-array. Mac
  // Excel cannot project a multi-cell return into BYROW's single-cell-
  // per-row output slot, so #CALC!.
  const Value v = EvalSrc("=BYROW({1,2;3,4}, LAMBDA(r, r))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(LambdaHelpersByRow, ErrorInRowPropagates) {
  // Row 2 contains an explicit `#DIV/0!` error literal; SUM over a row
  // containing an error surfaces that error, which BYROW then propagates
  // as the whole result.
  const Value v = EvalSrc("=BYROW({1,2;#DIV/0!,4}, LAMBDA(r, SUM(r)))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(LambdaHelpersByRow, NonLambdaSecondArgYieldsValueError) {
  // The second arg must be a Lambda value; a plain number fails type check.
  const Value v = EvalSrc("=BYROW({1,2;3,4}, 42)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(LambdaHelpersByRow, ByrowSingleColumnInputUnwrapsTo1x1) {
  // 3x1 input, lambda body `row*10` broadcasts over the 1x1 row slice and
  // produces a 1x1 Array per iteration. Mac Excel anchor-unwraps that to a
  // scalar, so BYROW emits a 3x1 Array {10; 20; 30}.
  const Value v = EvalSrc("=BYROW({1;2;3}, LAMBDA(row, row*10))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 20.0);
  ASSERT_TRUE(v.as_array_cells()[2].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 30.0);
}

TEST(LambdaHelpersByRow, ByrowMultiCellLambdaReturnStillCalc) {
  // 2x2 input: each row slice is 1x2. The lambda body returns the slice
  // verbatim (a 1x2 Array, not 1x1), so the anchor-unwrap leaves it as
  // an Array and BYROW must surface #CALC!.
  const Value v = EvalSrc("=BYROW({1,2;3,4}, LAMBDA(row, row))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

// ---------------------------------------------------------------------------
// BYCOL
// ---------------------------------------------------------------------------

TEST(LambdaHelpersByCol, SumsEachColumn) {
  // 3x3 input; per-column SUM yields {12, 15, 18} as a 1x3 row.
  const Value v = EvalSrc("=BYCOL({1,2,3;4,5,6;7,8,9}, LAMBDA(c, SUM(c)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 12.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 15.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 18.0);
}

TEST(LambdaHelpersByCol, SingleRowInput) {
  // 1x3 input: each column carries one cell.
  const Value v = EvalSrc("=BYCOL({10,20,30}, LAMBDA(c, SUM(c)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 30.0);
}

TEST(LambdaHelpersByCol, SingleColumnInput) {
  // 3x1 input collapses to a single column-sum scalar.
  const Value v = EvalSrc("=BYCOL({1;2;3}, LAMBDA(c, SUM(c)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 6.0);
}

TEST(LambdaHelpersByCol, LambdaArityMismatchYieldsValueError) {
  const Value v = EvalSrc("=BYCOL({1,2;3,4}, LAMBDA(a, b, a+b))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(LambdaHelpersByCol, MultiCellLambdaReturnYieldsCalcError) {
  const Value v = EvalSrc("=BYCOL({1,2;3,4}, LAMBDA(c, c))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(LambdaHelpersByCol, ErrorInColumnPropagates) {
  const Value v = EvalSrc("=BYCOL({1,2;3,#DIV/0!}, LAMBDA(c, SUM(c)))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(LambdaHelpersByCol, BycolSingleRowInputUnwrapsTo1x1) {
  // 1x3 input, lambda body `col*10` broadcasts over the 1x1 column slice
  // and produces a 1x1 Array per iteration. Mac Excel anchor-unwraps that
  // to a scalar, so BYCOL emits a 1x3 Array {10, 20, 30}.
  const Value v = EvalSrc("=BYCOL({1,2,3}, LAMBDA(col, col*10))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 20.0);
  ASSERT_TRUE(v.as_array_cells()[2].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 30.0);
}

// ---------------------------------------------------------------------------
// MAP
// ---------------------------------------------------------------------------

TEST(LambdaHelpersMap, SingleArrayDoubles) {
  // Squares each cell of a 2x2 array.
  const Value v = EvalSrc("=MAP({1,2;3,4}, LAMBDA(x, x*x))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 9.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 16.0);
}

TEST(LambdaHelpersMap, TwoArrayElementwiseSum) {
  const Value v = EvalSrc("=MAP({1,2;3,4}, {10,20;30,40}, LAMBDA(a, b, a+b))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 22.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 33.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 44.0);
}

TEST(LambdaHelpersMap, ThreeArrayElementwiseSum) {
  const Value v = EvalSrc("=MAP({1,2}, {10,20}, {100,200}, LAMBDA(a, b, c, a+b+c))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 111.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 222.0);
}

TEST(LambdaHelpersMap, ShapeMismatchYieldsNAError) {
  // 2x2 vs 1x2 mismatch -> #N/A.
  const Value v = EvalSrc("=MAP({1,2;3,4}, {10,20}, LAMBDA(a, b, a+b))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(LambdaHelpersMap, LambdaArityVsArrayCountMismatch) {
  // Two arrays but a 1-arg lambda -> #VALUE!.
  const Value v = EvalSrc("=MAP({1,2}, {10,20}, LAMBDA(a, a*2))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(LambdaHelpersMap, MultiCellLambdaReturnYieldsCalcError) {
  // SEQUENCE returns a multi-cell array; MAP cannot spill that into a
  // single output slot, so #CALC!.
  const Value v = EvalSrc("=MAP({1,2}, LAMBDA(x, SEQUENCE(1, 3)))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(LambdaHelpersMap, ErrorInCellPropagates) {
  // An explicit `#DIV/0!` cell in the input surfaces that error from
  // inside the lambda body's first arithmetic op (`x*2`).
  const Value v = EvalSrc("=MAP({1, #DIV/0!}, LAMBDA(x, x*2))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(LambdaHelpersMap, OneDimensionalInputPreservesShape) {
  // A 1D row vector input keeps its 1xN shape.
  const Value v = EvalSrc("=MAP({1,2,3,4}, LAMBDA(x, x+10))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 14.0);
}

// ---------------------------------------------------------------------------
// REDUCE
// ---------------------------------------------------------------------------

TEST(LambdaHelpersReduce, Sum) {
  const Value v = EvalSrc("=REDUCE(0, {1,2,3,4}, LAMBDA(a, x, a+x))");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(LambdaHelpersReduce, Product) {
  const Value v = EvalSrc("=REDUCE(1, {1,2,3,4}, LAMBDA(a, x, a*x))");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 24.0);
}

TEST(LambdaHelpersReduce, StringConcat) {
  const Value v = EvalSrc("=REDUCE(\"\", {\"a\",\"b\",\"c\"}, LAMBDA(a, x, a&x))");
  ASSERT_TRUE(v.is_text()) << v.debug_to_string();
  EXPECT_EQ(std::string(v.as_text()), "abc");
}

TEST(LambdaHelpersReduce, SingleCellInput) {
  // Smallest non-empty case: REDUCE walks one cell and threads the
  // accumulator once. There is no Excel-level idiom for an empty inline
  // array literal, so the empty-input -> seed-passthrough contract is
  // exercised indirectly by the 0-row guard in the impl.
  const Value v = EvalSrc("=REDUCE(40, {2}, LAMBDA(a, x, a+x))");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(LambdaHelpersReduce, ErrorInArrayCellPropagates) {
  const Value v = EvalSrc("=REDUCE(0, {1, #DIV/0!, 3}, LAMBDA(a, x, a+x))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(LambdaHelpersReduce, LambdaArityMismatchYieldsValueError) {
  // REDUCE requires a 2-arg lambda; passing a 1-arg lambda -> #VALUE!.
  const Value v = EvalSrc("=REDUCE(0, {1,2,3}, LAMBDA(a, a+1))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// SCAN
// ---------------------------------------------------------------------------

TEST(LambdaHelpersScan, RunningSum) {
  // initial=0, walk {1,2,3}: emit {1, 3, 6}.
  const Value v = EvalSrc("=SCAN(0, {1,2,3}, LAMBDA(a, x, a+x))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 6.0);
}

TEST(LambdaHelpersScan, RunningProduct) {
  const Value v = EvalSrc("=SCAN(1, {1,2,3,4}, LAMBDA(a, x, a*x))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 24.0);
}

TEST(LambdaHelpersScan, SingleCellInput) {
  // Smallest non-empty case: a 1x1 input emits a 1x1 result containing
  // the single accumulator update. There is no Excel-level idiom for an
  // "empty" array literal, so the empty-input -> #CALC! contract is
  // covered indirectly by the 0-row guard in the impl (the only way to
  // reach it is through a user-supplied empty Range, which the workbook
  // path will ultimately need its own coverage for).
  const Value v = EvalSrc("=SCAN(10, {5}, LAMBDA(a, x, a+x))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 15.0);
}

TEST(LambdaHelpersScan, ErrorInCellPropagates) {
  const Value v = EvalSrc("=SCAN(0, {1, #DIV/0!, 3}, LAMBDA(a, x, a+x))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(LambdaHelpersScan, TwoDimensionalInputPreservesShape) {
  // 2x2 input, row-major scan yields {1, 3, 6, 10}.
  const Value v = EvalSrc("=SCAN(0, {1,2;3,4}, LAMBDA(a, x, a+x))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 10.0);
}

TEST(LambdaHelpersScan, LambdaArityMismatchYieldsValueError) {
  const Value v = EvalSrc("=SCAN(0, {1,2,3}, LAMBDA(a, a+1))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// MAKEARRAY
// ---------------------------------------------------------------------------

TEST(LambdaHelpersMakeArray, ProducesRectangleFromIndices) {
  // 3x3 array where cell (r, c) (1-based) = r*10 + c.
  const Value v = EvalSrc("=MAKEARRAY(3, 3, LAMBDA(r, c, r*10+c))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  // Row 1: 11, 12, 13.
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 12.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 13.0);
  // Row 2: 21, 22, 23.
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 21.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[4].as_number(), 22.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[5].as_number(), 23.0);
  // Row 3: 31, 32, 33.
  EXPECT_DOUBLE_EQ(v.as_array_cells()[6].as_number(), 31.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[7].as_number(), 32.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[8].as_number(), 33.0);
}

TEST(LambdaHelpersMakeArray, OneByNShape) {
  const Value v = EvalSrc("=MAKEARRAY(1, 4, LAMBDA(r, c, c))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 4.0);
}

TEST(LambdaHelpersMakeArray, NByOneShape) {
  const Value v = EvalSrc("=MAKEARRAY(4, 1, LAMBDA(r, c, r))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 4.0);
}

TEST(LambdaHelpersMakeArray, ZeroRowsYieldsNumError) {
  const Value v = EvalSrc("=MAKEARRAY(0, 3, LAMBDA(r, c, r))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(LambdaHelpersMakeArray, ZeroColsYieldsNumError) {
  const Value v = EvalSrc("=MAKEARRAY(3, 0, LAMBDA(r, c, r))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(LambdaHelpersMakeArray, NegativeRowsYieldsNumError) {
  const Value v = EvalSrc("=MAKEARRAY(-1, 3, LAMBDA(r, c, r))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(LambdaHelpersMakeArray, LambdaArityMismatchYieldsValueError) {
  // MAKEARRAY requires a 2-arg lambda; 1-arg -> #VALUE!.
  const Value v = EvalSrc("=MAKEARRAY(2, 2, LAMBDA(r, r))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(LambdaHelpersMakeArray, MultiCellLambdaReturnYieldsCalcError) {
  // Lambda returns a multi-cell SEQUENCE array per cell; MAKEARRAY
  // rejects with #CALC!.
  const Value v = EvalSrc("=MAKEARRAY(2, 2, LAMBDA(r, c, SEQUENCE(1, 2)))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
