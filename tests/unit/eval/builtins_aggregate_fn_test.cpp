// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for AGGREGATE — the multi-mode aggregator that dispatches
// on a leading numeric `function_num` (1..19) plus a 0..7 options byte.
// Tests pin:
//   * Each of the 13 SUBTOTAL-aligned modes (1..13) returns the right scalar
//     on a simple {1, 2, 3, 4, 5} range.
//   * Each of the 6 k-arg modes (14..19) returns the right scalar with a
//     sensible k.
//   * The error-ignore bit (mask 2 of `options`) toggles "propagate vs
//     drop" for range-sourced errors.
//   * function_num truncation: 9.7 -> SUM, etc.
//   * function_num / options out-of-range surface `#VALUE!`.
//   * k-arg constraints: PERCENTILE.EXC with too-small p, QUARTILE.* with
//     bad quart, LARGE/SMALL with k > count -> `#NUM!`.
//   * MODE.SNGL with no duplicates -> `#N/A`.
//   * Mixed numeric / blank / text range cells: numeric branches drop
//     non-numerics; COUNTA counts non-blanks.

#include <cmath>
#include <cstdint>
#include <string>
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

// Seeds A1:A5 with 1, 2, 3, 4, 5.
Workbook MakeOneToFive() {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  return wb;
}

// ---------------------------------------------------------------------------
// Lazy dispatch routing
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, NotInEagerRegistry) {
  // AGGREGATE is wired through `kLazyDispatch` in `tree_walker.cpp`, not
  // through `default_registry().register_function`. The eager dispatcher
  // can't preserve where the data range ends and the trailing `k` begins
  // for codes 14..19, so a lazy registration is required.
  const FunctionDef* def = default_registry().lookup("AGGREGATE");
  EXPECT_EQ(def, nullptr);
}

// ---------------------------------------------------------------------------
// Codes 1..13 over {1, 2, 3, 4, 5}
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, Code1IsAverage) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(1,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsAggregateFn, Code2IsCount) {
  Workbook wb = MakeOneToFive();
  // Add a non-numeric cell to A6 to confirm COUNT skips it.
  wb.sheet(0).set_cell_value(5, 0, Value::text("text"));
  const Value v = EvalSourceIn("=AGGREGATE(2,0,A1:A6)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsAggregateFn, Code3IsCountA) {
  Workbook wb = MakeOneToFive();
  wb.sheet(0).set_cell_value(5, 0, Value::text("text"));
  wb.sheet(0).set_cell_value(6, 0, Value::boolean(true));
  // Row 8 (index 7) blank -> COUNTA must skip it.
  const Value v = EvalSourceIn("=AGGREGATE(3,0,A1:A8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsAggregateFn, Code4IsMax) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(4,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsAggregateFn, Code5IsMin) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(5,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsAggregateFn, Code6IsProduct) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(6,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 120.0);  // 1*2*3*4*5
}

TEST(BuiltinsAggregateFn, Code7IsStdevSample) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(7,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), std::sqrt(2.5));
}

TEST(BuiltinsAggregateFn, Code8IsStdevPopulation) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(8,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), std::sqrt(2.0));
}

TEST(BuiltinsAggregateFn, Code9IsSum) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(9,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsAggregateFn, Code10IsVarSample) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(10,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.5);
}

TEST(BuiltinsAggregateFn, Code11IsVarPopulation) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(11,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsAggregateFn, Code12IsMedian) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(12,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsAggregateFn, Code12MedianEvenCount) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=AGGREGATE(12,0,A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.5);
}

TEST(BuiltinsAggregateFn, Code13IsModeSngl) {
  // {1, 2, 2, 3, 4} -> mode is 2.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=AGGREGATE(13,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsAggregateFn, Code13ModeSnglNoDuplicateIsNA) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(13,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Codes 14..19 (k-arg modes)
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, Code14LargeSecond) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(14,0,A1:A5,2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);  // 2nd-largest of {1..5}
}

TEST(BuiltinsAggregateFn, Code15SmallSecond) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(15,0,A1:A5,2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);  // 2nd-smallest of {1..5}
}

TEST(BuiltinsAggregateFn, Code16PercentileIncMedian) {
  Workbook wb = MakeOneToFive();
  // PERCENTILE.INC at 0.5 -> median = 3.
  const Value v = EvalSourceIn("=AGGREGATE(16,0,A1:A5,0.5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsAggregateFn, Code16PercentileIncQuarter) {
  Workbook wb = MakeOneToFive();
  // PERCENTILE.INC at 0.25 — Excel: position = 1 + 0.25*4 = 2.0 -> xs[1] = 2.
  const Value v = EvalSourceIn("=AGGREGATE(16,0,A1:A5,0.25)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsAggregateFn, Code17QuartileIncSecond) {
  Workbook wb = MakeOneToFive();
  // QUARTILE.INC quart = 2 -> median = 3.
  const Value v = EvalSourceIn("=AGGREGATE(17,0,A1:A5,2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsAggregateFn, Code18PercentileExcMedian) {
  Workbook wb = MakeOneToFive();
  // PERCENTILE.EXC at 0.5 — position = 0.5*6 = 3.0 -> xs[2] = 3.
  const Value v = EvalSourceIn("=AGGREGATE(18,0,A1:A5,0.5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsAggregateFn, Code19QuartileExcSecond) {
  Workbook wb = MakeOneToFive();
  // QUARTILE.EXC quart = 2 -> PERCENTILE.EXC at 0.5 = 3.
  const Value v = EvalSourceIn("=AGGREGATE(19,0,A1:A5,2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// Error-ignore bit (mask 2)
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, OptionsZeroPropagatesError) {
  Workbook wb = MakeOneToFive();
  wb.sheet(0).set_cell_value(2, 0, Value::error(ErrorCode::Div0));
  const Value v = EvalSourceIn("=AGGREGATE(9,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsAggregateFn, OptionsSixIgnoresError) {
  Workbook wb = MakeOneToFive();
  wb.sheet(0).set_cell_value(2, 0, Value::error(ErrorCode::Div0));
  // Options 6 (mask 2 set) drops the #DIV/0! cell silently; SUM of the
  // remaining 1+2+4+5 = 12.
  const Value v = EvalSourceIn("=AGGREGATE(9,6,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 12.0);
}

TEST(BuiltinsAggregateFn, OptionsTwoIgnoresErrorMatchesSix) {
  Workbook wb = MakeOneToFive();
  wb.sheet(0).set_cell_value(2, 0, Value::error(ErrorCode::Div0));
  const Value v2 = EvalSourceIn("=AGGREGATE(9,2,A1:A5)", wb, wb.sheet(0));
  const Value v6 = EvalSourceIn("=AGGREGATE(9,6,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v2.is_number());
  ASSERT_TRUE(v6.is_number());
  // The nested-call bit is unobservable in Formulon today, so 2 and 6
  // collapse to identical behaviour.
  EXPECT_DOUBLE_EQ(v2.as_number(), v6.as_number());
}

// ---------------------------------------------------------------------------
// function_num truncation
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, FractionalCodeTruncatesToSum) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(9.7,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

// ---------------------------------------------------------------------------
// Bad function_num / options
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, ZeroFunctionNumRejected) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(0,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAggregateFn, OutOfRangeFunctionNumRejected) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(20,0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAggregateFn, OutOfRangeOptionsRejected) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(9,8,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAggregateFn, NegativeOptionsRejected) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(9,-1,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// k-arg constraints
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, PercentileExcTooSmallProbability) {
  Workbook wb = MakeOneToFive();
  // n = 5, p = 0.01 -> position = 0.01*6 = 0.06 < 1 -> #NUM!.
  const Value v = EvalSourceIn("=AGGREGATE(18,0,A1:A5,0.01)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsAggregateFn, QuartileIncBadQuart) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(17,0,A1:A5,5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsAggregateFn, QuartileExcQuartZeroRejected) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(19,0,A1:A5,0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsAggregateFn, QuartileExcQuartFourRejected) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(19,0,A1:A5,4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsAggregateFn, LargeKExceedsCount) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(14,0,A1:A5,6)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsAggregateFn, SmallKZeroRejected) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(15,0,A1:A5,0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsAggregateFn, KArgCodeRequiresExactArity) {
  // AGGREGATE(14, options, data, k) — extra args are rejected.
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(14,0,A1:A5,A1:A5,2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Provenance: range-sourced non-numerics are dropped (numeric branches);
// COUNTA counts non-blanks.
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, SumSkipsNonNumericInRange) {
  Workbook wb = MakeOneToFive();
  wb.sheet(0).set_cell_value(5, 0, Value::text("nope"));
  wb.sheet(0).set_cell_value(6, 0, Value::boolean(true));
  // Row 8 (index 7) blank.
  const Value v = EvalSourceIn("=AGGREGATE(9,0,A1:A8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);  // 1+2+3+4+5
}

TEST(BuiltinsAggregateFn, CountASkipsBlankCountsRest) {
  Workbook wb = MakeOneToFive();
  wb.sheet(0).set_cell_value(5, 0, Value::text("text"));
  wb.sheet(0).set_cell_value(6, 0, Value::boolean(true));
  // Row 8 (index 7) blank.
  const Value v = EvalSourceIn("=AGGREGATE(3,0,A1:A8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

// ---------------------------------------------------------------------------
// Empty / degenerate ranges
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, AverageOverEmptyRangeIsDiv0) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=AGGREGATE(1,0,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsAggregateFn, MinOverEmptyRangeIsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=AGGREGATE(5,0,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsAggregateFn, ProductOverEmptyRangeIsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=AGGREGATE(6,0,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsAggregateFn, LargeOverEmptyRangeIsNum) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=AGGREGATE(14,0,A1:A3,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Arity
// ---------------------------------------------------------------------------

TEST(BuiltinsAggregateFn, MissingDataRangeRejected) {
  Workbook wb = MakeOneToFive();
  const Value v = EvalSourceIn("=AGGREGATE(9,0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
