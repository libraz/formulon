// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the exclusive percentile / quartile variants:
// PERCENTILE.EXC and QUARTILE.EXC. Both share the `accepts_ranges = true`
// path with the inclusive variants and keep `propagate_errors = true`;
// non-numeric inputs are silently skipped inside the impl and errors
// short-circuit in the dispatcher. See the block comment above `Median`
// in src/eval/builtins/stats.cpp for the divergence from SUM / AVERAGE.
//
// The "exclusive" formulation uses the linear-interpolation formula
// `pos = k * (n + 1)`; the admissible domain for `k` is the open interval
// (1/(n+1), n/(n+1)) and the admissible domain for QUARTILE.EXC's `quart`
// is {1, 2, 3}. Out-of-range yields `#NUM!`; empty numeric slices also
// yield `#NUM!`.

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

// Parses `src` and evaluates it via the default function registry. Arenas
// are reset on each call.
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

// Parses `src` and evaluates it against a bound workbook + current sheet.
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
// Registry pins
// ---------------------------------------------------------------------------

TEST(BuiltinsStatsExcRegistry, NamesRegistered) {
  for (const char* name : {"PERCENTILE.EXC", "QUARTILE.EXC"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

TEST(BuiltinsStatsExcRegistry, RangeAwareAndPropagating) {
  for (const char* name : {"PERCENTILE.EXC", "QUARTILE.EXC"}) {
    const FunctionDef* def = default_registry().lookup(name);
    ASSERT_NE(def, nullptr) << name;
    EXPECT_TRUE(def->accepts_ranges) << name;
    EXPECT_TRUE(def->propagate_errors) << name;
    EXPECT_TRUE(def->range_filter_numeric_only) << name;
  }
}

// ---------------------------------------------------------------------------
// PERCENTILE.EXC
// ---------------------------------------------------------------------------

TEST(BuiltinsPercentileExc, Q1OnTenValues) {
  // n=10, k=0.25, pos = 0.25 * 11 = 2.75, idx=2, frac=0.75.
  // r = xs[1] + 0.75 * (xs[2] - xs[1]) = 2 + 0.75 * (3 - 2) = 2.75.
  const Value v = EvalSource("=PERCENTILE.EXC({1;2;3;4;5;6;7;8;9;10}, 0.25)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.75);
}

TEST(BuiltinsPercentileExc, MedianOnTenValues) {
  // n=10, k=0.5, pos = 5.5, idx=5, frac=0.5. r = 5 + 0.5 * (6 - 5) = 5.5.
  const Value v = EvalSource("=PERCENTILE.EXC({1;2;3;4;5;6;7;8;9;10}, 0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.5);
}

TEST(BuiltinsPercentileExc, Q3OnTenValues) {
  // n=10, k=0.75, pos = 8.25, idx=8, frac=0.25. r = 8 + 0.25 * (9 - 8) = 8.25.
  const Value v = EvalSource("=PERCENTILE.EXC({1;2;3;4;5;6;7;8;9;10}, 0.75)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 8.25);
}

TEST(BuiltinsPercentileExc, MedianOnThreeValuesLandsOnIndex) {
  // n=3, k=0.5, pos = 0.5 * 4 = 2.0, idx=2, frac=0. r = xs[1] = 20.
  const Value v = EvalSource("=PERCENTILE.EXC({10;20;30}, 0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsPercentileExc, BoundaryLowerIsInclusive) {
  // n=3, k = 1/(n+1) = 0.25 exactly -> pos = 1.0, idx=1, frac=0. The
  // implementation's guard `idx < 1` rejects idx==0 but admits idx==1, so
  // the boundary is accepted and returns xs[0]. Matches Excel documentation:
  // "If k is less than 1/(n + 1) or greater than n/(n + 1), returns #NUM!".
  const Value v = EvalSource("=PERCENTILE.EXC({10;20;30}, 0.25)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsPercentileExc, KBelowLowerBoundaryIsNum) {
  // n=3, k=0.24 < 1/(n+1)=0.25 -> pos=0.96, idx=0 < 1 -> #NUM!.
  const Value v = EvalSource("=PERCENTILE.EXC({10;20;30}, 0.24)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPercentileExc, KAboveUpperBoundaryIsNum) {
  // n=3, k=0.76 > n/(n+1)=0.75 -> pos=3.04, idx=3 >= n -> #NUM!.
  const Value v = EvalSource("=PERCENTILE.EXC({10;20;30}, 0.76)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPercentileExc, EmptyNumericSliceIsNum) {
  const Value v = EvalSource("=PERCENTILE.EXC(\"a\", \"b\", 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPercentileExc, DirectErrorPropagates) {
  const Value v = EvalSource("=PERCENTILE.EXC(1, 2, 3, #DIV/0!, 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsPercentileExc, RangeAndScalarK) {
  Workbook wb = Workbook::create();
  for (std::uint32_t i = 0; i < 10; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=PERCENTILE.EXC(A1:A10, 0.5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.5);
}

// ---------------------------------------------------------------------------
// QUARTILE.EXC
// ---------------------------------------------------------------------------

TEST(BuiltinsQuartileExc, Q1MatchesPercentileExcQuarter) {
  // quart=1 -> k=0.25 -> pos=0.25*11=2.75, idx=2, frac=0.75 -> 2 + 0.75 = 2.75.
  const Value v = EvalSource("=QUARTILE.EXC({1;2;3;4;5;6;7;8;9;10}, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.75);
}

TEST(BuiltinsQuartileExc, Q2MatchesMedian) {
  const Value v = EvalSource("=QUARTILE.EXC({1;2;3;4;5;6;7;8;9;10}, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.5);
}

TEST(BuiltinsQuartileExc, Q3MatchesPercentileExcThreeQuarters) {
  const Value v = EvalSource("=QUARTILE.EXC({1;2;3;4;5;6;7;8;9;10}, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 8.25);
}

TEST(BuiltinsQuartileExc, QZeroIsNum) {
  // Exclusive method has no Q0; {1,2,3} is the only valid domain.
  const Value v = EvalSource("=QUARTILE.EXC({1;2;3;4;5}, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsQuartileExc, QFourIsNum) {
  // Unlike QUARTILE.INC, QUARTILE.EXC rejects quart=4.
  const Value v = EvalSource("=QUARTILE.EXC({1;2;3;4;5}, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsQuartileExc, QFractionalIsNum) {
  const Value v = EvalSource("=QUARTILE.EXC({1;2;3;4;5}, 1.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsQuartileExc, EmptyNumericSliceIsNum) {
  const Value v = EvalSource("=QUARTILE.EXC(\"a\", \"b\", 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
