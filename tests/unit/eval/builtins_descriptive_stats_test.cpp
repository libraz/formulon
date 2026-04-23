// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the descriptive-statistics family:
// GEOMEAN / HARMEAN / DEVSQ / AVEDEV / TRIMMEAN / SKEW / SKEW.P / KURT /
// STANDARDIZE. Plus the regression standard-error STEYX, which rides the
// lazy-dispatch seam (paired-array shape check + pairwise numeric drop).
//
// These share the skip-non-numeric rule of MEDIAN / VAR (only Number kind
// participates; Text / Bool / Blank inside ranges are dropped silently).

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

// ---------------------------------------------------------------------------
// GEOMEAN
// ---------------------------------------------------------------------------

TEST(BuiltinsGeoMean, TwoValues) {
  // GEOMEAN(2, 8) = sqrt(16) = 4.
  const Value v = EvalSource("=GEOMEAN(2,8)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsGeoMean, ThreeValues) {
  // GEOMEAN(2, 4, 8) = cbrt(64) = 4.
  const Value v = EvalSource("=GEOMEAN(2,4,8)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 4.0, 1e-12);
}

TEST(BuiltinsGeoMean, ZeroIsNum) {
  const Value v = EvalSource("=GEOMEAN(1,0,3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGeoMean, NegativeIsNum) {
  const Value v = EvalSource("=GEOMEAN(1,-2,3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// HARMEAN
// ---------------------------------------------------------------------------

TEST(BuiltinsHarMean, TwoValues) {
  // HARMEAN(2, 4) = 2 / (1/2 + 1/4) = 2 / 0.75 = 8/3.
  const Value v = EvalSource("=HARMEAN(2,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 8.0 / 3.0, 1e-12);
}

TEST(BuiltinsHarMean, NegativeIsNum) {
  const Value v = EvalSource("=HARMEAN(1,-2,3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// DEVSQ
// ---------------------------------------------------------------------------

TEST(BuiltinsDevSq, SymmetricTriple) {
  // {2, 4, 6}: mean=4, SS = 4 + 0 + 4 = 8.
  const Value v = EvalSource("=DEVSQ(2,4,6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 8.0);
}

TEST(BuiltinsDevSq, SingleValueIsZero) {
  const Value v = EvalSource("=DEVSQ(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsDevSq, AllNonNumericIsZero) {
  // Range of only text -> filtered out -> empty -> returns 0 (empty sum).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  static thread_local Arena arena;
  arena.reset();
  parser::Parser p("=DEVSQ(A1:A2)", arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EvalState state;
  const EvalContext ctx(wb, wb.sheet(0), state);
  const Value v = evaluate(*root, arena, default_registry(), ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// AVEDEV
// ---------------------------------------------------------------------------

TEST(BuiltinsAveDev, SymmetricFive) {
  // {4, 5, 6, 7, 8}: mean=6, mean|dev| = (2+1+0+1+2)/5 = 1.2.
  const Value v = EvalSource("=AVEDEV(4,5,6,7,8)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.2);
}

TEST(BuiltinsAveDev, EmptyIsNum) {
  const Value v = EvalSource("=AVEDEV(\"a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// TRIMMEAN
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimMean, TrimOneEachSide) {
  // n=10, percent=0.2 -> floor(10*0.2/2)*2 = 2 trimmed (1 each side).
  // Remaining {2,3,4,5,6,7,8,9}, mean = 5.5.
  const Value v = EvalSource("=TRIMMEAN({1;2;3;4;5;6;7;8;9;10}, 0.2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.5);
}

TEST(BuiltinsTrimMean, ZeroPercentIsAverage) {
  const Value v = EvalSource("=TRIMMEAN({1;2;3;4;5}, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsTrimMean, PercentOneIsNum) {
  // percent must be strictly less than 1.
  const Value v = EvalSource("=TRIMMEAN({1;2;3;4;5}, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsTrimMean, NegativePercentIsNum) {
  const Value v = EvalSource("=TRIMMEAN({1;2;3;4;5}, -0.1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsTrimMean, SmallPercentTrimsNothing) {
  // percent=0.1 on n=5: floor(0.25) = 0, nothing trimmed -> mean of {1..5}.
  const Value v = EvalSource("=TRIMMEAN({1;2;3;4;5}, 0.1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// SKEW / SKEW.P
// ---------------------------------------------------------------------------

TEST(BuiltinsSkew, SymmetricIsZero) {
  // SKEW of symmetric data is 0.
  const Value v = EvalSource("=SKEW(1,2,3,4,5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-12);
}

TEST(BuiltinsSkew, RightSkewedIsPositive) {
  const Value v = EvalSource("=SKEW(1,1,1,2,10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
}

TEST(BuiltinsSkew, TooFewIsDiv0) {
  const Value v = EvalSource("=SKEW(1,2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSkewP, SymmetricIsZero) {
  const Value v = EvalSource("=SKEW.P(1,2,3,4,5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-12);
}

TEST(BuiltinsSkewP, ConstantIsDiv0) {
  const Value v = EvalSource("=SKEW.P(3,3,3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// KURT
// ---------------------------------------------------------------------------

TEST(BuiltinsKurt, SymmetricFive) {
  // Closed-form for {1,2,3,4,5}:
  //   n=5, mean=3, sample_var = 2.5, s = sqrt(2.5).
  //   sum((x-mean)/s)^4 = 2*(4^2 + 1^2) / 2.5^2 = 34 / 6.25 = 5.44.
  //   coeff_a = 5*6/(4*3*2) = 30/24 = 1.25
  //   coeff_b = 3*16/6 = 8
  //   KURT = 1.25 * 5.44 - 8 = 6.8 - 8 = -1.2.
  const Value v = EvalSource("=KURT(1,2,3,4,5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -1.2, 1e-12);
}

TEST(BuiltinsKurt, TooFewIsDiv0) {
  const Value v = EvalSource("=KURT(1,2,3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// STANDARDIZE
// ---------------------------------------------------------------------------

TEST(BuiltinsStandardize, Basic) {
  // (42 - 40) / 1.5 = 4/3.
  const Value v = EvalSource("=STANDARDIZE(42, 40, 1.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), (42.0 - 40.0) / 1.5, 1e-12);
}

TEST(BuiltinsStandardize, ZeroSdIsNum) {
  const Value v = EvalSource("=STANDARDIZE(42, 40, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsStandardize, NegativeSdIsNum) {
  const Value v = EvalSource("=STANDARDIZE(42, 40, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// STEYX
// ---------------------------------------------------------------------------

TEST(BuiltinsSteyx, PerfectFitIsZero) {
  // y = 2x exactly -> residual SS = 0 -> STEYX = 0.
  const Value v = EvalSource("=STEYX({2;4;6;8;10},{1;2;3;4;5})");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-12);
}

TEST(BuiltinsSteyx, NoisyData) {
  // y = {2,4,5,8,10}, x = {1,2,3,4,5}. Hand-computed:
  //   mean_x=3, mean_y=5.8.
  //   dx = {-2,-1,0,1,2}, dy = {-3.8,-1.8,-0.8,2.2,4.2}.
  //   sum_xx = 10, sum_xy = 20, sum_yy = 40.8.
  //   residual SS = sum_yy - sum_xy^2/sum_xx = 40.8 - 40 = 0.8.
  //   STEYX = sqrt(0.8 / (5 - 2)) = sqrt(0.8/3).
  const Value v = EvalSource("=STEYX({2;4;5;8;10},{1;2;3;4;5})");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_NEAR(v.as_number(), std::sqrt(0.8 / 3.0), 1e-12);
}

TEST(BuiltinsSteyx, TooFewPairsIsDiv0) {
  // n=2 -> degrees of freedom 0 -> #DIV/0!.
  const Value v = EvalSource("=STEYX({1;2},{3;4})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSteyx, ShapeMismatchIsNA) {
  const Value v = EvalSource("=STEYX({1;2;3},{1;2})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsSteyx, CollinearXIsDiv0) {
  // All x-values equal -> sum_xx = 0 -> #DIV/0!.
  const Value v = EvalSource("=STEYX({1;2;3;4},{5;5;5;5})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Registry pins
// ---------------------------------------------------------------------------

TEST(BuiltinsDescriptiveStatsRegistry, AllNamesRegistered) {
  for (const char* name :
       {"GEOMEAN", "HARMEAN", "DEVSQ", "AVEDEV", "TRIMMEAN", "SKEW", "SKEW.P", "KURT", "STANDARDIZE"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
