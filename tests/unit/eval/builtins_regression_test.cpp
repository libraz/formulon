// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the pairwise linear-regression family: CORREL,
// COVARIANCE.P, COVARIANCE.S, SLOPE, INTERCEPT, RSQ, and
// FORECAST.LINEAR (with FORECAST as a legacy alias). All seven ride
// the lazy dispatch table because their two array arguments must
// preserve `(rows, cols)` shape so the impl can reject mismatched
// rectangles with `#N/A`.
//
// Reference dataset used for most numeric cases:
//   A = [1, 2, 3, 4, 5]
//   B = [2, 4, 5, 4, 5]
// Derived stats (verified by hand and cross-checked against Python /
// NumPy):
//   mean_x = 3, mean_y = 4
//   sum_xx = 10, sum_yy = 6, sum_xy = 6
//   CORREL = sum_xy / sqrt(sum_xx*sum_yy) = 6 / sqrt(60)
//          ~ 0.77459666924148337
//   RSQ    = CORREL^2 = 0.6
//   SLOPE  = sum_xy / sum_xx = 0.6
//   INTERCEPT = mean_y - SLOPE*mean_x = 2.2
//   COVAR.P = sum_xy / n      = 1.2
//   COVAR.S = sum_xy / (n-1)  = 1.5

#include <cmath>
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

// Parses `src` and evaluates it through the default function registry
// with no bound workbook. Suitable for ArrayLiteral-only formulas.
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

// Parses `src` and evaluates it against a bound workbook + current
// sheet. Used whenever the formula references cell ranges.
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

// Populates a workbook with the reference dataset in columns A and B.
// Row order matches the top-of-file comment.
Workbook MakeReferenceWorkbook() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::number(2.0));
  s.set_cell_value(2, 0, Value::number(3.0));
  s.set_cell_value(3, 0, Value::number(4.0));
  s.set_cell_value(4, 0, Value::number(5.0));
  s.set_cell_value(0, 1, Value::number(2.0));
  s.set_cell_value(1, 1, Value::number(4.0));
  s.set_cell_value(2, 1, Value::number(5.0));
  s.set_cell_value(3, 1, Value::number(4.0));
  s.set_cell_value(4, 1, Value::number(5.0));
  return wb;
}

// ---------------------------------------------------------------------------
// CORREL
// ---------------------------------------------------------------------------

TEST(RegressionCORREL, ReferenceDataset) {
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=CORREL(B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 6.0 / std::sqrt(60.0), 1e-12);
}

TEST(RegressionCORREL, PerfectPositive) {
  // y_i = 2 * x_i -> CORREL = 1.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=CORREL(A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(RegressionCORREL, PerfectNegative) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(6.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=CORREL(A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -1.0, 1e-12);
}

TEST(RegressionCORREL, ShapeMismatchIsNA) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=CORREL(A1:A2, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(RegressionCORREL, ZeroVarianceXIsDiv0) {
  // All x values identical: sum_xx == 0 -> #DIV/0!.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(5.0));
  const Value v = EvalSourceIn("=CORREL(B1:B3, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(RegressionCORREL, SinglePairIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=CORREL(B1:B1, A1:A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(RegressionCORREL, ErrorInEitherCellPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(5.0));
  const Value v = EvalSourceIn("=CORREL(B1:B3, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(RegressionCORREL, SkipsTextPair) {
  // Rows 1 and 3 survive; they form a line through the origin so
  // CORREL of the surviving two points is 1.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=CORREL(A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(RegressionCORREL, ArityUnder) {
  const Value v = EvalSource("=CORREL({1,2,3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(RegressionCORREL, ArityOver) {
  const Value v = EvalSource("=CORREL({1,2,3}, {2,4,6}, {3,6,9})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// COVARIANCE.P / COVARIANCE.S
// ---------------------------------------------------------------------------

TEST(RegressionCOVARIANCE, PopulationReference) {
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=COVARIANCE.P(B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.2, 1e-12);
}

TEST(RegressionCOVARIANCE, SampleReference) {
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=COVARIANCE.S(B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.5, 1e-12);
}

TEST(RegressionCOVARIANCE, SampleOnSinglePairIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=COVARIANCE.S(A1:A1, B1:B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(RegressionCOVARIANCE, PopulationOnSinglePairIsZero) {
  // COVARIANCE.P(n=1) is defined as 0/1 = 0. Only an empty sequence
  // (after text-pair dropping) yields #DIV/0!.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=COVARIANCE.P(A1:A1, B1:B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(RegressionCOVARIANCE, PopulationEmptyPairsIsDiv0) {
  // Every pair dropped because one side is text -> n == 0 -> #DIV/0!.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  const Value v = EvalSourceIn("=COVARIANCE.P(B1:B2, A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(RegressionCOVARIANCE, ShapeMismatchIsNA) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=COVARIANCE.P(A1:A2, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// SLOPE / INTERCEPT
// ---------------------------------------------------------------------------

TEST(RegressionSLOPE, ReferenceDataset) {
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=SLOPE(B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.6, 1e-12);
}

TEST(RegressionSLOPE, ArgumentOrderMatters) {
  // Swapping arrays gives a different slope (inverse regression).
  // With the reference dataset: SLOPE(A on B) = sum_xy / sum_yy(B) =
  // 6 / 6 = 1.0.
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=SLOPE(A1:A5, B1:B5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(RegressionSLOPE, ZeroVarianceXIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(7.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(7.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(7.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=SLOPE(B1:B3, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(RegressionINTERCEPT, ReferenceDataset) {
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=INTERCEPT(B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.2, 1e-12);
}

TEST(RegressionINTERCEPT, ZeroVarianceXIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(7.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(7.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=INTERCEPT(B1:B2, A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// RSQ
// ---------------------------------------------------------------------------

TEST(RegressionRSQ, ReferenceDataset) {
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=RSQ(B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.6, 1e-12);
}

TEST(RegressionRSQ, EqualsCorrelSquared) {
  const Workbook wb = MakeReferenceWorkbook();
  // Compute the difference in a single formula so identity holds with
  // respect to the concrete double produced by the evaluator.
  const Value v = EvalSourceIn("=RSQ(B1:B5,A1:A5) - CORREL(B1:B5,A1:A5)^2", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-14);
}

TEST(RegressionRSQ, ZeroVarianceIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  const Value v = EvalSourceIn("=RSQ(B1:B2, A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// FORECAST.LINEAR / FORECAST
// ---------------------------------------------------------------------------

TEST(RegressionFORECAST, AtXEqualsZero) {
  // FORECAST at x=0 equals the intercept (2.2 for the reference set).
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=FORECAST.LINEAR(0, B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.2, 1e-12);
}

TEST(RegressionFORECAST, AtXEqualsThree) {
  // slope*3 + intercept = 0.6*3 + 2.2 = 4.0.
  const Workbook wb = MakeReferenceWorkbook();
  const Value v = EvalSourceIn("=FORECAST.LINEAR(3, B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 4.0, 1e-12);
}

TEST(RegressionFORECAST, LegacyAliasMatchesLinear) {
  const Workbook wb = MakeReferenceWorkbook();
  const Value a = EvalSourceIn("=FORECAST(3, B1:B5, A1:A5)", wb, wb.sheet(0));
  const Value b = EvalSourceIn("=FORECAST.LINEAR(3, B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(RegressionFORECAST, ErrorInScalarArgPropagates) {
  // An error in the scalar x argument must take precedence over any
  // shape or DIV/0 condition in the regression.
  const Value v = EvalSource("=FORECAST.LINEAR(#DIV/0!, {1,2,3}, {2,4,6})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(RegressionFORECAST, Arity) {
  const Value under = EvalSource("=FORECAST.LINEAR(1, {1,2,3})");
  ASSERT_TRUE(under.is_error());
  EXPECT_EQ(under.as_error(), ErrorCode::Value);
  const Value over = EvalSource("=FORECAST.LINEAR(1, {1,2,3}, {1,2,3}, {1,2,3})");
  ASSERT_TRUE(over.is_error());
  EXPECT_EQ(over.as_error(), ErrorCode::Value);
}

TEST(RegressionFORECAST, ShapeMismatchInRegressionIsNA) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=FORECAST.LINEAR(1, A1:A2, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// ArrayLiteral source (smoke test for the ArrayLiteral code path)
// ---------------------------------------------------------------------------

TEST(RegressionArrayLiteral, CorrelPerfectPositive) {
  const Value v = EvalSource("=CORREL({1,2,3}, {2,4,6})");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(RegressionArrayLiteral, SlopeMatchesRange) {
  // Inline array literal should yield the same slope as the
  // equivalent range expression.
  const Value v = EvalSource("=SLOPE({2,4,5,4,5}, {1,2,3,4,5})");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.6, 1e-12);
}

// ---------------------------------------------------------------------------
// Scalar / non-array argument rejection
// ---------------------------------------------------------------------------

TEST(RegressionCORREL, ScalarArgIsNA) {
  // A bare scalar is not an array; regression family reports #N/A.
  const Value v = EvalSource("=CORREL(5, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
