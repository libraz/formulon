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
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

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

TEST(RegressionFORECAST, ArrayXLiftsPerElement) {
  // Reference set: intercept 2.2, slope 0.6.
  const Workbook wb = MakeReferenceWorkbook();
  const Value at3 = EvalSourceIn("=FORECAST.LINEAR(3, B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(at3.is_number());
  const double y3 = at3.as_number();
  const Value at7 = EvalSourceIn("=FORECAST.LINEAR(7, B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(at7.is_number());
  const double y7 = at7.as_number();
  for (const char* fn : {"FORECAST", "FORECAST.LINEAR"}) {
    const Value v = EvalSourceIn("=" + std::string(fn) + "({3,7,\"x\",#N/A}, B1:B5, A1:A5)", wb, wb.sheet(0));
    ASSERT_TRUE(v.is_array()) << fn;
    ASSERT_EQ(v.as_array()->rows, 1U);
    ASSERT_EQ(v.as_array()->cols, 4U);
    const Value* c = v.as_array()->cells;
    EXPECT_EQ(c[0].as_number(), y3);
    EXPECT_EQ(c[1].as_number(), y7);
    EXPECT_EQ(c[2].as_error(), ErrorCode::Value);
    EXPECT_EQ(c[3].as_error(), ErrorCode::NA);
  }
}

TEST(RegressionFORECAST, RangeXSpills) {
  Workbook wb = MakeReferenceWorkbook();
  wb.sheet(0).set_cell_value(0, 2, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(7.0));
  const Value at3 = EvalSourceIn("=FORECAST.LINEAR(3, B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(at3.is_number());
  const double y3 = at3.as_number();
  const Value v = EvalSourceIn("=FORECAST.LINEAR(C1:C2, B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array()->rows, 2U);
  ASSERT_EQ(v.as_array()->cols, 1U);
  EXPECT_EQ(v.as_array()->cells[0].as_number(), y3);
}

TEST(RegressionFORECAST, ArrayXFitErrorPerElement) {
  // A zero-variance x set is #DIV/0! for every numeric element; an element
  // that is itself an error keeps its own error, as the scalar call does.
  const Value v = EvalSource("=FORECAST.LINEAR({1;#REF!}, {1,2,3}, {2,2,2})");
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array()->rows, 2U);
  EXPECT_EQ(v.as_array()->cells[0].as_error(), ErrorCode::Div0);
  EXPECT_EQ(v.as_array()->cells[1].as_error(), ErrorCode::Ref);
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

// Asserts a numeric result within `rel` relative tolerance of `expected`.
void ExpectRegRel(const char* formula, double expected, double rel) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_number()) << formula;
  EXPECT_NEAR(v.as_number(), expected, std::abs(expected) * rel) << formula;
}

void ExpectRegErr(const char* formula, ErrorCode code) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_error()) << formula;
  EXPECT_EQ(v.as_error(), code) << formula;
}

TEST(RegressionCovariance, ScalarArgumentsAreOneElementArrays) {
  ExpectRegRel("=COVARIANCE.P(1,2)", 0.0, 0.0);
  ExpectRegRel("=COVARIANCE.P(3,3)", 0.0, 0.0);
  ExpectRegRel("=COVAR(1,2)", 0.0, 0.0);
  ExpectRegErr("=COVARIANCE.S(1,2)", ErrorCode::Div0);
}

void ExpectRegExact(const char* formula, double expected) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_number()) << formula;
  EXPECT_EQ(v.as_number(), expected) << formula;
}

TEST(RegressionFamily, MatchesExcelValues) {
  constexpr const char* kY = "{2;4;5;4;5}";
  constexpr const char* kX = "{1;2;3;4;5}";
  ExpectRegExact((std::string("=CORREL(") + kY + "," + kX + ")").c_str(), 0.7745966692414833);
  ExpectRegExact((std::string("=PEARSON(") + kY + "," + kX + ")").c_str(), 0.7745966692414833);
  ExpectRegExact((std::string("=RSQ(") + kY + "," + kX + ")").c_str(), 0.6000000000000001);

  const std::string pair = "{3;6;7;8;6;5;4;2;1;9},{2.5;3.1;4.0;4.2;3.9;3.3;2.8;2.0;1.1;5.0})";
  ExpectRegExact(("=CORREL(" + pair).c_str(), 0.9740600040261604);
  ExpectRegExact(("=PEARSON(" + pair).c_str(), 0.9740600040261604);
  ExpectRegExact(("=RSQ(" + pair).c_str(), 0.9487928914434437);
  ExpectRegExact(("=SLOPE(" + pair).c_str(), 2.2045588358987303);
  ExpectRegExact(("=INTERCEPT(" + pair).c_str(), -1.9325426865169506);
  ExpectRegExact(("=STEYX(" + pair).c_str(), 0.6243509541009645);
  ExpectRegExact(("=COVARIANCE.S(" + pair).c_str(), 2.9122222222222223);
  ExpectRegExact(("=COVARIANCE.P(" + pair).c_str(), 2.621);
  ExpectRegExact(("=COVAR(" + pair).c_str(), 2.621);
  const std::string tail = ",{3;6;7;8;6;5;4;2;1;9},{2.5;3.1;4.0;4.2;3.9;3.3;2.8;2.0;1.1;5.0})";
  ExpectRegExact(("=FORECAST(2.5" + tail).c_str(), 3.578854403229875);
  ExpectRegExact(("=FORECAST.LINEAR(7" + tail).c_str(), 13.499369164774162);
  ExpectRegExact(("=FORECAST(-1" + tail).c_str(), -4.137101522415681);
}

TEST(RegressionFamily, NumericLiteralKeepsFifteenDigits) {
  ExpectRegExact("=3.8541666666666665", 3.85416666666666);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
