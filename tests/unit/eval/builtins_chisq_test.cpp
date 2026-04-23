// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for Excel's chi-squared distribution family
// (CHISQ.DIST, CHISQ.DIST.RT, CHISQ.INV, CHISQ.INV.RT). These functions
// are scalar-only (no range expansion) and share
// `eval/stats/special_functions.{h,cpp}` with the future T / F
// distribution batches; tests verify both the numerical surface
// (reference values, PDF/CDF consistency, round-trips) and every
// documented `#NUM!` domain case.

#include <cmath>
#include <iomanip>
#include <sstream>
#include <string>
#include <string_view>

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

// Parses `src` and evaluates it via the default function registry.
// Arenas are reset on each call.
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
// Registry pins
// ---------------------------------------------------------------------------

TEST(BuiltinsChisqRegistry, AllNamesRegistered) {
  for (const char* name : {"CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

// ---------------------------------------------------------------------------
// CHISQ.DIST
// ---------------------------------------------------------------------------

TEST(BuiltinsChisqDist, CdfAtFiveWithFourDf) {
  // Reference: scipy.stats.chi2.cdf(5, 4) ~ 0.7127025048...
  const Value v = EvalSource("=CHISQ.DIST(5, 4, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.7127025048163541, 1e-10);
}

TEST(BuiltinsChisqDist, CdfAtZeroIsZero) {
  const Value v = EvalSource("=CHISQ.DIST(0, 4, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsChisqDist, PdfAtZeroWithDfTwo) {
  // df = 2 is the unique case where the chi-squared PDF at 0 is finite
  // and exactly equal to 0.5 (the exponential distribution with rate 1/2).
  const Value v = EvalSource("=CHISQ.DIST(0, 2, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(BuiltinsChisqDist, PdfAtZeroWithLargeDfIsZero) {
  const Value v = EvalSource("=CHISQ.DIST(0, 5, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsChisqDist, PdfAtZeroWithDfOneIsNum) {
  // Chi-squared PDF at 0 with df = 1 diverges; Excel surfaces #NUM!.
  const Value v = EvalSource("=CHISQ.DIST(0, 1, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsChisqDist, CdfLargeXApproachesOne) {
  const Value v = EvalSource("=CHISQ.DIST(100, 10, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsChisqDist, PdfAtFiveWithFourDf) {
  // Closed-form reference: 5^(4/2-1) * e^(-5/2) / (2^(4/2) * Gamma(4/2))
  //                      = 5 * e^(-2.5) / 4
  //                      ~ 0.1026062482798735.
  const Value v = EvalSource("=CHISQ.DIST(5, 4, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.1026062482798735, 1e-10);
}

TEST(BuiltinsChisqDist, NegativeXIsNum) {
  const Value v = EvalSource("=CHISQ.DIST(-1, 4, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsChisqDist, DfZeroIsNum) {
  const Value v = EvalSource("=CHISQ.DIST(5, 0, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsChisqDist, DfAboveCapIsNum) {
  // Excel 365 rejects df above 1e10.
  const Value v = EvalSource("=CHISQ.DIST(5, 1E11, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsChisqDist, NonIntegerDfIsFloored) {
  // df = 4.7 floors to 4, so the CDF must match CHISQ.DIST(5, 4, TRUE).
  const Value a = EvalSource("=CHISQ.DIST(5, 4.7, TRUE)");
  const Value b = EvalSource("=CHISQ.DIST(5, 4, TRUE)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-15);
}

// Finite-difference check: numerical derivative of the CDF approximates
// the PDF. This catches sign errors and mismatched formulas between the
// two branches.
TEST(BuiltinsChisqDist, PdfMatchesFiniteDifferenceOfCdf) {
  const double x = 5.0;
  const double dx = 1e-5;
  const Value lo = EvalSource("=CHISQ.DIST(4.99999, 4, TRUE)");
  const Value hi = EvalSource("=CHISQ.DIST(5.00001, 4, TRUE)");
  const Value pdf = EvalSource("=CHISQ.DIST(5, 4, FALSE)");
  ASSERT_TRUE(lo.is_number());
  ASSERT_TRUE(hi.is_number());
  ASSERT_TRUE(pdf.is_number());
  const double fd = (hi.as_number() - lo.as_number()) / (2.0 * dx);
  EXPECT_NEAR(fd, pdf.as_number(), 1e-6) << "x=" << x;
}

// ---------------------------------------------------------------------------
// CHISQ.DIST.RT
// ---------------------------------------------------------------------------

TEST(BuiltinsChisqDistRt, BasicValue) {
  // 1 - CHISQ.DIST(5, 4, TRUE) ~ 0.2872974951836459.
  const Value v = EvalSource("=CHISQ.DIST.RT(5, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.2872974951836459, 1e-10);
}

TEST(BuiltinsChisqDistRt, SumsWithCdfToOne) {
  const Value v = EvalSource("=CHISQ.DIST(5, 4, TRUE) + CHISQ.DIST.RT(5, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsChisqDistRt, ZeroXIsOne) {
  const Value v = EvalSource("=CHISQ.DIST.RT(0, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsChisqDistRt, NegativeXIsNum) {
  const Value v = EvalSource("=CHISQ.DIST.RT(-1, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsChisqDistRt, DfZeroIsNum) {
  const Value v = EvalSource("=CHISQ.DIST.RT(5, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// CHISQ.INV
// ---------------------------------------------------------------------------

TEST(BuiltinsChisqInv, QuantileNinetyFiveFourDf) {
  // scipy.stats.chi2.ppf(0.95, 4) ~ 9.487729036781154.
  const Value v = EvalSource("=CHISQ.INV(0.95, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 9.487729036781154, 1e-8);
}

TEST(BuiltinsChisqInv, MedianFourDf) {
  // scipy.stats.chi2.ppf(0.5, 4) ~ 3.3566939800333233.
  const Value v = EvalSource("=CHISQ.INV(0.5, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.3566939800333233, 1e-8);
}

TEST(BuiltinsChisqInv, PZeroIsZero) {
  const Value v = EvalSource("=CHISQ.INV(0, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsChisqInv, POneIsNum) {
  const Value v = EvalSource("=CHISQ.INV(1, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsChisqInv, NegativeProbIsNum) {
  const Value v = EvalSource("=CHISQ.INV(-0.1, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsChisqInv, DfZeroIsNum) {
  const Value v = EvalSource("=CHISQ.INV(0.5, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// Round-trip: CHISQ.INV(CHISQ.DIST(x, df, TRUE), df) ~ x for several
// (x, df) pairs. We re-serialise the intermediate probability with
// setprecision(17) to preserve the full double.
TEST(BuiltinsChisqInv, RoundTrip) {
  struct Case {
    double x;
    double df;
  };
  const Case cases[] = {{1.5, 3}, {5.0, 4}, {9.488, 4}, {20.0, 10}, {3.0, 7}};
  for (const Case& c : cases) {
    std::ostringstream cdf_formula;
    cdf_formula << "=CHISQ.DIST(" << c.x << ", " << c.df << ", TRUE)";
    const Value cdf = EvalSource(cdf_formula.str());
    ASSERT_TRUE(cdf.is_number()) << "x=" << c.x << " df=" << c.df;
    std::ostringstream inv_formula;
    inv_formula << std::setprecision(17) << "=CHISQ.INV(" << cdf.as_number() << ", " << c.df << ")";
    const Value back = EvalSource(inv_formula.str());
    ASSERT_TRUE(back.is_number());
    EXPECT_NEAR(back.as_number(), c.x, 1e-7) << "x=" << c.x << " df=" << c.df;
  }
}

// ---------------------------------------------------------------------------
// CHISQ.INV.RT
// ---------------------------------------------------------------------------

TEST(BuiltinsChisqInvRt, Quantile5PctFourDf) {
  // CHISQ.INV.RT(0.05, 4) == CHISQ.INV(0.95, 4) ~ 9.487729036781154.
  const Value v = EvalSource("=CHISQ.INV.RT(0.05, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 9.487729036781154, 1e-8);
}

TEST(BuiltinsChisqInvRt, POneIsZero) {
  const Value v = EvalSource("=CHISQ.INV.RT(1, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsChisqInvRt, PZeroIsNum) {
  // Right-tail quantile at p=0 is +inf; surface #NUM!.
  const Value v = EvalSource("=CHISQ.INV.RT(0, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsChisqInvRt, RoundTrip) {
  // CHISQ.INV.RT(CHISQ.DIST.RT(x, df), df) ~ x for several (x, df).
  struct Case {
    double x;
    double df;
  };
  const Case cases[] = {{1.5, 3}, {5.0, 4}, {20.0, 10}};
  for (const Case& c : cases) {
    std::ostringstream rt_formula;
    rt_formula << "=CHISQ.DIST.RT(" << c.x << ", " << c.df << ")";
    const Value rt = EvalSource(rt_formula.str());
    ASSERT_TRUE(rt.is_number()) << "x=" << c.x << " df=" << c.df;
    std::ostringstream inv_formula;
    inv_formula << std::setprecision(17) << "=CHISQ.INV.RT(" << rt.as_number() << ", " << c.df << ")";
    const Value back = EvalSource(inv_formula.str());
    ASSERT_TRUE(back.is_number());
    EXPECT_NEAR(back.as_number(), c.x, 1e-7) << "x=" << c.x << " df=" << c.df;
  }
}

// ---------------------------------------------------------------------------
// Cross-function consistency
// ---------------------------------------------------------------------------

TEST(BuiltinsChisq, InvAndInvRtDual) {
  // CHISQ.INV(p, df) must equal CHISQ.INV.RT(1 - p, df) for any p in
  // (0, 1). Pick a mid-range value where both Newton iterations
  // converge well below our 1e-8 tolerance target.
  const Value a = EvalSource("=CHISQ.INV(0.3, 6)");
  const Value b = EvalSource("=CHISQ.INV.RT(0.7, 6)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-8);
}

TEST(BuiltinsChisq, ArityRejectsFewerArgs) {
  const Value v = EvalSource("=CHISQ.DIST(5, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChisq, ArityRejectsExtraArgs) {
  const Value v = EvalSource("=CHISQ.INV(0.5, 4, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
