// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for Excel's Student's t distribution family
// (T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T). These functions are
// scalar-only (no range expansion) and share
// `eval/stats/special_functions.{h,cpp}` with the F family; tests verify
// both the numerical surface (SciPy-verified reference values, PDF/CDF
// consistency, round-trips) and every documented `#NUM!` domain case.

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
// Arenas are reset on each call to keep one test's allocations from
// bleeding into the next.
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

TEST(BuiltinsTRegistry, AllNamesRegistered) {
  for (const char* name : {"T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

// ---------------------------------------------------------------------------
// T.DIST
// ---------------------------------------------------------------------------

TEST(BuiltinsTDist, CdfAtZeroIsHalf) {
  // Student's t is symmetric around 0; CDF at 0 is exactly 0.5.
  const Value v = EvalSource("=T.DIST(0, 5, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(BuiltinsTDist, CdfAtOneWithFiveDf) {
  // Reference: scipy.stats.t.cdf(1, 5) = 0.81839126617543867.
  const Value v = EvalSource("=T.DIST(1, 5, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.81839126617543867, 1e-10);
}

TEST(BuiltinsTDist, CdfNegativeXSymmetric) {
  // T.DIST(-1, 5, TRUE) + T.DIST(1, 5, TRUE) == 1 (symmetry).
  const Value lo = EvalSource("=T.DIST(-1, 5, TRUE)");
  const Value hi = EvalSource("=T.DIST(1, 5, TRUE)");
  ASSERT_TRUE(lo.is_number());
  ASSERT_TRUE(hi.is_number());
  EXPECT_NEAR(lo.as_number() + hi.as_number(), 1.0, 1e-12);
  EXPECT_NEAR(lo.as_number(), 1.0 - 0.81839126617543867, 1e-10);
}

TEST(BuiltinsTDist, PdfAtZeroWithFiveDf) {
  // Reference: scipy.stats.t.pdf(0, 5) = 0.3796066898...
  const Value v = EvalSource("=T.DIST(0, 5, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.3796066898224944, 1e-10);
}

TEST(BuiltinsTDist, PdfSymmetryAroundZero) {
  const Value lo = EvalSource("=T.DIST(-1, 5, FALSE)");
  const Value hi = EvalSource("=T.DIST(1, 5, FALSE)");
  ASSERT_TRUE(lo.is_number());
  ASSERT_TRUE(hi.is_number());
  EXPECT_NEAR(lo.as_number(), hi.as_number(), 1e-12);
}

TEST(BuiltinsTDist, NonIntegerDfIsFloored) {
  const Value a = EvalSource("=T.DIST(1, 5.9, TRUE)");
  const Value b = EvalSource("=T.DIST(1, 5, TRUE)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-15);
}

TEST(BuiltinsTDist, DfZeroIsNum) {
  const Value v = EvalSource("=T.DIST(0, 0, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// Finite-difference check: numerical derivative of CDF approximates PDF.
TEST(BuiltinsTDist, PdfMatchesFiniteDifferenceOfCdf) {
  const Value lo = EvalSource("=T.DIST(0.99999, 5, TRUE)");
  const Value hi = EvalSource("=T.DIST(1.00001, 5, TRUE)");
  const Value pdf = EvalSource("=T.DIST(1, 5, FALSE)");
  ASSERT_TRUE(lo.is_number());
  ASSERT_TRUE(hi.is_number());
  ASSERT_TRUE(pdf.is_number());
  const double fd = (hi.as_number() - lo.as_number()) / (2.0 * 1e-5);
  EXPECT_NEAR(fd, pdf.as_number(), 1e-6);
}

// ---------------------------------------------------------------------------
// T.DIST.RT
// ---------------------------------------------------------------------------

TEST(BuiltinsTDistRt, BasicValue) {
  // scipy.stats.t.sf(1.5, 10) = 0.08225366322...
  const Value v = EvalSource("=T.DIST.RT(1.5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.08225366322272008, 1e-10);
}

TEST(BuiltinsTDistRt, SumsWithCdfToOne) {
  const Value v = EvalSource("=T.DIST(1.5, 10, TRUE) + T.DIST.RT(1.5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsTDistRt, DfZeroIsNum) {
  const Value v = EvalSource("=T.DIST.RT(1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// T.DIST.2T
// ---------------------------------------------------------------------------

TEST(BuiltinsTDist2T, BasicValue) {
  // scipy.stats.t.sf(2, 10) * 2 = 0.073388034770740365.
  const Value v = EvalSource("=T.DIST.2T(2, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.073388034770740365, 1e-10);
}

TEST(BuiltinsTDist2T, EqualsTwoTimesRt) {
  const Value twoT = EvalSource("=T.DIST.2T(2, 10)");
  const Value twoRt = EvalSource("=2 * T.DIST.RT(2, 10)");
  ASSERT_TRUE(twoT.is_number());
  ASSERT_TRUE(twoRt.is_number());
  EXPECT_NEAR(twoT.as_number(), twoRt.as_number(), 1e-12);
}

TEST(BuiltinsTDist2T, NegativeXIsNum) {
  // T.DIST.2T requires x >= 0 (matches Excel 365).
  const Value v = EvalSource("=T.DIST.2T(-1, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// T.INV
// ---------------------------------------------------------------------------

TEST(BuiltinsTInv, QuantileNinetyFiveTenDf) {
  // scipy.stats.t.ppf(0.95, 10) = 1.812461122811676.
  const Value v = EvalSource("=T.INV(0.95, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.812461122811676, 1e-8);
}

TEST(BuiltinsTInv, MedianIsZero) {
  const Value v = EvalSource("=T.INV(0.5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsTInv, PZeroIsNum) {
  const Value v = EvalSource("=T.INV(0, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsTInv, POneIsNum) {
  const Value v = EvalSource("=T.INV(1, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsTInv, DfZeroIsNum) {
  const Value v = EvalSource("=T.INV(0.5, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// Round-trip: T.INV(T.DIST(x, df, TRUE), df) ~ x within tolerance.
TEST(BuiltinsTInv, RoundTrip) {
  struct Case {
    double x;
    double df;
  };
  const Case cases[] = {{-1.5, 3}, {-0.5, 5}, {0.5, 5}, {1.5, 10}, {2.0, 20}};
  for (const Case& c : cases) {
    std::ostringstream cdf_formula;
    cdf_formula << "=T.DIST(" << c.x << ", " << c.df << ", TRUE)";
    const Value cdf = EvalSource(cdf_formula.str());
    ASSERT_TRUE(cdf.is_number()) << "x=" << c.x << " df=" << c.df;
    std::ostringstream inv_formula;
    inv_formula << std::setprecision(17) << "=T.INV(" << cdf.as_number() << ", " << c.df << ")";
    const Value back = EvalSource(inv_formula.str());
    ASSERT_TRUE(back.is_number());
    EXPECT_NEAR(back.as_number(), c.x, 1e-6) << "x=" << c.x << " df=" << c.df;
  }
}

// ---------------------------------------------------------------------------
// T.INV.2T
// ---------------------------------------------------------------------------

TEST(BuiltinsTInv2T, FivePercentTenDf) {
  // scipy.stats.t.isf(0.025, 10) = 2.2281388519862748.
  const Value v = EvalSource("=T.INV.2T(0.05, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.2281388519862748, 1e-8);
}

TEST(BuiltinsTInv2T, EquivalentToOneSidedHalfP) {
  // T.INV.2T(p, df) == T.INV(1 - p/2, df).
  const Value a = EvalSource("=T.INV.2T(0.1, 15)");
  const Value b = EvalSource("=T.INV(0.95, 15)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-8);
}

TEST(BuiltinsTInv2T, ProbZeroIsNum) {
  const Value v = EvalSource("=T.INV.2T(0, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Cross-function consistency
// ---------------------------------------------------------------------------

TEST(BuiltinsT, ArityRejectsWrongCount) {
  // T.DIST needs 3 args.
  const Value a = EvalSource("=T.DIST(1, 5)");
  ASSERT_TRUE(a.is_error());
  EXPECT_EQ(a.as_error(), ErrorCode::Value);
  // T.INV needs 2 args.
  const Value b = EvalSource("=T.INV(0.5, 10, 1)");
  ASSERT_TRUE(b.is_error());
  EXPECT_EQ(b.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
