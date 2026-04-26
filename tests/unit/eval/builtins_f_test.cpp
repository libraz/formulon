// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for Excel's Snedecor's F distribution family
// (F.DIST, F.DIST.RT, F.INV, F.INV.RT). All four are scalar-only and
// share `eval/stats/special_functions.{h,cpp}` with the T family; tests
// verify the numerical surface against SciPy-verified references, PDF
// vs CDF consistency via finite differences, and every documented
// `#NUM!` domain case.

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

TEST(BuiltinsFRegistry, AllNamesRegistered) {
  for (const char* name : {"F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

// ---------------------------------------------------------------------------
// F.DIST
// ---------------------------------------------------------------------------

TEST(BuiltinsFDist, CdfBasicValue) {
  // scipy.stats.f.cdf(2, 5, 10) = 0.8358050491002611.
  const Value v = EvalSource("=F.DIST(2, 5, 10, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.8358050491002611, 1e-10);
}

TEST(BuiltinsFDist, PdfBasicValue) {
  // scipy.stats.f.pdf(2, 5, 10) = 0.16200574218011515.
  const Value v = EvalSource("=F.DIST(2, 5, 10, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.16200574218011515, 1e-10);
}

TEST(BuiltinsFDist, CdfAtZeroIsZero) {
  const Value v = EvalSource("=F.DIST(0, 5, 10, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsFDist, PdfAtZeroD1GtTwoIsZero) {
  // For d1 > 2 the F PDF at 0 is exactly 0.
  const Value v = EvalSource("=F.DIST(0, 5, 10, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsFDist, PdfAtZeroD1EqTwoIsOne) {
  // For d1 == 2 the F PDF at 0 equals 1 exactly.
  const Value v = EvalSource("=F.DIST(0, 2, 10, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsFDist, PdfAtZeroD1EqOneIsNum) {
  // d1 == 1 gives a divergent PDF at x = 0; Excel surfaces #NUM!.
  const Value v = EvalSource("=F.DIST(0, 1, 10, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsFDist, NegativeXIsNum) {
  const Value v = EvalSource("=F.DIST(-1, 5, 10, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsFDist, D1ZeroIsNum) {
  const Value v = EvalSource("=F.DIST(2, 0, 10, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsFDist, D2ZeroIsNum) {
  const Value v = EvalSource("=F.DIST(2, 5, 0, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsFDist, NonIntegerDfFloored) {
  const Value a = EvalSource("=F.DIST(2, 5.9, 10.7, TRUE)");
  const Value b = EvalSource("=F.DIST(2, 5, 10, TRUE)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-15);
}

TEST(BuiltinsFDist, PdfMatchesFiniteDifferenceOfCdf) {
  const Value lo = EvalSource("=F.DIST(1.99999, 5, 10, TRUE)");
  const Value hi = EvalSource("=F.DIST(2.00001, 5, 10, TRUE)");
  const Value pdf = EvalSource("=F.DIST(2, 5, 10, FALSE)");
  ASSERT_TRUE(lo.is_number());
  ASSERT_TRUE(hi.is_number());
  ASSERT_TRUE(pdf.is_number());
  const double fd = (hi.as_number() - lo.as_number()) / (2.0 * 1e-5);
  EXPECT_NEAR(fd, pdf.as_number(), 1e-6);
}

// ---------------------------------------------------------------------------
// F.DIST.RT
// ---------------------------------------------------------------------------

TEST(BuiltinsFDistRt, BasicValue) {
  // scipy.stats.f.sf(2, 5, 10) = 0.1641949508997389.
  const Value v = EvalSource("=F.DIST.RT(2, 5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.1641949508997389, 1e-10);
}

TEST(BuiltinsFDistRt, SumsWithDistToOne) {
  const Value v = EvalSource("=F.DIST(2, 5, 10, TRUE) + F.DIST.RT(2, 5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsFDistRt, NegativeXIsNum) {
  const Value v = EvalSource("=F.DIST.RT(-1, 5, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsFDistRt, D1ZeroIsNum) {
  const Value v = EvalSource("=F.DIST.RT(2, 0, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// F.INV
// ---------------------------------------------------------------------------

TEST(BuiltinsFInv, QuantileNinetyFive) {
  // scipy.stats.f.ppf(0.95, 5, 10) = 3.3258345304130104.
  const Value v = EvalSource("=F.INV(0.95, 5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.3258345304130104, 1e-6);
}

TEST(BuiltinsFInv, Median) {
  // scipy.stats.f.ppf(0.5, 5, 10) = 0.93193316085104805.
  const Value v = EvalSource("=F.INV(0.5, 5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.93193316085104805, 1e-6);
}

TEST(BuiltinsFInv, PZeroIsLeftBoundary) {
  // The F distribution's support starts at 0, so F.INV(0, ...) is the
  // natural left-edge value and Mac Excel 365 returns 0.0 rather than
  // #NUM! (mirroring how F.INV.RT accepts p == 1 at its right edge).
  const Value v = EvalSource("=F.INV(0, 5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsFInv, POneIsNum) {
  const Value v = EvalSource("=F.INV(1, 5, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsFInv, D1ZeroIsNum) {
  const Value v = EvalSource("=F.INV(0.5, 0, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// Round-trip: F.INV(F.DIST(x, d1, d2, TRUE), d1, d2) ~ x within tolerance.
TEST(BuiltinsFInv, RoundTrip) {
  struct Case {
    double x;
    double d1;
    double d2;
  };
  const Case cases[] = {{0.5, 5, 10}, {1.0, 5, 10}, {2.0, 5, 10}, {3.5, 4, 8}, {1.25, 10, 20}};
  for (const Case& c : cases) {
    std::ostringstream cdf_formula;
    cdf_formula << "=F.DIST(" << c.x << ", " << c.d1 << ", " << c.d2 << ", TRUE)";
    const Value cdf = EvalSource(cdf_formula.str());
    ASSERT_TRUE(cdf.is_number()) << "x=" << c.x << " d1=" << c.d1 << " d2=" << c.d2;
    std::ostringstream inv_formula;
    inv_formula << std::setprecision(17) << "=F.INV(" << cdf.as_number() << ", " << c.d1 << ", " << c.d2 << ")";
    const Value back = EvalSource(inv_formula.str());
    ASSERT_TRUE(back.is_number());
    EXPECT_NEAR(back.as_number(), c.x, 1e-5) << "x=" << c.x << " d1=" << c.d1 << " d2=" << c.d2;
  }
}

// ---------------------------------------------------------------------------
// F.INV.RT
// ---------------------------------------------------------------------------

TEST(BuiltinsFInvRt, FivePercentQuantileMatchesLeftTail) {
  // F.INV.RT(0.05, 5, 10) == F.INV(0.95, 5, 10) (dual tail identity).
  const Value rt = EvalSource("=F.INV.RT(0.05, 5, 10)");
  const Value lt = EvalSource("=F.INV(0.95, 5, 10)");
  ASSERT_TRUE(rt.is_number());
  ASSERT_TRUE(lt.is_number());
  EXPECT_NEAR(rt.as_number(), lt.as_number(), 1e-6);
}

TEST(BuiltinsFInvRt, POneIsZero) {
  // Right-tail at p = 1 places all mass on the left; quantile is 0.
  const Value v = EvalSource("=F.INV.RT(1, 5, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsFInvRt, PZeroIsNum) {
  // p = 0 places all mass on +inf; surface #NUM!.
  const Value v = EvalSource("=F.INV.RT(0, 5, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsFInvRt, RoundTrip) {
  struct Case {
    double x;
    double d1;
    double d2;
  };
  const Case cases[] = {{1.0, 5, 10}, {2.0, 5, 10}, {3.0, 4, 8}};
  for (const Case& c : cases) {
    std::ostringstream rt_formula;
    rt_formula << "=F.DIST.RT(" << c.x << ", " << c.d1 << ", " << c.d2 << ")";
    const Value rt = EvalSource(rt_formula.str());
    ASSERT_TRUE(rt.is_number()) << "x=" << c.x << " d1=" << c.d1 << " d2=" << c.d2;
    std::ostringstream inv_formula;
    inv_formula << std::setprecision(17) << "=F.INV.RT(" << rt.as_number() << ", " << c.d1 << ", " << c.d2 << ")";
    const Value back = EvalSource(inv_formula.str());
    ASSERT_TRUE(back.is_number());
    EXPECT_NEAR(back.as_number(), c.x, 1e-5) << "x=" << c.x << " d1=" << c.d1 << " d2=" << c.d2;
  }
}

// ---------------------------------------------------------------------------
// Cross-function consistency
// ---------------------------------------------------------------------------

TEST(BuiltinsF, InvAndInvRtDual) {
  const Value a = EvalSource("=F.INV(0.3, 6, 12)");
  const Value b = EvalSource("=F.INV.RT(0.7, 6, 12)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-6);
}

TEST(BuiltinsF, ArityRejectsWrongCount) {
  // F.DIST needs 4 args.
  const Value a = EvalSource("=F.DIST(2, 5, 10)");
  ASSERT_TRUE(a.is_error());
  EXPECT_EQ(a.as_error(), ErrorCode::Value);
  // F.INV needs 3 args.
  const Value b = EvalSource("=F.INV(0.95, 5)");
  ASSERT_TRUE(b.is_error());
  EXPECT_EQ(b.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
