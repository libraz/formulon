// Copyright 2026 libraz. Licensed under the MIT License.
//
// Direct unit tests for the regularized incomplete gamma helpers
// `p_gamma(a, x)` and `q_gamma(a, x)`. The CHISQ.* builtins lean on
// these for their CDF surface, but we pin numerical behaviour here
// against closed-form references (the exponential case a=1, and the
// identity P + Q == 1) so that a regression in the special-functions
// layer surfaces at this level instead of as a drift in the Excel
// distribution tests.

#include "eval/stats/special_functions.h"

#include <cmath>
#include <limits>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace stats {
namespace {

// P(a, x) + Q(a, x) should equal 1 to within a few ulps whenever both
// values are finite. The pairs below sample both the series branch
// (`x < a + 1`) and the continued-fraction branch (`x >= a + 1`).
TEST(SpecialFunctionsIdentity, PPlusQEqualsOne) {
  const double pairs[][2] = {
      {0.5, 0.1},   // series
      {0.5, 2.0},   // CF
      {1.0, 0.5},   // series
      {1.0, 3.0},   // CF
      {2.0, 1.0},   // series
      {2.0, 5.0},   // CF
      {5.0, 3.0},   // series
      {5.0, 10.0},  // CF
      {10.0, 8.0},  // series
      {10.0, 20.0}  // CF
  };
  for (const auto& pair : pairs) {
    const double a = pair[0];
    const double x = pair[1];
    const double p = p_gamma(a, x);
    const double q = q_gamma(a, x);
    EXPECT_NEAR(p + q, 1.0, 1e-13) << "a=" << a << " x=" << x;
  }
}

// For a == 1 the regularized incomplete gamma collapses to the
// exponential CDF / survival function:
//   P(1, x) = 1 - exp(-x)
//   Q(1, x) = exp(-x)
// Exercise both branches of the switch (x < 2 vs x >= 2).
TEST(SpecialFunctionsExponential, PGammaAtAEqualsOne) {
  for (double x : {0.0, 0.25, 1.0, 1.5, 2.0, 5.0, 10.0}) {
    const double expected_p = 1.0 - std::exp(-x);
    EXPECT_NEAR(p_gamma(1.0, x), expected_p, 1e-12) << "x=" << x;
    EXPECT_NEAR(q_gamma(1.0, x), std::exp(-x), 1e-12) << "x=" << x;
  }
}

// Reference values from the spec: p_gamma(1, 1) = 1 - 1/e ~ 0.6321...,
// q_gamma(1, 1) = 1/e ~ 0.3679...
TEST(SpecialFunctionsReference, PGammaAtOneOne) {
  EXPECT_NEAR(p_gamma(1.0, 1.0), 1.0 - 1.0 / std::exp(1.0), 1e-12);
  EXPECT_NEAR(q_gamma(1.0, 1.0), 1.0 / std::exp(1.0), 1e-12);
}

// Boundary at x == 0: the incomplete gamma is 0 (lower) and 1 (upper)
// for any positive a.
TEST(SpecialFunctionsBoundary, XZero) {
  for (double a : {0.5, 1.0, 2.5, 10.0}) {
    EXPECT_DOUBLE_EQ(p_gamma(a, 0.0), 0.0) << "a=" << a;
    EXPECT_DOUBLE_EQ(q_gamma(a, 0.0), 1.0) << "a=" << a;
  }
}

// Domain errors: a <= 0 or x < 0 surface as NaN, matching the
// documented contract. Callers translate NaN into `#NUM!`.
TEST(SpecialFunctionsDomain, ReturnsNaN) {
  EXPECT_TRUE(std::isnan(p_gamma(-1.0, 1.0)));
  EXPECT_TRUE(std::isnan(p_gamma(0.0, 1.0)));
  EXPECT_TRUE(std::isnan(p_gamma(1.0, -1.0)));
  EXPECT_TRUE(std::isnan(q_gamma(-1.0, 1.0)));
  EXPECT_TRUE(std::isnan(q_gamma(0.0, 1.0)));
  EXPECT_TRUE(std::isnan(q_gamma(1.0, -1.0)));
}

// Both sides of the x = a + 1 boundary must agree: compute
// p_gamma(a, a+1) via the series branch (x < a+1 is false at equality,
// so this lands in the CF branch) and directly compare to the CF
// evaluation that q_gamma takes. The identity P + Q == 1 already
// enforces this, but we pin it on a value that historically drifted
// in older Numerical Recipes implementations.
TEST(SpecialFunctionsBoundary, AtTransition) {
  const double a = 3.0;
  const double x = a + 1.0;  // exactly on the boundary
  const double p = p_gamma(a, x);
  const double q = q_gamma(a, x);
  EXPECT_GT(p, 0.0);
  EXPECT_LT(p, 1.0);
  EXPECT_NEAR(p + q, 1.0, 1e-14);
}

// Large-a smoke test: the series path is used when x < a + 1. Confirm
// the result is a valid probability and that P + Q == 1 remains
// accurate at higher degrees of freedom than any CHISQ.* test covers.
TEST(SpecialFunctionsLargeA, Stable) {
  const double a = 100.0;
  const double x = 50.0;  // series branch
  const double p = p_gamma(a, x);
  const double q = q_gamma(a, x);
  EXPECT_GT(p, 0.0);
  EXPECT_LT(p, 1.0);
  EXPECT_NEAR(p + q, 1.0, 1e-12);
}

}  // namespace
}  // namespace stats
}  // namespace eval
}  // namespace formulon
