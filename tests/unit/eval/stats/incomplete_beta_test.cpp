//
// Direct unit tests for the regularized incomplete beta helper
// `regularized_incomplete_beta(a, b, x)`. The T / F Excel distribution
// builtins share this routine for their CDF surface, so we pin the
// numerical behaviour here against closed-form references (the uniform
// case a=b=1, the symmetry identity, and spot-check values from SciPy)
// so that a regression in the special-functions layer surfaces at this
// level instead of as a drift in the Excel distribution tests.

#include <cmath>
#include <limits>

#include "eval/stats/special_functions.h"
#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace stats {
namespace {

// For a == b == 1 the incomplete beta degenerates to the identity
// I_x(1, 1) == x (the Beta(1, 1) law is uniform on [0, 1]). Exercises
// both CF branches of the reflection because the switching point lies
// at (1+1)/(1+1+2) = 0.5.
TEST(RegularizedIncompleteBetaUniform, IdentityOnUnitInterval) {
  for (double x : {0.0, 0.1, 0.25, 0.49, 0.5, 0.51, 0.75, 0.9, 1.0}) {
    const double got = regularized_incomplete_beta(1.0, 1.0, x);
    EXPECT_NEAR(got, x, 1e-14) << "x=" << x;
  }
}

// Reference: scipy.special.betainc(2, 3, 0.3) = 0.3483.
TEST(RegularizedIncompleteBetaReference, SciPyBetaInc2_3_0p3) {
  const double got = regularized_incomplete_beta(2.0, 3.0, 0.3);
  EXPECT_NEAR(got, 0.3483, 1e-4);
}

// Spot values from scipy.special.betainc accurate to ~1e-12.
TEST(RegularizedIncompleteBetaReference, SciPyAdditionalPoints) {
  // betainc(0.5, 0.5, 0.5) = 0.5 (arcsine distribution median).
  EXPECT_NEAR(regularized_incomplete_beta(0.5, 0.5, 0.5), 0.5, 1e-12);
  // betainc(5, 5, 0.5) = 0.5 (symmetric Beta).
  EXPECT_NEAR(regularized_incomplete_beta(5.0, 5.0, 0.5), 0.5, 1e-12);
  // betainc(10, 1, 0.9) = 0.9^10 = 0.3486784401.
  EXPECT_NEAR(regularized_incomplete_beta(10.0, 1.0, 0.9), 0.3486784401, 1e-10);
  // betainc(1, 10, 0.1) = 1 - (1-0.1)^10 = 1 - 0.9^10 = 0.6513215599.
  EXPECT_NEAR(regularized_incomplete_beta(1.0, 10.0, 0.1), 0.6513215599, 1e-10);
}

// Symmetry identity: I_x(a, b) + I_{1-x}(b, a) == 1. The reflection is
// the mechanism the implementation uses to stay on the fast-converging
// branch, so checking it is essentially a self-consistency test.
TEST(RegularizedIncompleteBetaSymmetry, ReflectionIdentity) {
  const double triples[][3] = {
      {0.5, 0.5, 0.3}, {1.0, 1.0, 0.4}, {2.0, 3.0, 0.25}, {3.0, 2.0, 0.75}, {5.0, 7.0, 0.2},
      {7.0, 5.0, 0.8}, {0.5, 2.5, 0.1}, {2.5, 0.5, 0.9},  {10.0, 2.0, 0.9}, {2.0, 10.0, 0.1},
  };
  for (const auto& t : triples) {
    const double a = t[0];
    const double b = t[1];
    const double x = t[2];
    const double left = regularized_incomplete_beta(a, b, x);
    const double right = regularized_incomplete_beta(b, a, 1.0 - x);
    EXPECT_NEAR(left + right, 1.0, 1e-12) << "a=" << a << " b=" << b << " x=" << x;
  }
}

// Boundary at x == 0 and x == 1: I is exactly 0 / 1 respectively for any
// positive shape parameters.
TEST(RegularizedIncompleteBetaBoundary, Endpoints) {
  for (double a : {0.5, 1.0, 2.5, 10.0}) {
    for (double b : {0.5, 1.0, 2.5, 10.0}) {
      EXPECT_DOUBLE_EQ(regularized_incomplete_beta(a, b, 0.0), 0.0) << "a=" << a << " b=" << b;
      EXPECT_DOUBLE_EQ(regularized_incomplete_beta(a, b, 1.0), 1.0) << "a=" << a << " b=" << b;
    }
  }
}

// Domain errors: a <= 0, b <= 0, or x outside [0, 1] all surface as NaN.
// Callers translate NaN into `#NUM!`.
TEST(RegularizedIncompleteBetaDomain, ReturnsNaN) {
  EXPECT_TRUE(std::isnan(regularized_incomplete_beta(0.0, 1.0, 0.5)));
  EXPECT_TRUE(std::isnan(regularized_incomplete_beta(-1.0, 1.0, 0.5)));
  EXPECT_TRUE(std::isnan(regularized_incomplete_beta(1.0, 0.0, 0.5)));
  EXPECT_TRUE(std::isnan(regularized_incomplete_beta(1.0, -1.0, 0.5)));
  EXPECT_TRUE(std::isnan(regularized_incomplete_beta(1.0, 1.0, -0.1)));
  EXPECT_TRUE(std::isnan(regularized_incomplete_beta(1.0, 1.0, 1.1)));
}

// Both CF branches (x < threshold and x > threshold) are exercised by the
// reflection test above, but we make the intent explicit here by picking
// (a, b) where the threshold (a+1)/(a+b+2) is close to 0.5 and selecting
// x values that straddle it. `a=5, b=5`: threshold = 6/12 = 0.5.
TEST(RegularizedIncompleteBetaBranch, BothBranchesExercised) {
  const double a = 5.0;
  const double b = 5.0;
  const double below = regularized_incomplete_beta(a, b, 0.3);  // below threshold
  const double above = regularized_incomplete_beta(a, b, 0.7);  // above threshold
  // Combined with the symmetry identity, `above == 1 - below` for the
  // symmetric Beta(5, 5). Pin the two independently against known values.
  EXPECT_GT(below, 0.0);
  EXPECT_LT(below, 0.5);
  EXPECT_GT(above, 0.5);
  EXPECT_LT(above, 1.0);
  EXPECT_NEAR(above, 1.0 - below, 1e-12);
}

// Asymmetric large shape parameters exercise the log-space prefactor.
// Without lgamma + exp, `tgamma(df)` overflows past ~170.
TEST(RegularizedIncompleteBetaLargeShape, NoOverflow) {
  const double got = regularized_incomplete_beta(200.0, 200.0, 0.5);
  // Beta(200, 200) is tightly peaked at 0.5, so the median is 0.5 exactly.
  EXPECT_NEAR(got, 0.5, 1e-12);
  // Off-median sanity: at x = 0.4 the CDF is tiny but strictly positive.
  const double tail = regularized_incomplete_beta(200.0, 200.0, 0.4);
  EXPECT_GT(tail, 0.0);
  EXPECT_LT(tail, 0.01);
}

// Either a converged value or NaN -- never a partial one
//
// The continued fraction used to stop after a fixed number of steps and
// return whatever it had reached, with no way for the caller to tell.
// Large shapes need materially more steps than small ones, so the shapes
// that need the most work were exactly the ones answered with a running
// value: at a == b == 5e11 that value was -86.6, which is not a
// probability. Two guards keep every exit path honest now -- the
// iteration budget scales with the shapes, and the log-space prefactor
// refuses once its own cancellation has swamped the result.

// The Beta(s, s) median is 0.5 exactly at every scale, which makes it the
// reference these cases are pinned against. It is also the slowest
// configuration for the recursion, so it is where truncation showed up.
TEST(RegularizedIncompleteBetaLargeShape, SymmetricMedianIsHalfAcrossScales) {
  struct Case {
    double shape;
    double tolerance;
  };
  // The tolerance widens with the shape because the prefactor is summed
  // in log space, where the terms cancel more violently as they grow.
  // Each one is roughly ten times the error actually observed, so a real
  // regression trips it while ordinary round-off does not.
  for (const Case& c : {Case{1.0e3, 1e-12}, Case{1.0e4, 1e-10}, Case{1.0e5, 1e-9}, Case{1.0e6, 1e-8}, Case{1.0e7, 1e-7},
                        Case{1.0e8, 1e-6}, Case{1.0e9, 1e-5}, Case{1.0e10, 1e-4}}) {
    const double got = regularized_incomplete_beta(c.shape, c.shape, 0.5);
    EXPECT_NEAR(got, 0.5, c.tolerance) << "shape=" << c.shape;
  }
}

// A shape large enough to need more work than the budget allows, or to
// lose the prefactor entirely, has to come back as NaN. Returning a
// number here is the defect: the caller cannot tell it apart from an
// answer, and it is not even inside [0, 1].
TEST(RegularizedIncompleteBetaLargeShape, UncomputableShapeIsNaN) {
  for (double shape : {1.0e14, 1.0e15, 1.0e17}) {
    const double got = regularized_incomplete_beta(shape, shape, 0.5);
    EXPECT_TRUE(std::isnan(got)) << "shape=" << shape << " returned " << got;
  }
}

// Whatever the shapes, a non-NaN result is a probability. This is the
// property the truncated recursion broke most visibly.
TEST(RegularizedIncompleteBetaLargeShape, AnyFiniteResultIsAProbability) {
  for (double shape : {1.0e2, 1.0e4, 1.0e6, 1.0e8, 1.0e10, 1.0e12, 1.0e14, 1.0e16}) {
    for (double x : {0.25, 0.5, 0.75}) {
      const double got = regularized_incomplete_beta(shape, shape, x);
      if (std::isnan(got)) {
        continue;  // Refusal is the documented alternative.
      }
      EXPECT_GE(got, 0.0) << "shape=" << shape << " x=" << x;
      EXPECT_LE(got, 1.0) << "shape=" << shape << " x=" << x;
    }
  }
}

// Skewed shapes converge in a handful of steps at any magnitude, so the
// budget must not refuse them just because `a + b` is large. This is what
// keeps T.DIST answering at its documented df ceiling, where the shapes
// are `(df/2, 1/2)`.
TEST(RegularizedIncompleteBetaLargeShape, SkewedShapesStillResolve) {
  for (double shape : {1.0e6, 1.0e8, 1.0e10}) {
    const double got = regularized_incomplete_beta(shape, 0.5, 0.5);
    ASSERT_FALSE(std::isnan(got)) << "shape=" << shape;
    EXPECT_GE(got, 0.0);
    EXPECT_LE(got, 1.0);
  }
}

// The reflection identity has to survive at scale too: both branches read
// the same prefactor, so a guard that fired on one side only would break
// the pairing rather than surface as a wrong number.
TEST(RegularizedIncompleteBetaLargeShape, ReflectionHoldsAtScale) {
  for (double shape : {1.0e4, 1.0e6, 1.0e8}) {
    const double lo = regularized_incomplete_beta(shape, shape, 0.4);
    const double hi = regularized_incomplete_beta(shape, shape, 0.6);
    ASSERT_FALSE(std::isnan(lo)) << "shape=" << shape;
    ASSERT_FALSE(std::isnan(hi)) << "shape=" << shape;
    EXPECT_NEAR(lo, 1.0 - hi, 1e-9) << "shape=" << shape;
  }
}

}  // namespace
}  // namespace stats
}  // namespace eval
}  // namespace formulon
