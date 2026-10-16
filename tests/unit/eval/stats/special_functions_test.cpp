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

TEST(SpecialFunctionsLargeA, ConvergesAtDistributionScale) {
  // These shapes used to stop at the unconditional 200th term and return
  // the partial sum as if it were a probability. At the transition point,
  // both branches must now converge and remain complementary.
  for (double a : {1e3, 1e4, 1e5}) {
    const double p = p_gamma(a, a);
    const double q = q_gamma(a, a);
    EXPECT_TRUE(std::isfinite(p)) << "a=" << a;
    EXPECT_TRUE(std::isfinite(q)) << "a=" << a;
    EXPECT_GT(p, 0.0) << "a=" << a;
    EXPECT_LT(p, 1.0) << "a=" << a;
    EXPECT_NEAR(p + q, 1.0, 1e-12) << "a=" << a;
  }
}

// Reference values spanning both branches of the ERF/ERFC rationals, the 0.5
// and 4 breakpoints and both signs.
struct ErfPoint {
  double x;
  double erf;
  double erfc;
  double norm_s_dist;
};

const ErfPoint kErfPoints[] = {
    {-7.0, -1.0, 2.0, 1.2798125438858352e-12},
    {-4.500000000000053, -0.9999999998033839, 1.999999999803384, 3.397673124729209e-06},
    {-4.000000000000064, -0.9999999845827421, 1.999999984582742, 3.16712418331113e-05},
    {-3.0000000000000853, -0.9999779095030015, 1.9999779095030015, 0.001349898031629714},
    {-2.0000000000001066, -0.9953222650189549, 1.995322265018955, 0.022750131948173456},
    {-1.000000000000128, -0.842700792949768, 1.8427007929497679, 0.15865525393142607},
    {-0.5000000000001275, -0.5204998778131585, 1.5204998778131587, 0.30853753872594203},
    {-0.25000000000012723, -0.27632639016837185, 1.276326390168372, 0.40129367431702706},
    {-0.005000000000127077, -0.005641848820174939, 1.0056418488201748, 0.4980052969092085},
    {0.004999999999872923, 0.005641848819888163, 0.9943581511801118, 0.5019947030906902},
    {0.24999999999987307, 0.2763263901681024, 0.7236736098318977, 0.5987063256828746},
    {0.3999999999998732, 0.4283923550465466, 0.5716076449534534, 0.6554217416102774},
    {0.48999999999987326, 0.5116682611884107, 0.4883317388115893, 0.6879330505825647},
    {0.49999999999987327, 0.5204998778129353, 0.4795001221870647, 0.6914624612739686},
    {0.5049999999998732, 0.5248827877318527, 0.47511721226814724, 0.6932205820110979},
    {0.7499999999998734, 0.7111556336534337, 0.28884436634656624, 0.7733726476230937},
    {0.9999999999998737, 0.8427007929496624, 0.15729920705033756, 0.8413447460685124},
    {1.499999999999863, 0.9661051464752944, 0.033894853524705566, 0.9331927987311242},
    {1.9999999999998523, 0.9953222650189497, 0.004677734981050318, 0.9772498680518128},
    {2.9999999999998312, 0.9999779095030014, 2.209049699860894e-05, 0.9986501019683691},
    {3.99499999999981, 0.9999999839349666, 1.6065033432913357e-08, 0.999967652873523},
    {3.99999999999981, 0.9999999845827421, 1.541725790030415e-08, 0.9999683287581669},
    {4.00499999999981, 0.9999999852051181, 1.4794881905905571e-08, 0.9999689912594256},
    {4.999999999999789, 0.9999999999984626, 1.5374597944313477e-12, 0.9999997133484281},
    {5.999999999999767, 1.0, 2.1519736712559825e-17, 0.9999999990134123},
    {6.999999999999746, 1.0, 4.183825607794428e-23, 0.9999999999987201},
};

TEST(ErfKernel, MatchesExcelValues) {
  for (const ErfPoint& p : kErfPoints) {
    EXPECT_EQ(erf_excel(p.x), p.erf) << p.x;
    EXPECT_EQ(erfc_excel(p.x), p.erfc) << p.x;
    EXPECT_EQ(normal_cdf(p.x), p.norm_s_dist) << p.x;
  }
}

TEST(ErfKernel, ZeroAndFarTail) {
  EXPECT_EQ(erf_excel(0.0), 0.0);
  EXPECT_EQ(erfc_excel(0.0), 1.0);
  EXPECT_EQ(erfc_excel(101.0), 0.0);
  EXPECT_EQ(erfc_excel(-101.0), 2.0);
  EXPECT_EQ(erf_excel(101.0), 1.0);
  EXPECT_EQ(erf_excel(-101.0), -1.0);
  EXPECT_EQ(normal_cdf(-150.0), 0.0);
  EXPECT_EQ(normal_cdf(0.0), 0.5);
  EXPECT_TRUE(std::isnan(erf_excel(std::numeric_limits<double>::quiet_NaN())));
  EXPECT_TRUE(std::isnan(erfc_excel(std::numeric_limits<double>::quiet_NaN())));
}

// Reference values across the tail, central, middle and outer regions and the
// 0.075 / 0.9375 breakpoints.
struct QuantilePoint {
  double p;
  double z;
};

const QuantilePoint kQuantilePoints[] = {
    {7.466108948025751e-301, -37.05497728591007},
    {6.776263578034403e-21, -9.303791061748415},
    {5.820766091346741e-11, -6.4439345265385635},
    {1.52587890625e-05, -4.169569323349103},
    {0.0009765625, -3.097269078198784},
    {0.010009765625, -2.325981619598469},
    {0.02490234375, -1.9616376339009494},
    {0.050048828125, -1.6443803751786208},
    {0.070068359375, -1.4752820926704961},
    {0.075, -1.4395314709384572},
    {0.07503814697265626, -1.439262038830567},
    {0.10009765625, -1.2809953126722822},
    {0.14990234375, -1.0368523205641345},
    {0.199951171875, -0.8417956562558121},
    {0.300048828125, -0.5242600831294301},
    {0.39990234375, -0.2535998823915192},
    {0.5, 0.0},
    {0.60009765625, 0.2535998823915192},
    {0.699951171875, 0.5242600831294301},
    {0.85009765625, 1.0368523205641345},
    {0.89990234375, 1.2809953126722822},
    {0.929931640625, 1.4752820926704973},
    {0.9375, 1.5341205443525465},
    {0.9375381469726562, 1.5344307966549564},
    {0.949951171875, 1.6443803751786208},
    {0.989990234375, 2.325981619598469},
    {0.9990234375, 3.097269078198784},
    {0.999755859375, 3.4871041041144313},
};

TEST(NormalQuantile, MatchesExcelValues) {
  for (const QuantilePoint& q : kQuantilePoints) {
    EXPECT_EQ(normal_quantile(q.p), q.z) << q.p;
  }
}

}  // namespace
}  // namespace stats
}  // namespace eval
}  // namespace formulon
