// Probability distribution tests grouped by distribution family.

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
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {
using formulon::test::EvalSource;

TEST(BuiltinsBetaInv, Symmetric) {
  // BETA(2, 2) is symmetric around 0.5, so the median is 0.5.
  const Value v = EvalSource("=BETA.INV(0.5, 2, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-9);
}

TEST(BuiltinsBetaInv, ScaledSupport) {
  // Median of BETA(2, 2) rescaled to [0, 10] is 5.
  const Value v = EvalSource("=BETA.INV(0.5, 2, 2, 0, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 5.0, 1e-9);
}

TEST(BuiltinsBetaInv, PZeroIsNum) {
  const Value v = EvalSource("=BETA.INV(0, 2, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBetaInv, POneIsNum) {
  const Value v = EvalSource("=BETA.INV(1, 2, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBetaInv, TailTinyProbabilityRecoversNearZeroAnswer) {
  // For p far below the absolute Newton tolerance, the inverter must
  // converge in x-space rather than terminating on |cdf(x)-p| < kTol.
  // The true root for BETA.INV(3.038194444441917e-24, 2, 2.5, 0, 1.2)
  // is ~1e-12 (i.e. y = 8.33e-13 in standard support, then rescaled
  // by B - A = 1.2).
  const Value v = EvalSource("=BETA.INV(3.038194444441917e-24, 2, 2.5, 0, 1.2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0e-12, 5e-8);
}

TEST(BuiltinsBetaInv, RoundTripTinyAtSupportBoundary) {
  // BETA.INV must round-trip BETA.DIST near the support boundary even
  // when the resulting CDF probability is extreme.
  const Value cdf = EvalSource("=BETA.DIST(1e-12, 2, 2.5, TRUE, 0, 1.2)");
  ASSERT_TRUE(cdf.is_number());
  std::ostringstream oss;
  oss << std::setprecision(17) << cdf.as_number();
  const std::string formula = "=BETA.INV(" + oss.str() + ", 2, 2.5, 0, 1.2)";
  const Value back = EvalSource(formula);
  ASSERT_TRUE(back.is_number());
  EXPECT_NEAR(back.as_number(), 1.0e-12, 5e-8);
}

TEST(BuiltinsBetaInv, ExtremeTailDoesNotClampToFixedLowerBound) {
  // BETA(1, 1) has CDF(x) = x, making this an exact probe of the inverter's
  // lower search bound rather than of the incomplete-beta approximation.
  const Value v = EvalSource("=BETA.INV(5.77e-151, 1, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 5.77e-151, 1e-160);
}

TEST(BuiltinsGamma, IntegerFactorial) {
  // Γ(5) = 4! = 24.
  const Value v = EvalSource("=GAMMA(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 24.0, 1e-10);
}

TEST(BuiltinsGamma, HalfIsSqrtPi) {
  // Γ(0.5) = sqrt(pi) ~ 1.7724538509055159.
  const Value v = EvalSource("=GAMMA(0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.7724538509055159, 1e-12);
}

TEST(BuiltinsGamma, NegativeNonIntegerIsFinite) {
  // Γ(-0.5) = -2 * sqrt(pi) ~ -3.544907701811032.
  const Value v = EvalSource("=GAMMA(-0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -3.544907701811032, 1e-12);
}

TEST(BuiltinsGamma, ZeroIsNum) {
  const Value v = EvalSource("=GAMMA(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGamma, NegativeIntegerIsNum) {
  const Value v = EvalSource("=GAMMA(-3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGammaln, LnFactorialOfTen) {
  // ln Γ(10) = ln(9!) = ln(362880) ~ 12.801827480081469.
  const Value v = EvalSource("=GAMMALN(10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 12.801827480081469, 1e-10);
}

TEST(BuiltinsGammalnPrecise, AgreesWithGammaln) {
  const Value a = EvalSource("=GAMMALN(7.5)");
  const Value b = EvalSource("=GAMMALN.PRECISE(7.5)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsGammaln, ZeroIsNum) {
  const Value v = EvalSource("=GAMMALN(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGammaln, NegativeIsNum) {
  const Value v = EvalSource("=GAMMALN(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGammaDist, CdfBasic) {
  // GAMMA(alpha=3, beta=2) CDF at x=2 ~ 0.080301 (scipy:
  // gammainc(3, 1) where x/beta = 2/2 = 1).
  const Value v = EvalSource("=GAMMA.DIST(2, 3, 2, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.080301, 1e-5);
}

TEST(BuiltinsGammaDist, PdfBasic) {
  // PDF of GAMMA(3, 2) at x=2: (1 / (2^3 * Γ(3))) * 2^2 * exp(-1)
  //   = (1 / 16) * 4 * e^-1 ~ 0.09196986029286058.
  const Value v = EvalSource("=GAMMA.DIST(2, 3, 2, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.09196986029286058, 1e-10);
}

TEST(BuiltinsGammaDist, CdfZero) {
  const Value v = EvalSource("=GAMMA.DIST(0, 3, 2, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsGammaDist, NegativeXIsNum) {
  const Value v = EvalSource("=GAMMA.DIST(-1, 3, 2, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGammaDist, NegativeAlphaIsNum) {
  const Value v = EvalSource("=GAMMA.DIST(1, -1, 2, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGammaDist, NegativeBetaIsNum) {
  const Value v = EvalSource("=GAMMA.DIST(1, 3, -2, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGammaInv, RoundTrip) {
  // GAMMA.INV(GAMMA.DIST(x, a, b, TRUE), a, b) == x.
  const Value cdf = EvalSource("=GAMMA.DIST(4, 3, 2, TRUE)");
  ASSERT_TRUE(cdf.is_number());
  std::ostringstream oss;
  oss << std::setprecision(17) << cdf.as_number();
  const std::string formula = "=GAMMA.INV(" + oss.str() + ", 3, 2)";
  const Value back = EvalSource(formula);
  ASSERT_TRUE(back.is_number());
  EXPECT_NEAR(back.as_number(), 4.0, 1e-6);
}

TEST(BuiltinsGammaInv, PZeroIsZero) {
  const Value v = EvalSource("=GAMMA.INV(0, 3, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsGammaInv, POneIsNum) {
  const Value v = EvalSource("=GAMMA.INV(1, 3, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGammaInv, NegativeAlphaIsNum) {
  const Value v = EvalSource("=GAMMA.INV(0.5, -1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsWeibullDist, CdfBasic) {
  // WEIBULL(alpha=20, beta=100) CDF at x=105 ~ 0.929581.
  const Value v = EvalSource("=WEIBULL.DIST(105, 20, 100, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.929581, 1e-5);
}

TEST(BuiltinsWeibullDist, PdfAtBoundaryIsZero) {
  // Mac Excel 365 returns exactly 0 for the Weibull PDF at x == 0
  // regardless of alpha (including the alpha == 1 exponential case where
  // the mathematical limit would be 1/beta = 1). Verified against the
  // oracle.
  const Value v = EvalSource("=WEIBULL.DIST(0, 1, 1, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsWeibullDist, PdfAtBoundarySubLinearIsZero) {
  // alpha < 1 would mathematically diverge at x == 0, but Excel's
  // "boundary is zero" rule still applies: no #NUM!, just 0.
  const Value v = EvalSource("=WEIBULL.DIST(0, 0.5, 1, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsWeibullDist, PdfAwayFromBoundary) {
  // Weibull PDF with alpha=2, beta=1, x=1: (2/1) * 1^1 * exp(-1) = 2/e.
  const Value v = EvalSource("=WEIBULL.DIST(1, 2, 1, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.0 / std::exp(1.0), 1e-12);
}

TEST(BuiltinsWeibullDist, CdfAtZero) {
  const Value v = EvalSource("=WEIBULL.DIST(0, 2, 1, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsWeibullDist, NegativeXIsNum) {
  const Value v = EvalSource("=WEIBULL.DIST(-1, 2, 3, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsWeibullDist, ZeroAlphaIsNum) {
  const Value v = EvalSource("=WEIBULL.DIST(1, 0, 2, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBetaDist, CdfStandardSupport) {
  // BETA(alpha=8, beta=10) CDF at 0.4 ~ 0.35949234 (numerical integration
  // of the Beta(8, 10) PDF from 0 to 0.4; also matches scipy's betainc).
  const Value v = EvalSource("=BETA.DIST(0.4, 8, 10, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.35949234293309229, 1e-10);
}

TEST(BuiltinsBetaDist, PdfStandardSupport) {
  // BETA(alpha=2, beta=2) PDF at 0.5 = 6 * 0.5 * 0.5 = 1.5.
  const Value v = EvalSource("=BETA.DIST(0.5, 2, 2, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.5, 1e-12);
}

TEST(BuiltinsBetaDist, ScaledSupportCdf) {
  // Rescaled BETA(8, 10) on [1, 3] at x=1.8: y = (1.8 - 1)/(3 - 1) = 0.4,
  // so the CDF matches the standard-support case at y = 0.4.
  const Value v = EvalSource("=BETA.DIST(1.8, 8, 10, TRUE, 1, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.35949234293309229, 1e-10);
}

TEST(BuiltinsBetaDist, ScaledSupportPdfJacobian) {
  // Rescaled BETA(2, 2) on [0, 2] at x=1 is PDF(0.5) / span = 1.5 / 2 = 0.75.
  const Value v = EvalSource("=BETA.DIST(1, 2, 2, FALSE, 0, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.75, 1e-12);
}

TEST(BuiltinsBetaDist, CdfLowerBound) {
  const Value v = EvalSource("=BETA.DIST(0, 2, 3, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsBetaDist, CdfUpperBound) {
  const Value v = EvalSource("=BETA.DIST(1, 2, 3, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsBetaDist, NegativeAlphaIsNum) {
  const Value v = EvalSource("=BETA.DIST(0.5, -1, 2, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBetaDist, XOutOfSupportIsNum) {
  const Value v = EvalSource("=BETA.DIST(1.5, 2, 3, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBetaDist, AGtBIsNum) {
  const Value v = EvalSource("=BETA.DIST(0.5, 2, 3, TRUE, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBetaInv, RoundTrip) {
  // BETA.INV(BETA.DIST(0.4, 8, 10, TRUE), 8, 10) should recover 0.4.
  const Value cdf = EvalSource("=BETA.DIST(0.4, 8, 10, TRUE)");
  ASSERT_TRUE(cdf.is_number());
  std::ostringstream oss;
  oss << std::setprecision(17) << cdf.as_number();
  const std::string formula = "=BETA.INV(" + oss.str() + ", 8, 10)";
  const Value back = EvalSource(formula);
  ASSERT_TRUE(back.is_number());
  EXPECT_NEAR(back.as_number(), 0.4, 1e-9);
}

// Asserts a numeric result within `rel` relative tolerance of `expected`.
void ExpectGdRel(const char* formula, double expected, double rel) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_number()) << formula;
  EXPECT_NEAR(v.as_number(), expected, std::abs(expected) * rel) << formula;
}

void ExpectGdErr(const char* formula, ErrorCode code) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_error()) << formula;
  EXPECT_EQ(v.as_error(), code) << formula;
}

TEST(BuiltinsGammaDist, PdfAtZeroIsNumUpToAlphaOne) {
  ExpectGdErr("=GAMMA.DIST(0, 0.5, 2, FALSE)", ErrorCode::Num);
  ExpectGdErr("=GAMMA.DIST(0, 1, 2, FALSE)", ErrorCode::Num);
  ExpectGdErr("=GAMMADIST(0, 1, 1, FALSE)", ErrorCode::Num);
  ExpectGdRel("=GAMMA.DIST(0, 1.5, 2, FALSE)", 0.0, 0.0);
}

TEST(BuiltinsGammaInv, SmallProbabilityRoundTrips) {
  ExpectGdRel("=GAMMA.INV(GAMMA.DIST(1E-5,9,2,TRUE),9,2)", 9.999999999999992e-06, 1e-12);
  ExpectGdRel("=GAMMA.INV(0.5,3,2)", 5.348120627447123, 1e-13);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
