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

TEST(BuiltinsDistributionsRegistry, AllNamesRegistered) {
  for (const char* name : {"NORM.DIST", "NORM.S.DIST", "NORM.INV", "NORM.S.INV", "BINOM.DIST", "POISSON.DIST",
                           "EXPON.DIST", "BETA.DIST", "BETA.INV", "GAMMA", "GAMMALN", "GAMMALN.PRECISE", "GAMMA.DIST",
                           "GAMMA.INV", "WEIBULL.DIST", "LOGNORM.DIST", "LOGNORM.INV", "HYPGEOM.DIST"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

TEST(BuiltinsNormDist, StandardPdfAtZero) {
  // PDF of N(0, 1) at 0 is 1 / sqrt(2*pi) ~ 0.3989422804...
  const Value v = EvalSource("=NORM.DIST(0, 0, 1, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.39894228040143267, 1e-12);
}

TEST(BuiltinsNormDist, StandardCdfAtZero) {
  const Value v = EvalSource("=NORM.DIST(0, 0, 1, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(BuiltinsNormDist, StandardCdfAtOne) {
  const Value v = EvalSource("=NORM.DIST(1, 0, 1, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.8413447460685429, 1e-12);
}

TEST(BuiltinsNormDist, SymmetryIdentityCdf) {
  // P(X <= -x) + P(X <= x) == 1 for any standard normal x.
  const Value lo = EvalSource("=NORM.DIST(-1.25, 0, 1, TRUE)");
  const Value hi = EvalSource("=NORM.DIST(1.25, 0, 1, TRUE)");
  ASSERT_TRUE(lo.is_number());
  ASSERT_TRUE(hi.is_number());
  EXPECT_NEAR(lo.as_number() + hi.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsNormDist, ScaledDistribution) {
  // N(100, 15) CDF at 105: equivalent to NORM.S.DIST((105-100)/15) ~ 0.63056.
  const Value v = EvalSource("=NORM.DIST(105, 100, 15, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.6305586598182363, 1e-10);
}

TEST(BuiltinsNormDist, CumulativeNumericTrue) {
  // A non-zero numeric cumulative argument must coerce to TRUE.
  const Value v = EvalSource("=NORM.DIST(0, 0, 1, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(BuiltinsNormDist, NegativeSdIsNum) {
  const Value v = EvalSource("=NORM.DIST(0, 0, -1, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNormDist, ZeroSdIsNum) {
  const Value v = EvalSource("=NORM.DIST(0, 0, 0, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNormSDist, CdfAtZero) {
  const Value v = EvalSource("=NORM.S.DIST(0, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(BuiltinsNormSDist, CdfAtTwo) {
  // Standard-normal CDF at z=2 ~ 0.97724987.
  const Value v = EvalSource("=NORM.S.DIST(2, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.9772498680518208, 1e-12);
}

TEST(BuiltinsNormSDist, PdfAtZero) {
  const Value v = EvalSource("=NORM.S.DIST(0, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.39894228040143267, 1e-12);
}

TEST(BuiltinsNormInv, HalfIsZero) {
  const Value v = EvalSource("=NORM.INV(0.5, 0, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-9);
}

TEST(BuiltinsNormInv, Quantile95) {
  // 95th percentile of N(0, 1) ~ 1.6448536269514722.
  const Value v = EvalSource("=NORM.INV(0.95, 0, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.6448536269514722, 1e-9);
}

TEST(BuiltinsNormInv, Scaled975) {
  // 100 + 15 * Phi^{-1}(0.975). Phi^{-1}(0.975) = 1.9599639845400545,
  // so the exact double-precision result is ~129.39945976810081.
  const Value v = EvalSource("=NORM.INV(0.975, 100, 15)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 129.39945976810081, 1e-9);
}

TEST(BuiltinsNormInv, PEqualsZeroIsNum) {
  const Value v = EvalSource("=NORM.INV(0, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNormInv, PEqualsOneIsNum) {
  const Value v = EvalSource("=NORM.INV(1, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNormInv, NegativeSdIsNum) {
  const Value v = EvalSource("=NORM.INV(0.5, 0, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNormSInv, HalfIsZero) {
  const Value v = EvalSource("=NORM.S.INV(0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-9);
}

TEST(BuiltinsNormSInv, Quantile95) {
  const Value v = EvalSource("=NORM.S.INV(0.95)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.6448536269514722, 1e-9);
}

TEST(BuiltinsNormSInv, PZeroIsNum) {
  const Value v = EvalSource("=NORM.S.INV(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNormInv, RoundTripFromDist) {
  // Round-trip: NORM.INV(NORM.DIST(x, 0, 1, TRUE), 0, 1) ~= x. After the
  // Halley polish NORM.INV is accurate to ~1e-14, but the probability
  // must be re-serialised to text to be embedded in the formula, so we
  // use `std::setprecision(17)` to preserve the full double.
  const Value cdf = EvalSource("=NORM.DIST(1.25, 0, 1, TRUE)");
  ASSERT_TRUE(cdf.is_number());
  std::ostringstream oss;
  oss << std::setprecision(17) << cdf.as_number();
  const std::string formula = "=NORM.INV(" + oss.str() + ", 0, 1)";
  const Value back = EvalSource(formula);
  ASSERT_TRUE(back.is_number());
  EXPECT_NEAR(back.as_number(), 1.25, 1e-9);
}

TEST(BuiltinsBinomDist, PmfFairCoinFiveHeads) {
  // BINOM(10, 0.5) PMF at 5 = C(10, 5) / 2^10 = 252 / 1024 = 0.24609375.
  const Value v = EvalSource("=BINOM.DIST(5, 10, 0.5, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.24609375, 1e-12);
}

TEST(BuiltinsBinomDist, CdfFairCoinFiveOrFewer) {
  // P(X <= 5) for BINOM(10, 0.5) ~ 0.623046875.
  const Value v = EvalSource("=BINOM.DIST(5, 10, 0.5, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.623046875, 1e-12);
}

TEST(BuiltinsBinomDist, PmfZeroSuccesses) {
  // P(X = 0) for BINOM(10, 0.5) = 1/1024 ~ 0.0009765625.
  const Value v = EvalSource("=BINOM.DIST(0, 10, 0.5, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0009765625, 1e-12);
}

TEST(BuiltinsBinomDist, PmfAllSuccesses) {
  const Value v = EvalSource("=BINOM.DIST(10, 10, 0.5, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0009765625, 1e-12);
}

TEST(BuiltinsBinomDist, CdfFullRangeSumsToOne) {
  // BINOM.DIST(n, n, p, TRUE) covers the full support, so it must equal 1.
  const Value v = EvalSource("=BINOM.DIST(10, 10, 0.3, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsBinomDist, ProbZeroPmfOnZeroIsOne) {
  // prob == 0 boundary: PMF(0) = 1, PMF(k>0) = 0.
  const Value zero = EvalSource("=BINOM.DIST(0, 10, 0, FALSE)");
  ASSERT_TRUE(zero.is_number());
  EXPECT_DOUBLE_EQ(zero.as_number(), 1.0);
  const Value one = EvalSource("=BINOM.DIST(1, 10, 0, FALSE)");
  ASSERT_TRUE(one.is_number());
  EXPECT_DOUBLE_EQ(one.as_number(), 0.0);
}

TEST(BuiltinsBinomDist, ProbOnePmfOnNIsOne) {
  // prob == 1 boundary: PMF(n) = 1, PMF(k<n) = 0.
  const Value n = EvalSource("=BINOM.DIST(10, 10, 1, FALSE)");
  ASSERT_TRUE(n.is_number());
  EXPECT_DOUBLE_EQ(n.as_number(), 1.0);
  const Value below = EvalSource("=BINOM.DIST(9, 10, 1, FALSE)");
  ASSERT_TRUE(below.is_number());
  EXPECT_DOUBLE_EQ(below.as_number(), 0.0);
}

TEST(BuiltinsBinomDist, NumberGreaterThanTrialsIsNum) {
  const Value v = EvalSource("=BINOM.DIST(11, 10, 0.5, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBinomDist, NegativeNumberIsNum) {
  const Value v = EvalSource("=BINOM.DIST(-1, 10, 0.5, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBinomDist, ProbOutOfRangeIsNum) {
  const Value hi = EvalSource("=BINOM.DIST(5, 10, 1.5, FALSE)");
  ASSERT_TRUE(hi.is_error());
  EXPECT_EQ(hi.as_error(), ErrorCode::Num);
  const Value lo = EvalSource("=BINOM.DIST(5, 10, -0.1, FALSE)");
  ASSERT_TRUE(lo.is_error());
  EXPECT_EQ(lo.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPoissonDist, PmfMean3AtTwo) {
  // Poisson(3) PMF at 2 = 3^2 * e^-3 / 2! ~ 0.224042.
  const Value v = EvalSource("=POISSON.DIST(2, 3, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.22404180765538775, 1e-12);
}

TEST(BuiltinsPoissonDist, CdfMean3AtTwo) {
  const Value v = EvalSource("=POISSON.DIST(2, 3, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.42319008112684359, 1e-12);
}

TEST(BuiltinsPoissonDist, PmfZeroCount) {
  // Poisson(3) PMF at 0 = e^-3 ~ 0.04978706836786394.
  const Value v = EvalSource("=POISSON.DIST(0, 3, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.04978706836786394, 1e-12);
}

TEST(BuiltinsPoissonDist, CdfFarTailApproachesOne) {
  // For mean = 5, the CDF at mean + 10*sqrt(mean) ~ 27 should be
  // indistinguishable from 1 at 1e-12 precision.
  const Value v = EvalSource("=POISSON.DIST(27, 5, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsPoissonDist, NegativeXIsNum) {
  const Value v = EvalSource("=POISSON.DIST(-1, 3, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPoissonDist, ZeroMeanIsPointMass) {
  // Mac Excel 365: `mean == 0` is the degenerate Poisson concentrated at
  // k = 0. PMF(0) = 1, PMF(k > 0) = 0, CDF(k >= 0) = 1. We align with the
  // oracle rather than rejecting the edge case.
  {
    const Value v = EvalSource("=POISSON.DIST(0, 0, FALSE)");
    ASSERT_TRUE(v.is_number());
    EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
  }
  {
    const Value v = EvalSource("=POISSON.DIST(1, 0, FALSE)");
    ASSERT_TRUE(v.is_number());
    EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
  }
  {
    const Value v = EvalSource("=POISSON.DIST(3, 0, TRUE)");
    ASSERT_TRUE(v.is_number());
    EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
  }
}

TEST(BuiltinsPoissonDist, NegativeMeanIsNum) {
  const Value v = EvalSource("=POISSON.DIST(2, -1, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsDistributionBounds, PoissonCdfAnswersFarPastTheTermCeiling) {
  // Every term beyond a handful is negligible against mean = 1, so the
  // CDF is 1 to within double precision. The point is that it comes back
  // at all: this argument is 1e12 terms of summation.
  const Value v = EvalSource("=POISSON.DIST(1E12, 1, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsDistributionBounds, PoissonCdfClosedFormAgreesWithTheSummation) {
  // Straddle the ceiling where the implementation switches from summing
  // terms to the closed form: the two sides must differ by exactly the
  // one PMF term between them, or the switch-over is visible to callers.
  const Value below = EvalSource("=POISSON.DIST(1048575, 1048576, TRUE)");
  const Value above = EvalSource("=POISSON.DIST(1048576, 1048576, TRUE)");
  const Value term = EvalSource("=POISSON.DIST(1048576, 1048576, FALSE)");
  ASSERT_TRUE(below.is_number());
  ASSERT_TRUE(above.is_number());
  ASSERT_TRUE(term.is_number());
  EXPECT_NEAR(above.as_number() - below.as_number(), term.as_number(), 1e-9);
}

TEST(BuiltinsDistributionBounds, BinomCdfAnswersPastTheTermCeiling) {
  // 1e8 trials is a hundred times the summation ceiling, so this is the
  // closed form answering. The median of Binomial(n, 0.5) sits at n/2,
  // and the exact value there is 0.5 by symmetry for odd n.
  const Value v = EvalSource("=BINOM.DIST(50000000, 100000001, 0.5, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-6);
}

TEST(BuiltinsDistributionBounds, BinomCdfFullSupportIsExactlyOne) {
  // `number_s == trials` is the whole support at any magnitude and needs
  // no beta evaluation, so it answers regardless of the shape bound.
  const Value v = EvalSource("=BINOM.DIST(1E12, 1E12, 0.5, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsDistributionBounds, BinomCdfUnresolvableShapeIsNum) {
  // Balanced shapes this large cannot be held to the family's tolerance:
  // the closed form's prefactor cancels in log space and what comes back
  // is off by more than the oracle tolerance. Refused rather than
  // answered with a number we cannot stand behind -- and refused
  // promptly, rather than summed for hours.
  const Value v = EvalSource("=BINOM.DIST(5E11, 1E12, 0.5, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsDistributionBounds, BinomPmfIsUnaffectedByTheCeiling) {
  // The ceiling is on summed terms. A point query does one evaluation at
  // any magnitude, so it must keep answering.
  const Value v = EvalSource("=BINOM.DIST(1E12, 1E12, 0.5, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GE(v.as_number(), 0.0);
}

TEST(BuiltinsDistributionBounds, BinomCdfJustBelowTheCeilingStillSums) {
  // The boundary is exclusive: one term short of the ceiling is still
  // answered by summation, which is what keeps the ordinary range intact.
  // The full support sums to 1 only to within the round-off a million
  // added terms accumulate -- which is itself why the ceiling sits here
  // rather than further out.
  const Value v = EvalSource("=BINOM.DIST(1048575, 1048575, 0.5, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-8);
}

TEST(BuiltinsExponDist, PdfAtZero) {
  // Exponential(lambda=1) PDF at 0 = lambda = 1.
  const Value v = EvalSource("=EXPON.DIST(0, 1, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsExponDist, CdfAtZero) {
  const Value v = EvalSource("=EXPON.DIST(0, 1, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsExponDist, CdfAtLn2IsHalf) {
  // CDF at x = ln(2)/lambda is exactly 0.5 (median of the distribution).
  const Value v = EvalSource("=EXPON.DIST(LN(2), 1, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(BuiltinsExponDist, NegativeXIsNum) {
  const Value v = EvalSource("=EXPON.DIST(-1, 1, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsExponDist, ZeroLambdaIsNum) {
  const Value v = EvalSource("=EXPON.DIST(1, 0, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsExponDist, NegativeLambdaIsNum) {
  const Value v = EvalSource("=EXPON.DIST(1, -1, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsDistributions, NormDistAndSDistAgree) {
  // NORM.DIST with mean=0, sd=1 must match NORM.S.DIST for the same inputs.
  const Value a = EvalSource("=NORM.DIST(1.5, 0, 1, TRUE)");
  const Value b = EvalSource("=NORM.S.DIST(1.5, TRUE)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-15);
}

TEST(BuiltinsDistributions, NormInvAndSInvAgree) {
  const Value a = EvalSource("=NORM.INV(0.8, 0, 1)");
  const Value b = EvalSource("=NORM.S.INV(0.8)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-15);
}

TEST(BuiltinsDistributions, ArityRejectsFewerArgs) {
  // NORM.DIST requires 4 args; 3 args must surface #VALUE! from the
  // dispatcher's arity check.
  const Value v = EvalSource("=NORM.DIST(0, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDistributions, ArityRejectsExtraArgs) {
  const Value v = EvalSource("=NORM.S.INV(0.5, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
