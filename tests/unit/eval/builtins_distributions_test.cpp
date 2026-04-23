// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the probability-distribution family registered by
// `register_stats_builtins`: NORM.DIST, NORM.S.DIST, NORM.INV, NORM.S.INV,
// BINOM.DIST, POISSON.DIST, and EXPON.DIST.
//
// These functions are scalar-only (no range expansion), so every case
// here exercises the tree walker's eager path: each argument is
// pre-evaluated, coerced via `coerce_to_number` / `coerce_to_bool`, and
// handed to the impl. Tests cover PDF/CDF symmetry identities,
// PDF/INV round-trips, tail summation, and every documented
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

TEST(BuiltinsDistributionsRegistry, AllNamesRegistered) {
  for (const char* name : {"NORM.DIST", "NORM.S.DIST", "NORM.INV", "NORM.S.INV", "BINOM.DIST", "POISSON.DIST",
                           "EXPON.DIST", "BETA.DIST", "BETA.INV", "GAMMA", "GAMMALN", "GAMMALN.PRECISE", "GAMMA.DIST",
                           "GAMMA.INV", "WEIBULL.DIST", "LOGNORM.DIST", "LOGNORM.INV", "HYPGEOM.DIST"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

// ---------------------------------------------------------------------------
// NORM.DIST
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// NORM.S.DIST
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// NORM.INV / NORM.S.INV
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// BINOM.DIST
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// POISSON.DIST
// ---------------------------------------------------------------------------

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

TEST(BuiltinsPoissonDist, ZeroMeanIsNum) {
  const Value v = EvalSource("=POISSON.DIST(0, 0, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPoissonDist, NegativeMeanIsNum) {
  const Value v = EvalSource("=POISSON.DIST(2, -1, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// EXPON.DIST
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Cross-function consistency
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// BETA.DIST
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// BETA.INV
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// GAMMA
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// GAMMALN / GAMMALN.PRECISE
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// GAMMA.DIST
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// GAMMA.INV
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// WEIBULL.DIST
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// LOGNORM.DIST
// ---------------------------------------------------------------------------

TEST(BuiltinsLognormDist, CdfBasic) {
  // LOGNORM(mean=3.5, sd=1.2) CDF at x=4 ~ 0.039083.
  const Value v = EvalSource("=LOGNORM.DIST(4, 3.5, 1.2, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.039083, 1e-5);
}

TEST(BuiltinsLognormDist, PdfBasic) {
  // LOGNORM(mean=0, sd=1) PDF at x=1 = 1/(1 * sqrt(2*pi)) ~ 0.398942.
  const Value v = EvalSource("=LOGNORM.DIST(1, 0, 1, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.39894228040143267, 1e-12);
}

TEST(BuiltinsLognormDist, ZeroXIsNum) {
  const Value v = EvalSource("=LOGNORM.DIST(0, 0, 1, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLognormDist, NegativeXIsNum) {
  const Value v = EvalSource("=LOGNORM.DIST(-1, 0, 1, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLognormDist, ZeroSdIsNum) {
  const Value v = EvalSource("=LOGNORM.DIST(1, 0, 0, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// LOGNORM.INV
// ---------------------------------------------------------------------------

TEST(BuiltinsLognormInv, Median) {
  // Median of LOGNORM(mean=0, sd=1) is exp(0) = 1.
  const Value v = EvalSource("=LOGNORM.INV(0.5, 0, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-9);
}

TEST(BuiltinsLognormInv, RoundTrip) {
  const Value cdf = EvalSource("=LOGNORM.DIST(2, 0, 1, TRUE)");
  ASSERT_TRUE(cdf.is_number());
  std::ostringstream oss;
  oss << std::setprecision(17) << cdf.as_number();
  const std::string formula = "=LOGNORM.INV(" + oss.str() + ", 0, 1)";
  const Value back = EvalSource(formula);
  ASSERT_TRUE(back.is_number());
  EXPECT_NEAR(back.as_number(), 2.0, 1e-9);
}

TEST(BuiltinsLognormInv, PZeroIsNum) {
  const Value v = EvalSource("=LOGNORM.INV(0, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLognormInv, POneIsNum) {
  const Value v = EvalSource("=LOGNORM.INV(1, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// HYPGEOM.DIST
// ---------------------------------------------------------------------------

TEST(BuiltinsHypgeomDist, PmfBasic) {
  // HYPGEOM(k=1, n=4, K=8, N=20) PMF = C(8,1)*C(12,3)/C(20,4)
  //   = 8 * 220 / 4845 = 1760/4845 ~ 0.36326109391124871.
  const Value v = EvalSource("=HYPGEOM.DIST(1, 4, 8, 20, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1760.0 / 4845.0, 1e-12);
}

TEST(BuiltinsHypgeomDist, CdfBasic) {
  // CDF at k=2 for (n=4, K=8, N=20) should equal sum of PMFs 0..2.
  const Value cdf = EvalSource("=HYPGEOM.DIST(2, 4, 8, 20, TRUE)");
  const Value sum = EvalSource(
      "=HYPGEOM.DIST(0, 4, 8, 20, FALSE) + HYPGEOM.DIST(1, 4, 8, 20, FALSE) + HYPGEOM.DIST(2, 4, 8, 20, FALSE)");
  ASSERT_TRUE(cdf.is_number());
  ASSERT_TRUE(sum.is_number());
  EXPECT_NEAR(cdf.as_number(), sum.as_number(), 1e-12);
}

TEST(BuiltinsHypgeomDist, FullSupportCdfIsOne) {
  // Summing PMFs over the full support must equal 1. For (n=4, K=8, N=20)
  // the support is k in [max(0, 4+8-20), min(4, 8)] = [0, 4].
  const Value v = EvalSource("=HYPGEOM.DIST(4, 4, 8, 20, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsHypgeomDist, PmfZeroSuccesses) {
  // PMF(0; n=4, K=8, N=20) = C(8, 0) * C(12, 4) / C(20, 4).
  // = 1 * 495 / 4845 ~ 0.10216718266.
  const Value v = EvalSource("=HYPGEOM.DIST(0, 4, 8, 20, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 495.0 / 4845.0, 1e-12);
}

TEST(BuiltinsHypgeomDist, NegativeKIsNum) {
  // Mac Excel 365 rejects k < 0 with #NUM! (not 0): a negative sample
  // count is treated as malformed input, not merely infeasible. Verified
  // against the oracle golden.
  const Value v = EvalSource("=HYPGEOM.DIST(-1, 4, 8, 20, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsHypgeomDist, SampleBiggerThanPopIsNum) {
  // Genuinely malformed: sample size > population. Excel surfaces #NUM!.
  const Value v = EvalSource("=HYPGEOM.DIST(0, 25, 8, 20, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsHypgeomDist, KBiggerThanSampleIsZero) {
  // k = 5 exceeds n = 4 (no way to draw 5 successes in a sample of 4).
  // Excel treats this as an infeasible-k case and returns 0, not #NUM!.
  const Value v = EvalSource("=HYPGEOM.DIST(5, 4, 8, 20, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsHypgeomDist, KBiggerThanKIsZero) {
  // k = 9 exceeds K = 8 (can't have more successes than the total number
  // of successes in the population). Out of support -> 0.
  const Value v = EvalSource("=HYPGEOM.DIST(9, 10, 8, 20, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsHypgeomDist, CdfBeyondSupportIsOne) {
  // CDF at k = 10 (beyond k_max = 4 for this configuration) is 1.0.
  const Value v = EvalSource("=HYPGEOM.DIST(10, 4, 8, 20, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsHypgeomDist, CdfNegativeKIsNum) {
  // Symmetric with the PMF case: k < 0 is malformed, not infeasible.
  const Value v = EvalSource("=HYPGEOM.DIST(-1, 4, 8, 20, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsHypgeomDist, BigKBiggerThanPopIsNum) {
  // K > N is a genuinely malformed arrangement (more successes in the
  // population than the population itself). Excel surfaces #NUM!.
  const Value v = EvalSource("=HYPGEOM.DIST(0, 4, 25, 20, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
