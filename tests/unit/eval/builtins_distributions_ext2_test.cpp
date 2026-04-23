// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the probability / combinatoric distribution extension
// registered by `register_stats_builtins`: CONFIDENCE, CONFIDENCE.NORM,
// CONFIDENCE.T, BINOM.INV, CRITBINOM, FISHER, FISHERINV, GAUSS, PHI,
// NEGBINOM.DIST, NEGBINOMDIST, BINOM.DIST.RANGE.
//
// These functions are all scalar-only. Tests cover canonical values, the
// documented alias equivalences (CRITBINOM == BINOM.INV, NEGBINOMDIST
// delegates to NEGBINOM.DIST's PMF branch), and every documented `#NUM!` /
// `#DIV/0!` domain case.

#include <cmath>
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

// Parses `src` and evaluates it via the default function registry. Arenas
// are thread-local and reset per call to keep text payloads readable while
// the immediately following EXPECT_* runs.
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

TEST(BuiltinsDistributionsExt2Registry, AllNamesRegistered) {
  for (const char* name :
       {"CONFIDENCE", "CONFIDENCE.NORM", "CONFIDENCE.T", "BINOM.INV", "CRITBINOM", "BINOM.DIST.RANGE", "FISHER",
        "FISHERINV", "GAUSS", "PHI", "NEGBINOM.DIST", "NEGBINOMDIST"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

// ---------------------------------------------------------------------------
// CONFIDENCE / CONFIDENCE.NORM
// ---------------------------------------------------------------------------

TEST(BuiltinsConfidence, NormBasic) {
  // z_{0.975} * 20 / sqrt(100) = 1.959963984540054 * 20 / 10 = 3.919927969080107
  const Value v = EvalSource("=CONFIDENCE(0.05, 20, 100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.9199279690801083, 1e-8);
}

TEST(BuiltinsConfidence, NormDotMatchesConfidence) {
  const Value a = EvalSource("=CONFIDENCE(0.05, 20, 100)");
  const Value b = EvalSource("=CONFIDENCE.NORM(0.05, 20, 100)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsConfidence, AlphaZeroIsNum) {
  const Value v = EvalSource("=CONFIDENCE(0, 20, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsConfidence, AlphaOneIsNum) {
  const Value v = EvalSource("=CONFIDENCE(1, 20, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsConfidence, NegativeStdevIsNum) {
  const Value v = EvalSource("=CONFIDENCE(0.05, -1, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsConfidence, SizeBelowOneIsNum) {
  const Value v = EvalSource("=CONFIDENCE(0.05, 20, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// CONFIDENCE.T
// ---------------------------------------------------------------------------

TEST(BuiltinsConfidenceT, Basic) {
  // scipy.stats.t.ppf(0.975, 29) = 2.045229642132703. Half-width =
  // 2.045229642 * 20 / sqrt(30) ~= 7.4681227.
  const Value v = EvalSource("=CONFIDENCE.T(0.05, 20, 30)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 7.468122735162, 1e-8);
}

TEST(BuiltinsConfidenceT, SizeOneIsDiv0) {
  // df = size - 1 = 0 collapses the distribution.
  const Value v = EvalSource("=CONFIDENCE.T(0.05, 20, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsConfidenceT, AlphaOutOfRangeIsNum) {
  const Value v = EvalSource("=CONFIDENCE.T(-0.1, 20, 30)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// BINOM.INV / CRITBINOM
// ---------------------------------------------------------------------------

TEST(BuiltinsBinomInv, MedianFairCoin) {
  // P(X <= 50) under BINOM(100, 0.5) ~ 0.5398; P(X <= 49) ~ 0.4602. So the
  // smallest k with CDF(k) >= 0.5 is 50.
  const Value v = EvalSource("=BINOM.INV(100, 0.5, 0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsBinomInv, ThreeQuarterQuantile) {
  // P(X <= 53) under BINOM(100, 0.5) ~ 0.7579 >= 0.75; P(X <= 52) ~ 0.7106.
  const Value v = EvalSource("=BINOM.INV(100, 0.5, 0.75)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 53.0);
}

TEST(BuiltinsBinomInv, AlphaZeroReturnsZero) {
  const Value v = EvalSource("=BINOM.INV(10, 0.3, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsBinomInv, AlphaOneReturnsTrials) {
  const Value v = EvalSource("=BINOM.INV(10, 0.3, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsBinomInv, ProbOutOfRangeIsNum) {
  const Value v = EvalSource("=BINOM.INV(10, 1.5, 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBinomInv, CritBinomAlias) {
  const Value a = EvalSource("=BINOM.INV(100, 0.5, 0.75)");
  const Value b = EvalSource("=CRITBINOM(100, 0.5, 0.75)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

// ---------------------------------------------------------------------------
// FISHER / FISHERINV
// ---------------------------------------------------------------------------

TEST(BuiltinsFisher, ZeroMapsToZero) {
  const Value v = EvalSource("=FISHER(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsFisher, Half) {
  // 0.5 * ln(1.5 / 0.5) = 0.5 * ln(3) ~ 0.5493061443340548.
  const Value v = EvalSource("=FISHER(0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5493061443340548, 1e-12);
}

TEST(BuiltinsFisher, NegativeHalfIsOpposite) {
  const Value v = EvalSource("=FISHER(-0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -0.5493061443340548, 1e-12);
}

TEST(BuiltinsFisher, BoundaryIsNum) {
  const Value a = EvalSource("=FISHER(1)");
  const Value b = EvalSource("=FISHER(-1)");
  ASSERT_TRUE(a.is_error());
  ASSERT_TRUE(b.is_error());
  EXPECT_EQ(a.as_error(), ErrorCode::Num);
  EXPECT_EQ(b.as_error(), ErrorCode::Num);
}

TEST(BuiltinsFisherInv, RoundTripWithFisher) {
  // FISHERINV(FISHER(0.3)) == 0.3 to full precision.
  const Value v = EvalSource("=FISHERINV(FISHER(0.3))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.3, 1e-12);
}

TEST(BuiltinsFisherInv, ZeroMapsToZero) {
  const Value v = EvalSource("=FISHERINV(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// GAUSS / PHI
// ---------------------------------------------------------------------------

TEST(BuiltinsGauss, Zero) {
  const Value v = EvalSource("=GAUSS(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsGauss, One) {
  // NORM.S.DIST(1, TRUE) ~ 0.841344746; GAUSS(1) = 0.341344746.
  const Value v = EvalSource("=GAUSS(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.3413447460685429, 1e-9);
}

TEST(BuiltinsGauss, NegativeSymmetry) {
  const Value a = EvalSource("=GAUSS(1)");
  const Value b = EvalSource("=GAUSS(-1)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), -b.as_number(), 1e-12);
}

TEST(BuiltinsPhi, Zero) {
  // 1 / sqrt(2*pi) ~ 0.3989422804014327.
  const Value v = EvalSource("=PHI(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.3989422804014327, 1e-12);
}

TEST(BuiltinsPhi, One) {
  // exp(-0.5) / sqrt(2*pi) ~ 0.24197072451914337.
  const Value v = EvalSource("=PHI(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.24197072451914337, 1e-12);
}

// ---------------------------------------------------------------------------
// NEGBINOM.DIST / NEGBINOMDIST
// ---------------------------------------------------------------------------

TEST(BuiltinsNegBinomDist, PmfBasic) {
  // C(14, 4) * (0.25)^5 * (0.75)^10 = 1001 * 0.00097656 * 0.05631351 ~= 0.05504...
  const Value v = EvalSource("=NEGBINOM.DIST(10, 5, 0.25, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.05504866037517518, 1e-9);
}

TEST(BuiltinsNegBinomDist, CdfExceedsPmf) {
  const Value pmf = EvalSource("=NEGBINOM.DIST(10, 5, 0.25, FALSE)");
  const Value cdf = EvalSource("=NEGBINOM.DIST(10, 5, 0.25, TRUE)");
  ASSERT_TRUE(pmf.is_number());
  ASSERT_TRUE(cdf.is_number());
  EXPECT_GT(cdf.as_number(), pmf.as_number());
  EXPECT_LE(cdf.as_number(), 1.0);
}

TEST(BuiltinsNegBinomDist, LegacyMatchesPmf) {
  const Value a = EvalSource("=NEGBINOM.DIST(10, 5, 0.25, FALSE)");
  const Value b = EvalSource("=NEGBINOMDIST(10, 5, 0.25)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsNegBinomDist, ProbZeroIsNum) {
  const Value v = EvalSource("=NEGBINOM.DIST(10, 5, 0, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNegBinomDist, SuccessBelowOneIsNum) {
  const Value v = EvalSource("=NEGBINOM.DIST(10, 0, 0.25, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNegBinomDist, NegativeFailuresIsNum) {
  const Value v = EvalSource("=NEGBINOM.DIST(-1, 5, 0.25, FALSE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// BINOM.DIST.RANGE
// ---------------------------------------------------------------------------

TEST(BuiltinsBinomDistRange, FourArgBasic) {
  // BINOM(10, 0.5): PMF(4)+PMF(5)+PMF(6) = 210/1024 + 252/1024 + 210/1024
  // = 672/1024 = 0.65625.
  const Value v = EvalSource("=BINOM.DIST.RANGE(10, 0.5, 4, 6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.65625, 1e-12);
}

TEST(BuiltinsBinomDistRange, ThreeArgMatchesPmf) {
  // Single-point query: BINOM.DIST.RANGE(10, 0.5, 5) == BINOM.DIST(5, 10, 0.5, FALSE).
  const Value a = EvalSource("=BINOM.DIST.RANGE(10, 0.5, 5)");
  const Value b = EvalSource("=BINOM.DIST(5, 10, 0.5, FALSE)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsBinomDistRange, ReversedRangeIsNum) {
  // number_s2 < number_s violates the domain.
  const Value v = EvalSource("=BINOM.DIST.RANGE(10, 0.5, 6, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsBinomDistRange, SuccessBeyondTrialsIsNum) {
  const Value v = EvalSource("=BINOM.DIST.RANGE(10, 0.5, 11)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
