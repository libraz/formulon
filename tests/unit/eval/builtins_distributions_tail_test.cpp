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

// Asserts a numeric result within `rel` relative tolerance of `expected`.
void ExpectTailRel(const char* formula, double expected, double rel) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_number()) << formula;
  EXPECT_NEAR(v.as_number(), expected, std::abs(expected) * rel) << formula;
}

void ExpectTailErr(const char* formula, ErrorCode code) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_error()) << formula;
  EXPECT_EQ(v.as_error(), code) << formula;
}

TEST(BuiltinsTDistZeroDf, PdfIsDiv0WhileTheRestStayNum) {
  ExpectTailErr("=T.DIST(1,0,FALSE)", ErrorCode::Div0);
  ExpectTailErr("=T.DIST(,,)", ErrorCode::Div0);
  ExpectTailErr("=T.DIST(1,0,TRUE)", ErrorCode::Num);
  ExpectTailErr("=T.DIST.2T(1,0)", ErrorCode::Num);
  ExpectTailErr("=T.DIST.RT(1,0)", ErrorCode::Num);
  ExpectTailErr("=T.INV(0.9,0)", ErrorCode::Num);
  ExpectTailErr("=T.INV.2T(0.5,0)", ErrorCode::Num);
}

TEST(BuiltinsBinomInvProbabilityBounds, EndpointsAreNum) {
  ExpectTailErr("=BINOM.INV(10,0,0.5)", ErrorCode::Num);
  ExpectTailErr("=BINOM.INV(10,1,0.5)", ErrorCode::Num);
  ExpectTailErr("=BINOM.INV(10,,0.5)", ErrorCode::Num);
  ExpectTailErr("=CRITBINOM(10,0,0.5)", ErrorCode::Num);
}

TEST(BuiltinsTInvPrecision, QuantilesNearTheCentreAndTheTails) {
  ExpectTailRel("=T.INV.2T(1.0000001,10)", -1.284989018215513e-07, 1e-8);
  ExpectTailRel("=T.INV.2T(1.5,10)", -0.6998120613124317, 1e-13);
  ExpectTailRel("=T.INV(1E-9,5)", -98.93722464836996, 1e-11);
  ExpectTailRel("=T.INV.2T(1E-20,10)", 274.8423853162235, 1e-11);
  ExpectTailErr("=T.INV.2T(1E6,10)", ErrorCode::Num);
  ExpectTailErr("=T.INV.2T(2,1E6)", ErrorCode::Num);
}

TEST(BuiltinsChisqPrecision, VeryLargeDegreesOfFreedom) {
  ExpectTailRel("=CHISQ.DIST.RT(1E10,1E10)", 0.4999981193680548, 1e-12);
  ExpectTailRel("=CHISQ.INV(CHISQ.DIST(1E10,1E10,TRUE),1E10)/1E10", 1.0, 1e-12);
  ExpectTailRel("=CHISQ.INV.RT(CHISQ.DIST.RT(1E10,1E10),1E10)/1E10", 1.0, 1e-12);
  ExpectTailRel("=CHISQ.DIST(10000,10000,TRUE)", 0.5018806340338173, 1e-12);
  ExpectTailRel("=CHISQ.INV(0.001,10000)", 9568.668495093969, 1e-12);
  ExpectTailRel("=CHISQ.INV.RT(1E-12,3)", 58.91975568320216, 1e-12);
  ExpectTailRel("=CHISQ.INV(0.95,4)", 9.487729036781154, 1e-12);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
