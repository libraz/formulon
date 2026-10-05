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

}  // namespace
}  // namespace eval
}  // namespace formulon
