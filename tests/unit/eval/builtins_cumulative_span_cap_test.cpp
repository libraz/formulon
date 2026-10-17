// Work caps on the cumulative-term financial functions, and the HYPGEOM.DIST count bound.

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

TEST(HypgeomDistCumulative, CountAtInt32LimitIsNum) {
  for (const char* src :
       {"=HYPGEOM.DIST(5,10,50,2147483647,TRUE)", "=HYPGEOM.DIST(5,10,50,2147483647,FALSE)",
        "=HYPGEOM.DIST(2147483647,10,50,100,TRUE)", "=HYPGEOM.DIST(5,2147483647,50,2147483647,TRUE)",
        "=HYPGEOM.DIST(5,10,2147483647,2147483647,TRUE)", "=HYPGEOM.DIST(1E12,2E12,2E12,4E12,TRUE)"}) {
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_error()) << src;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << src;
  }
}

TEST(HypgeomDistCumulative, CountBelowInt32LimitAnswers) {
  const Value v = EvalSource("=HYPGEOM.DIST(5,10,50,2147483646,TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
  const Value w = EvalSource("=HYPGEOM.DIST(2147483646,10,50,100,TRUE)");
  ASSERT_TRUE(w.is_number());
  EXPECT_DOUBLE_EQ(w.as_number(), 1.0);
}

TEST(HypgeomDistCumulative, LongSpanIsWalkedWithoutCap) {
  const Value v = EvalSource("=HYPGEOM.DIST(5000000,10000000,50000000,100000000,TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.5001329807567633);
}

TEST(HypgeomDistCumulative, SmallSpanStillSums) {
  // Population 20 with 8 successes, sample 4: CDF at the support top is 1.
  const Value top = EvalSource("=HYPGEOM.DIST(4,4,8,20,TRUE)");
  ASSERT_TRUE(top.is_number());
  EXPECT_DOUBLE_EQ(top.as_number(), 1.0);
  const Value mid = EvalSource("=HYPGEOM.DIST(2,4,8,20,TRUE)");
  ASSERT_TRUE(mid.is_number());
  EXPECT_GT(mid.as_number(), 0.0);
  EXPECT_LT(mid.as_number(), 1.0);
}

TEST(FinancialCumulative, CumipmtHugeSpanIsNum) {
  const Value v = EvalSource("=CUMIPMT(0.01,1E15,1000,1,1E15,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCumulative, CumprincHugeSpanIsNum) {
  const Value v = EvalSource("=CUMPRINC(0.01,1E15,1000,1,1E15,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCumulative, NormalSpanUnchanged) {
  const Value v = EvalSource("=CUMIPMT(0.09/12,360,125000,1,1,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -937.5, 1e-9);
}

TEST(ComplexInspectors, ImabsOverflowIsNum) {
  const Value v = EvalSource("=IMABS(\"1.7e308+1.7e308i\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
