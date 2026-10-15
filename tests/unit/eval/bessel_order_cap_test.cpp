// The BESSEL family accepts a truncated order up to 1E7 and returns #NUM!
// past it, for every x (Mac Excel 365 ja-JP).

#include <string>

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {
using formulon::test::EvalSource;

TEST(BesselOrderCap, OrderAtTheCapEvaluates) {
  for (const char* formula :
       {"=BESSELI(1,10000000)", "=BESSELJ(1,10000000)", "=BESSELJ(1,10000000.9)", "=BESSELK(1E7,10000000)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), 0.0) << formula;
  }
  const Value y = EvalSource("=BESSELY(1E7,10000000)");
  ASSERT_TRUE(y.is_number());
  EXPECT_NEAR(y.as_number(), -3.59611273479326e-03, 1e-12);
}

TEST(BesselOrderCap, OrderPastTheCapIsNum) {
  for (const char* formula : {"=BESSELI(1,10000001)", "=BESSELJ(1,10000001)", "=BESSELK(1E7,10000001)",
                              "=BESSELY(1E7,10000001)", "=BESSELI(1,1073741824)", "=BESSELJ(1,2147483648)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
