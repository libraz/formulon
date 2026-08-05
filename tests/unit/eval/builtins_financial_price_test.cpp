//
// End-to-end tests for the regular-period bond-pricing builtin: PRICE.
// The implementation lives in `eval/builtins/financial_price.cpp` and
// shares the coupon-schedule engine in `eval/coupon_schedule.h`.

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

// ---------------------------------------------------------------------------
// PRICE -- canonical Microsoft documented case.
// ---------------------------------------------------------------------------

TEST(FinancialPrice, MicrosoftDocCanonicalCase) {
  // Microsoft's canonical example (Excel docs, "PRICE function"):
  //   settlement = 2008-02-15, maturity = 2017-11-15,
  //   rate = 5.75%, yld = 6.5%, redemption = 100,
  //   frequency = 2 (semi-annual), basis = 0 (US 30/360)
  // Expected: 94.6343616213 (matches MS reference & manual closed-form).
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 94.6343616213, 1e-9);
}

TEST(FinancialPrice, DefaultBasisZero) {
  // basis omitted defaults to 0 (US 30/360); same result as the explicit
  // basis=0 case above.
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 94.6343616213, 1e-9);
}

TEST(FinancialPrice, RejectsDateSerialBeyondExcelCalendar) {
  const Value v = EvalSource("=PRICE(2958466, 2958467, 0.0575, 0.065, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, BasisOneActualActual) {
  // Settlement falls inside a coupon period under actual/actual day
  // counting, exercising the v^(t1+i) path with a non-trivial fractional
  // t1. PCD = 2010-04-01, NCD = 2010-10-01, settlement = 2010-05-15.
  //   actual days_bs = 44, actual period_days = 183, days_nc = 139
  //   t1 = 139/183 = 0.7595628..., n = 11 coupons, cf = 2.5, AI ~ 0.6011
  // Expected (computed from first principles): 95.4528335058.
  const Value v = EvalSource("=PRICE(DATE(2010,5,15), DATE(2015,10,1), 0.05, 0.06, 100, 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 95.4528335058, 1e-9);
}

TEST(FinancialPrice, SinglePeriodLinearDiscount) {
  // Settlement falls inside the last coupon period (n == 1), so the
  // simple-interest branch fires:
  //   PRICE = (redemption + cf) / (1 + t1*(yld/freq)) - AI
  // settlement = 2017-08-15, maturity = 2017-11-15, basis 0 (30/360):
  //   t1 = 90/180 = 0.5, cf = 2.875, AI = 1.4375
  //   PRICE = 102.875 / (1 + 0.5*0.0325) - 1.4375 = 99.7925123001
  const Value v = EvalSource("=PRICE(DATE(2017,8,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 99.7925123001, 1e-9);
}

TEST(FinancialPrice, PremiumWhenYieldBelowRate) {
  // yld < rate => price > 100 (bondholder pays premium for above-market
  // coupons). Same shape as the canonical case but yld swapped from 6.5%
  // to 5%. Expected: 105.7233426194.
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.05, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 100.0);
  EXPECT_NEAR(v.as_number(), 105.7233426194, 1e-9);
}

TEST(FinancialPrice, ZeroCouponBond) {
  // rate == 0 => CF_i == 0 for i < n-1 and CF_{n-1} = redemption. AI is
  // also 0 (no accrual on a zero-coupon). The formula collapses to
  //   PRICE = redemption * v^(t1+n-1)
  // For the canonical shape (n=20, t1=0.5, freq=2, yld=0.065):
  //   100 * (1/1.0325)^19.5 = 53.5974124569
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0, 0.065, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 53.5974124569, 1e-9);
}

TEST(FinancialPrice, ZeroYieldUndiscountedSum) {
  // yld == 0 => v == 1 and the dirty price is just redemption + n*cf.
  // For the canonical shape: 100 + 20*2.875 - 1.4375 = 156.0625.
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 156.0625, 1e-9);
}

// ---------------------------------------------------------------------------
// PRICE -- domain / validation errors.
// ---------------------------------------------------------------------------

TEST(FinancialPrice, SettlementEqualMaturityIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2017,11,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2018,1,1), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, FrequencyThreeIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 3, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, FrequencyZeroIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, BasisFiveIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, BasisNegativeIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, NegativeRateIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), -0.01, 0.065, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, NegativeYieldIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, -0.01, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, ZeroRedemptionIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 0, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPrice, NegativeRedemptionIsNum) {
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, -50, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// PRICE -- dispatcher / propagation behaviour.
// ---------------------------------------------------------------------------

TEST(FinancialPrice, ErrorInArgPropagates) {
  // `propagate_errors = true` short-circuits before the impl runs;
  // the first `#N/A` argument flows through unchanged.
  const Value v = EvalSource("=PRICE(NA(), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(FinancialPrice, TooFewArgsIsValue) {
  // arity 5 is below the min_arity of 6.
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialPrice, TooManyArgsIsValue) {
  // arity 8 exceeds max_arity of 7.
  const Value v = EvalSource("=PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
