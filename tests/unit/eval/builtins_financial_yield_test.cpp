//
// End-to-end tests for the regular-period bond yield-to-maturity
// builtin: YIELD. The implementation lives in
// `eval/builtins/financial_yield.cpp` and shares the clean-price kernel
// in `eval/builtins/financial_clean_price.h` with PRICE.
//
// The primary verification anchor is round-trip parity with PRICE: for
// every (settlement, maturity, rate, redemption, frequency, basis)
// tuple, `YIELD(..., PRICE(..., y, ...), ...)` must recover `y` to
// ~1e-9. This is the strongest possible test of the closed-form
// inversion (n==1 branch) and the Newton-Raphson iteration (n>1
// branch) -- if either drifts, the round-trip residual blows up.

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
// YIELD -- canonical Microsoft documented case.
// ---------------------------------------------------------------------------

TEST(FinancialYield, MicrosoftDocCanonicalCaseRoundTrip) {
  // Round-trip the Microsoft canonical PRICE example. PRICE returns
  // 94.6343616213 for yld = 0.065; YIELD with that price as input must
  // recover 0.065.
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 94.6343616213, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.065, 1e-9);
}

TEST(FinancialYield, RoundTripPriceCanonical) {
  // PRICE(...) -> 94.6343616213; feeding that price into YIELD must
  // return back 0.065. End-to-end via parser; verifies the wiring at
  // both PRICE and YIELD's registration.
  const Value v = EvalSource(
      "=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, "
      "PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0), "
      "100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.065, 1e-9);
}

TEST(FinancialYield, DefaultBasisZero) {
  // basis omitted defaults to 0; same result as the explicit basis=0
  // case above.
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 94.6343616213, 100, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.065, 1e-9);
}

TEST(FinancialYield, BasisOneActualActualRoundTrip) {
  // Settlement falls inside a coupon period under actual/actual day
  // counting (PRICE's BasisOneActualActual case): yld = 0.06 produces
  // PRICE = 95.4528335058. YIELD must recover yld.
  const Value v = EvalSource(
      "=YIELD(DATE(2010,5,15), DATE(2015,10,1), 0.05, "
      "PRICE(DATE(2010,5,15), DATE(2015,10,1), 0.05, 0.06, 100, 2, 1), "
      "100, 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.06, 1e-9);
}

TEST(FinancialYield, SinglePeriodLinearDiscountClosedForm) {
  // n == 1 branch: settlement falls inside the last coupon period.
  // Same arguments as PRICE's SinglePeriodLinearDiscount test, where
  // PRICE returned 99.7925123001 for yld = 0.065. The closed-form
  // analytic inversion must recover 0.065 to high precision (no Newton
  // iteration involved).
  const Value v = EvalSource("=YIELD(DATE(2017,8,15), DATE(2017,11,15), 0.0575, 99.7925123001, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.065, 1e-9);
}

TEST(FinancialYield, SinglePeriodRoundTripViaPrice) {
  // n == 1 branch round-tripped through PRICE: full closed-form
  // inversion verification end-to-end.
  const Value v = EvalSource(
      "=YIELD(DATE(2017,8,15), DATE(2017,11,15), 0.0575, "
      "PRICE(DATE(2017,8,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0), "
      "100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.065, 1e-9);
}

TEST(FinancialYield, PremiumYieldBelowRate) {
  // yld < rate => price > 100 (premium). PRICE returned 105.7233426194
  // for yld = 0.05 in the canonical shape. YIELD must recover 0.05.
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 105.7233426194, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_LT(v.as_number(), 0.0575);  // yld < rate (sanity)
  EXPECT_NEAR(v.as_number(), 0.05, 1e-9);
}

TEST(FinancialYield, PremiumRoundTripViaPrice) {
  // Premium case round-tripped through PRICE.
  const Value v = EvalSource(
      "=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, "
      "PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.05, 100, 2, 0), "
      "100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.05, 1e-9);
}

TEST(FinancialYield, DiscountRoundTrip) {
  // Discount case (yld > rate => price < 100). Round-trip the
  // canonical case where yld = 0.065 produced price = 94.6343...
  const Value v = EvalSource(
      "=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, "
      "PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.07, 100, 2, 0), "
      "100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0575);  // yld > rate (sanity)
  EXPECT_NEAR(v.as_number(), 0.07, 1e-9);
}

TEST(FinancialYield, ZeroCouponBondRoundTrip) {
  // rate == 0: zero-coupon bond. PRICE collapsed to redemption *
  // v^(t1+n-1) = 53.5974124569 for yld = 0.065. YIELD must invert that.
  const Value v = EvalSource(
      "=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0, "
      "PRICE(DATE(2008,2,15), DATE(2017,11,15), 0, 0.065, 100, 2, 0), "
      "100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.065, 1e-9);
}

// ---------------------------------------------------------------------------
// YIELD -- domain / validation errors.
// ---------------------------------------------------------------------------

TEST(FinancialYield, SettlementEqualMaturityIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2017,11,15), DATE(2017,11,15), 0.0575, 95, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2018,1,1), DATE(2017,11,15), 0.0575, 95, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, FrequencyThreeIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95, 100, 3, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, FrequencyZeroIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95, 100, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, BasisFiveIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95, 100, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, BasisNegativeIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95, 100, 2, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, NegativeRateIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), -0.01, 95, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, ZeroPriceIsNum) {
  // pr <= 0 is YIELD-specific: a zero market price would imply infinite
  // yield, which Excel rejects with #NUM!.
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, NegativePriceIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, -1, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, ZeroRedemptionIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95, 0, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialYield, NegativeRedemptionIsNum) {
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95, -50, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// YIELD -- dispatcher / propagation behaviour.
// ---------------------------------------------------------------------------

TEST(FinancialYield, ErrorInArgPropagates) {
  // `propagate_errors = true` short-circuits before the impl runs;
  // the first `#N/A` argument flows through unchanged.
  const Value v = EvalSource("=YIELD(NA(), DATE(2017,11,15), 0.0575, 95, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(FinancialYield, TooFewArgsIsValue) {
  // arity 5 is below the min_arity of 6.
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialYield, TooManyArgsIsValue) {
  // arity 8 exceeds max_arity of 7.
  const Value v = EvalSource("=YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95, 100, 2, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
