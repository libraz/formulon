//
// End-to-end tests for the irregular-first-period bond-pricing
// builtins: ODDFPRICE and ODDFYIELD. The implementations live in
// `eval/builtins/financial_oddfprice.cpp` /
// `eval/builtins/financial_oddfyield.cpp` and share the schedule
// helper in `eval/builtins/financial_oddf_helpers.h`.

#include <cmath>
#include <cstdio>
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

// ---------------------------------------------------------------------------
// ODDFPRICE -- canonical Microsoft documented case.
// ---------------------------------------------------------------------------

TEST(FinancialOddf, MicrosoftDocCanonicalOddfPrice) {
  // Microsoft's canonical ODDFPRICE example:
  //   settlement = 2008-11-11, maturity = 2021-03-01,
  //   issue = 2008-10-15, first_coupon = 2009-03-01,
  //   rate = 7.85%, yld = 6.25%, redemption = 100,
  //   frequency = 2 (semi-annual), basis = 1 (Actual/Actual).
  // Microsoft's docs publish ~113.5977. The brief tightens "match
  // within 1e-3 of the published value"; exact-bit parity will be
  // measured against Mac later.
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_number()) << "got error: " << static_cast<int>(v.is_error() ? v.as_error() : ErrorCode::NA);
  EXPECT_NEAR(v.as_number(), 113.598, 1e-3);
}

TEST(FinancialOddf, OddfPriceDefaultBasisZero) {
  // basis omitted defaults to 0 (US 30/360); same result as the
  // explicit basis=0 case. Confirm both are finite + equal.
  const Value v_default = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2)");
  const Value v_explicit = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 0)");
  ASSERT_TRUE(v_default.is_number());
  ASSERT_TRUE(v_explicit.is_number());
  EXPECT_DOUBLE_EQ(v_default.as_number(), v_explicit.as_number());
}

TEST(FinancialOddf, OddfYieldDefaultBasisZero) {
  // basis omitted defaults to 0 for ODDFYIELD as well; round-trip
  // against our own self-consistent ODDFPRICE output.
  const Value price_v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2)");
  ASSERT_TRUE(price_v.is_number());
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", price_v.as_number());
  const std::string yield_src =
      std::string("=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, ") + buf +
      ", 100, 2)";
  const Value yield_v = EvalSource(yield_src);
  ASSERT_TRUE(yield_v.is_number());
  EXPECT_NEAR(yield_v.as_number(), 0.0625, 1e-9);
}

// ---------------------------------------------------------------------------
// ODDFPRICE / ODDFYIELD -- round-trip across all five bases.
// ---------------------------------------------------------------------------

// Helper: compute ODDFPRICE for a given (yld, basis), then feed the
// resulting price into ODDFYIELD and check it recovers `yld`. Anchors
// on Microsoft's canonical case so the irregular first span is
// non-trivial across bases (issue 2008-10-15 -> first_coupon
// 2009-03-01 fits within one normal semi-annual period under most
// bases, but exercise NC=1 with a meaningful test span).
void RoundTripPriceFromYield(double yld, int basis) {
  const std::string price_src =
      std::string("=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, ") +
      std::to_string(yld) + ", 100, 2, " + std::to_string(basis) + ")";
  const Value price_v = EvalSource(price_src);
  ASSERT_TRUE(price_v.is_number()) << "basis=" << basis << " yld=" << yld;
  const double price = price_v.as_number();

  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", price);
  const std::string yield_src =
      std::string("=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, ") + buf +
      ", 100, 2, " + std::to_string(basis) + ")";
  const Value yield_v = EvalSource(yield_src);
  ASSERT_TRUE(yield_v.is_number()) << "basis=" << basis << " yld=" << yld;
  EXPECT_NEAR(yield_v.as_number(), yld, 1e-9) << "basis=" << basis << " yld=" << yld;
}

// Helper: compute ODDFYIELD from a given (pr, basis), then feed the
// resulting yield back into ODDFPRICE and check it recovers `pr`.
void RoundTripYieldFromPrice(double pr, int basis) {
  char buf_pr[64];
  std::snprintf(buf_pr, sizeof(buf_pr), "%.17g", pr);
  const std::string yield_src =
      std::string("=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, ") + buf_pr +
      ", 100, 2, " + std::to_string(basis) + ")";
  const Value yield_v = EvalSource(yield_src);
  ASSERT_TRUE(yield_v.is_number()) << "basis=" << basis << " pr=" << pr;
  const double yld = yield_v.as_number();

  char buf_y[64];
  std::snprintf(buf_y, sizeof(buf_y), "%.17g", yld);
  const std::string price_src =
      std::string("=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, ") + buf_y +
      ", 100, 2, " + std::to_string(basis) + ")";
  const Value price_v = EvalSource(price_src);
  ASSERT_TRUE(price_v.is_number()) << "basis=" << basis << " pr=" << pr;
  EXPECT_NEAR(price_v.as_number(), pr, 1e-9) << "basis=" << basis << " pr=" << pr;
}

TEST(FinancialOddf, RoundTripPriceFromYieldBasis0) {
  RoundTripPriceFromYield(0.0625, 0);
}
TEST(FinancialOddf, RoundTripPriceFromYieldBasis1) {
  RoundTripPriceFromYield(0.0625, 1);
}
TEST(FinancialOddf, RoundTripPriceFromYieldBasis2) {
  RoundTripPriceFromYield(0.0625, 2);
}
TEST(FinancialOddf, RoundTripPriceFromYieldBasis3) {
  RoundTripPriceFromYield(0.0625, 3);
}
TEST(FinancialOddf, RoundTripPriceFromYieldBasis4) {
  RoundTripPriceFromYield(0.0625, 4);
}

TEST(FinancialOddf, RoundTripYieldFromPriceBasis0) {
  RoundTripYieldFromPrice(113.5, 0);
}
TEST(FinancialOddf, RoundTripYieldFromPriceBasis1) {
  RoundTripYieldFromPrice(113.5, 1);
}
TEST(FinancialOddf, RoundTripYieldFromPriceBasis2) {
  RoundTripYieldFromPrice(113.5, 2);
}
TEST(FinancialOddf, RoundTripYieldFromPriceBasis3) {
  RoundTripYieldFromPrice(113.5, 3);
}
TEST(FinancialOddf, RoundTripYieldFromPriceBasis4) {
  RoundTripYieldFromPrice(113.5, 4);
}

// ---------------------------------------------------------------------------
// Short first period (NC == 1): issue and first_coupon fit within one
// normal coupon period.
// ---------------------------------------------------------------------------

TEST(FinancialOddf, ShortFirstPeriodSensiblePrice) {
  // NC == 1 case: issue 2024-03-15, first_coupon 2024-09-15
  // (semi-annual = 6 months apart -> exactly one normal period under
  // 30/360, so NC=1 by definition). Settlement is in the middle.
  // Coupon rate = yld -> price should be near par minus a small amount
  // for the partial first-period accrual reduction (vs. a regular
  // bond, ODDFPRICE accounts for the partial first coupon).
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2024,6,1), DATE(2027,9,15), DATE(2024,3,15), DATE(2024,9,15), 0.05, 0.05, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  // For rate == yld and a non-trivial irregular first span, the price
  // is close to par but not exactly par; sanity-check the range.
  EXPECT_GT(v.as_number(), 80.0);
  EXPECT_LT(v.as_number(), 110.0);

  // Confirm round-trip.
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", v.as_number());
  const std::string yield_src =
      std::string("=ODDFYIELD(DATE(2024,6,1), DATE(2027,9,15), DATE(2024,3,15), DATE(2024,9,15), 0.05, ") + buf +
      ", 100, 2, 0)";
  const Value yld_v = EvalSource(yield_src);
  ASSERT_TRUE(yld_v.is_number());
  EXPECT_NEAR(yld_v.as_number(), 0.05, 1e-9);
}

// ---------------------------------------------------------------------------
// Long first period (NC > 1): issue spans more than one normal period
// before first_coupon.
// ---------------------------------------------------------------------------

TEST(FinancialOddf, LongFirstPeriodSensiblePriceNc3) {
  // Long-first case: issue 2008-01-01, first_coupon 2009-07-01,
  // semi-annual frequency. Walking back from 2009-07-01 by 6 months:
  //   2009-01-01, 2008-07-01, 2008-01-01 = issue
  // Three quasi-periods: NC = 3. Settlement somewhere inside.
  const Value v =
      EvalSource("=ODDFPRICE(DATE(2008,4,1), DATE(2015,1,1), DATE(2008,1,1), DATE(2009,7,1), 0.06, 0.05, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 90.0);
  EXPECT_LT(v.as_number(), 130.0);

  // Round-trip.
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", v.as_number());
  const std::string yield_src =
      std::string("=ODDFYIELD(DATE(2008,4,1), DATE(2015,1,1), DATE(2008,1,1), DATE(2009,7,1), 0.06, ") + buf +
      ", 100, 2, 0)";
  const Value yld_v = EvalSource(yield_src);
  ASSERT_TRUE(yld_v.is_number());
  EXPECT_NEAR(yld_v.as_number(), 0.05, 1e-9);
}

TEST(FinancialOddf, LongFirstPeriodAcrossAllBases) {
  // Same NC>1 setup, exercise each basis. Round-trip must still close.
  for (int basis = 0; basis <= 4; ++basis) {
    const std::string price_src =
        std::string("=ODDFPRICE(DATE(2008,4,1), DATE(2015,1,1), DATE(2008,1,1), DATE(2009,7,1), 0.06, 0.05, 100, 2, ") +
        std::to_string(basis) + ")";
    const Value price_v = EvalSource(price_src);
    ASSERT_TRUE(price_v.is_number()) << "basis=" << basis;

    char buf[64];
    std::snprintf(buf, sizeof(buf), "%.17g", price_v.as_number());
    const std::string yield_src =
        std::string("=ODDFYIELD(DATE(2008,4,1), DATE(2015,1,1), DATE(2008,1,1), DATE(2009,7,1), 0.06, ") + buf +
        ", 100, 2, " + std::to_string(basis) + ")";
    const Value yld_v = EvalSource(yield_src);
    ASSERT_TRUE(yld_v.is_number()) << "basis=" << basis;
    EXPECT_NEAR(yld_v.as_number(), 0.05, 1e-9) << "basis=" << basis;
  }
}

// ---------------------------------------------------------------------------
// Frequency = 4 (quarterly).
// ---------------------------------------------------------------------------

TEST(FinancialOddf, OddfPriceQuarterlyFrequency) {
  // Quarterly frequency exercises period_days = 360/4 = 90 (basis 0).
  // Issue 2024-01-15, first_coupon 2024-07-15 -> 6 months = 2
  // quarterly periods, so NC = 2.
  const Value price_v = EvalSource(
      "=ODDFPRICE(DATE(2024,4,1), DATE(2026,7,15), DATE(2024,1,15), DATE(2024,7,15), 0.04, 0.045, 100, 4, 0)");
  ASSERT_TRUE(price_v.is_number());
  EXPECT_GT(price_v.as_number(), 80.0);
  EXPECT_LT(price_v.as_number(), 120.0);

  // Round-trip.
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", price_v.as_number());
  const std::string yield_src =
      std::string("=ODDFYIELD(DATE(2024,4,1), DATE(2026,7,15), DATE(2024,1,15), DATE(2024,7,15), 0.04, ") + buf +
      ", 100, 4, 0)";
  const Value yield_v = EvalSource(yield_src);
  ASSERT_TRUE(yield_v.is_number());
  EXPECT_NEAR(yield_v.as_number(), 0.045, 1e-9);
}

TEST(FinancialOddf, OddfPriceAnnualFrequency) {
  // Annual frequency: period_days = 360 (basis 0). Issue 2018-01-01,
  // first_coupon 2020-01-01 -> NC = 2 (long first).
  const Value price_v =
      EvalSource("=ODDFPRICE(DATE(2018,7,1), DATE(2025,1,1), DATE(2018,1,1), DATE(2020,1,1), 0.06, 0.05, 100, 1, 0)");
  ASSERT_TRUE(price_v.is_number());
  EXPECT_GT(price_v.as_number(), 80.0);
  EXPECT_LT(price_v.as_number(), 130.0);
}

// ---------------------------------------------------------------------------
// ODDFPRICE -- domain / validation errors.
// ---------------------------------------------------------------------------

TEST(FinancialOddf, OddfPriceIssueEqualSettlementIsNum) {
  // issue must be strictly less than settlement.
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,10,15), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceIssueAfterSettlementIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,10,1), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceSettlementEqualFirstCouponIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2009,3,1), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceSettlementAfterFirstCouponIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2009,4,1), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceFirstCouponEqualMaturityIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2009,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceFirstCouponAfterMaturityIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2009,1,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceFrequencyThreeIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 3, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceFrequencyZeroIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceBasisFiveIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceBasisNegativeIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceNegativeRateIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), -0.01, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceNegativeYieldIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, -0.01, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceZeroRedemptionIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 0, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfPriceNegativeRedemptionIsNum) {
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, -50, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// ODDFYIELD -- domain / validation errors.
// ---------------------------------------------------------------------------

TEST(FinancialOddf, OddfYieldIssueEqualSettlementIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,10,15), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 113.5, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfYieldSettlementAfterFirstCouponIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2009,4,1), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 113.5, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfYieldFirstCouponAfterMaturityIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2009,1,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 113.5, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfYieldFrequencyThreeIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 113.5, 100, 3, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfYieldBasisFiveIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 113.5, 100, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfYieldNegativeRateIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), -0.01, 113.5, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfYieldZeroPriceIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfYieldNegativePriceIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, -50, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddf, OddfYieldZeroRedemptionIsNum) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 113.5, 0, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// ODDFPRICE / ODDFYIELD -- dispatcher / propagation behaviour.
// ---------------------------------------------------------------------------

TEST(FinancialOddf, OddfPriceErrorInArgPropagates) {
  // `propagate_errors = true` short-circuits before the impl runs;
  // the first `#N/A` argument flows through unchanged.
  const Value v =
      EvalSource("=ODDFPRICE(NA(), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(FinancialOddf, OddfYieldErrorInArgPropagates) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), NA(), 113.5, 100, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(FinancialOddf, OddfPriceTooFewArgsIsValue) {
  // arity 7 is below the min_arity of 8.
  const Value v =
      EvalSource("=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialOddf, OddfPriceTooManyArgsIsValue) {
  // arity 10 exceeds max_arity of 9.
  const Value v = EvalSource(
      "=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 0.0625, 100, 2, 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialOddf, OddfYieldTooFewArgsIsValue) {
  const Value v =
      EvalSource("=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 113.5, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialOddf, OddfYieldTooManyArgsIsValue) {
  const Value v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 113.5, 100, 2, 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// ODDFYIELD -- convergence robustness.
// ---------------------------------------------------------------------------

TEST(FinancialOddf, OddfYieldConvergesFromMicrosoftHeuristic) {
  // The Microsoft "approximate yield" initial guess is the only
  // public formula and is reliably within a few percent of the truth.
  // Verify it converges from a far-off price (deep discount). Bonds
  // trading well below par (yld >> rate) are the most adversarial
  // case for the heuristic.
  const Value yld_v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 60, 100, 2, 1)");
  ASSERT_TRUE(yld_v.is_number());
  // Round-trip: pricing the recovered yield must reproduce the input.
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", yld_v.as_number());
  const std::string price_src =
      std::string("=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, ") + buf +
      ", 100, 2, 1)";
  const Value price_v = EvalSource(price_src);
  ASSERT_TRUE(price_v.is_number());
  EXPECT_NEAR(price_v.as_number(), 60.0, 1e-9);
}

TEST(FinancialOddf, OddfYieldConvergesFromHighPremium) {
  // Deep premium (yld << rate). Same robustness check.
  const Value yld_v = EvalSource(
      "=ODDFYIELD(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, 140, 100, 2, 1)");
  ASSERT_TRUE(yld_v.is_number());
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", yld_v.as_number());
  const std::string price_src =
      std::string("=ODDFPRICE(DATE(2008,11,11), DATE(2021,3,1), DATE(2008,10,15), DATE(2009,3,1), 0.0785, ") + buf +
      ", 100, 2, 1)";
  const Value price_v = EvalSource(price_src);
  ASSERT_TRUE(price_v.is_number());
  EXPECT_NEAR(price_v.as_number(), 140.0, 1e-9);
}

TEST(FinancialOddf, OddfYieldLowYieldRoundTrip) {
  const Value yld_v = EvalSource(
      "=ODDFYIELD(DATE(2024,1,1), DATE(2054,1,1), DATE(2023,1,1), DATE(2024,7,1), 0.005, "
      "ODDFPRICE(DATE(2024,1,1), DATE(2054,1,1), DATE(2023,1,1), DATE(2024,7,1), 0.005, 0.0125, 100, 2, 0), "
      "100, 2, 0)");
  ASSERT_TRUE(yld_v.is_number());
  EXPECT_NEAR(yld_v.as_number(), 0.0125, 1e-9);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
