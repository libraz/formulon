// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end tests for the irregular-last-period bond-pricing
// builtins: ODDLPRICE and ODDLYIELD. The implementations live in
// `eval/builtins/financial_oddlprice.cpp` /
// `eval/builtins/financial_oddlyield.cpp` and share the schedule
// helper in `eval/builtins/financial_oddl_helpers.h`.

#include <cmath>
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
// ODDLPRICE -- canonical Microsoft documented case.
// ---------------------------------------------------------------------------

TEST(FinancialOddl, MicrosoftDocCanonicalOddlPrice) {
  // Microsoft's canonical ODDLPRICE example:
  //   settlement = 2008-02-07, maturity = 2008-06-15,
  //   last_interest = 2007-10-15,
  //   rate = 3.75%, yld = 4.05%, redemption = 100,
  //   frequency = 2 (semi-annual), basis = 0 (US 30/360).
  // Microsoft's docs print the expected as `99.8782860831`. Working the
  // closed form by hand with integer 30/360 day counts (DC=240, A=112,
  // DSC=128, E=180) gives `99.87828601472134` exactly -- the trailing
  // `.31` in the docs is sub-1e-7 noise from a slightly different
  // intermediate rounding (Microsoft's published value cannot be
  // reproduced from the integer day-count formula at full precision).
  // Verify both that we produce the closed-form value and that we land
  // within ~1e-7 of the documented anchor (matches "at 4 decimal places"
  // in task spec; tighter than that against the published docs is not
  // achievable without a Mac oracle measurement).
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 99.87828601472134, 1e-12);
  EXPECT_NEAR(v.as_number(), 99.8782860831, 1e-6);
}

TEST(FinancialOddl, MicrosoftDocCanonicalOddlYield) {
  // ODDLYIELD inverse: feeding our own ODDLPRICE output back through
  // ODDLYIELD must recover the original 4.05% yld exactly (1e-12).
  // Using the documented Microsoft price `99.8782860831` instead lands
  // ~2e-9 off because the documented price differs by ~7e-8 from our
  // closed-form value (see MicrosoftDocCanonicalOddlPrice). Anchor the
  // tight check on our self-consistent inverse and the looser check on
  // the documented value.
  const Value yld_from_doc_pr =
      EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.8782860831, 100, 2, 0)");
  ASSERT_TRUE(yld_from_doc_pr.is_number());
  EXPECT_NEAR(yld_from_doc_pr.as_number(), 0.0405, 1e-7);

  const Value yld_from_self_pr =
      EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.87828601472134, 100, 2, 0)");
  ASSERT_TRUE(yld_from_self_pr.is_number());
  EXPECT_NEAR(yld_from_self_pr.as_number(), 0.0405, 1e-12);
}

TEST(FinancialOddl, OddlPriceDefaultBasisZero) {
  // basis omitted defaults to 0 (US 30/360); same result as the
  // explicit basis=0 case.
  const Value v = EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 99.87828601472134, 1e-12);
}

TEST(FinancialOddl, OddlYieldDefaultBasisZero) {
  // basis omitted defaults to 0 for ODDLYIELD as well; round-trip
  // against our own self-consistent ODDLPRICE output.
  const Value v =
      EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.87828601472134, 100, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0405, 1e-12);
}

// ---------------------------------------------------------------------------
// ODDLPRICE / ODDLYIELD -- round-trip across all five bases.
// ---------------------------------------------------------------------------

// Helper: compute ODDLPRICE for a given (yld, basis), then feed the
// resulting price into ODDLYIELD and check it recovers `yld`.
void RoundTripPriceFromYield(double yld, int basis) {
  const std::string price_src = std::string("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, ") +
                                std::to_string(yld) + ", 100, 2, " + std::to_string(basis) + ")";
  const Value price_v = EvalSource(price_src);
  ASSERT_TRUE(price_v.is_number()) << "basis=" << basis << " yld=" << yld;
  const double price = price_v.as_number();

  // Format `price` with full precision so the parser sees the exact
  // double bits round-tripped through the formula.
  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", price);
  const std::string yield_src = std::string("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, ") +
                                buf + ", 100, 2, " + std::to_string(basis) + ")";
  const Value yield_v = EvalSource(yield_src);
  ASSERT_TRUE(yield_v.is_number()) << "basis=" << basis << " yld=" << yld;
  EXPECT_NEAR(yield_v.as_number(), yld, 1e-12) << "basis=" << basis << " yld=" << yld;
}

// Helper: compute ODDLYIELD from a given (pr, basis), then feed the
// resulting yield back into ODDLPRICE and check it recovers `pr`.
void RoundTripYieldFromPrice(double pr, int basis) {
  char buf_pr[64];
  std::snprintf(buf_pr, sizeof(buf_pr), "%.17g", pr);
  const std::string yield_src = std::string("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, ") +
                                buf_pr + ", 100, 2, " + std::to_string(basis) + ")";
  const Value yield_v = EvalSource(yield_src);
  ASSERT_TRUE(yield_v.is_number()) << "basis=" << basis << " pr=" << pr;
  const double yld = yield_v.as_number();

  char buf_y[64];
  std::snprintf(buf_y, sizeof(buf_y), "%.17g", yld);
  const std::string price_src = std::string("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, ") +
                                buf_y + ", 100, 2, " + std::to_string(basis) + ")";
  const Value price_v = EvalSource(price_src);
  ASSERT_TRUE(price_v.is_number()) << "basis=" << basis << " pr=" << pr;
  EXPECT_NEAR(price_v.as_number(), pr, 1e-12) << "basis=" << basis << " pr=" << pr;
}

TEST(FinancialOddl, RoundTripPriceFromYieldBasis0) {
  RoundTripPriceFromYield(0.0405, 0);
}
TEST(FinancialOddl, RoundTripPriceFromYieldBasis1) {
  RoundTripPriceFromYield(0.0405, 1);
}
TEST(FinancialOddl, RoundTripPriceFromYieldBasis2) {
  RoundTripPriceFromYield(0.0405, 2);
}
TEST(FinancialOddl, RoundTripPriceFromYieldBasis3) {
  RoundTripPriceFromYield(0.0405, 3);
}
TEST(FinancialOddl, RoundTripPriceFromYieldBasis4) {
  RoundTripPriceFromYield(0.0405, 4);
}

TEST(FinancialOddl, RoundTripYieldFromPriceBasis0) {
  RoundTripYieldFromPrice(99.5, 0);
}
TEST(FinancialOddl, RoundTripYieldFromPriceBasis1) {
  RoundTripYieldFromPrice(99.5, 1);
}
TEST(FinancialOddl, RoundTripYieldFromPriceBasis2) {
  RoundTripYieldFromPrice(99.5, 2);
}
TEST(FinancialOddl, RoundTripYieldFromPriceBasis3) {
  RoundTripYieldFromPrice(99.5, 3);
}
TEST(FinancialOddl, RoundTripYieldFromPriceBasis4) {
  RoundTripYieldFromPrice(99.5, 4);
}

// ---------------------------------------------------------------------------
// ODDLPRICE -- coverage for non-default frequencies.
// ---------------------------------------------------------------------------

TEST(FinancialOddl, OddlPriceQuarterlyFrequency) {
  // frequency = 4 (quarterly): exercises the period_days = 360/4 = 90
  // (basis 0) path. We don't have a Mac value to anchor against, but
  // confirm the result is a finite number in the sensible range and
  // that ODDLYIELD recovers the input yld exactly.
  const Value price_v =
      EvalSource("=ODDLPRICE(DATE(2024,5,15), DATE(2024,11,15), DATE(2024,2,15), 0.04, 0.045, 100, 4, 0)");
  ASSERT_TRUE(price_v.is_number());
  // Sanity: discount + small accrual swing => price near par for a
  // small (rate - yld) gap over a sub-year horizon.
  EXPECT_GT(price_v.as_number(), 50.0);
  EXPECT_LT(price_v.as_number(), 150.0);

  char buf[64];
  std::snprintf(buf, sizeof(buf), "%.17g", price_v.as_number());
  const std::string yield_src =
      std::string("=ODDLYIELD(DATE(2024,5,15), DATE(2024,11,15), DATE(2024,2,15), 0.04, ") + buf + ", 100, 4, 0)";
  const Value yield_v = EvalSource(yield_src);
  ASSERT_TRUE(yield_v.is_number());
  EXPECT_NEAR(yield_v.as_number(), 0.045, 1e-12);
}

TEST(FinancialOddl, OddlPriceAnnualFrequency) {
  // frequency = 1 (annual): period_days = 360 (basis 0).
  const Value price_v =
      EvalSource("=ODDLPRICE(DATE(2010,3,15), DATE(2010,12,31), DATE(2009,6,30), 0.06, 0.05, 100, 1, 0)");
  ASSERT_TRUE(price_v.is_number());
  EXPECT_GT(price_v.as_number(), 50.0);
  EXPECT_LT(price_v.as_number(), 150.0);
}

// ---------------------------------------------------------------------------
// ODDLPRICE -- domain / validation errors.
// ---------------------------------------------------------------------------

TEST(FinancialOddl, OddlPriceLastInterestEqualSettlementIsNum) {
  // last_interest must be strictly less than settlement.
  const Value v = EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2008,2,7), 0.0375, 0.0405, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceLastInterestAfterSettlementIsNum) {
  // last_interest > settlement is rejected.
  const Value v = EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2008,2,8), 0.0375, 0.0405, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceSettlementEqualMaturityIsNum) {
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,6,15), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceSettlementAfterMaturityIsNum) {
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,7,1), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceFrequencyThreeIsNum) {
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 3, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceFrequencyZeroIsNum) {
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceBasisFiveIsNum) {
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceBasisNegativeIsNum) {
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 2, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceNegativeRateIsNum) {
  const Value v = EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), -0.01, 0.0405, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceNegativeYieldIsNum) {
  const Value v = EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, -0.01, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceZeroRedemptionIsNum) {
  const Value v = EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 0, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlPriceNegativeRedemptionIsNum) {
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, -50, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// ODDLYIELD -- domain / validation errors.
// ---------------------------------------------------------------------------

TEST(FinancialOddl, OddlYieldLastInterestEqualSettlementIsNum) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2008,2,7), 0.0375, 99.5, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlYieldSettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,7,1), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.5, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlYieldFrequencyThreeIsNum) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.5, 100, 3, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlYieldBasisFiveIsNum) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.5, 100, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlYieldNegativeRateIsNum) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), -0.01, 99.5, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlYieldZeroPriceIsNum) {
  // pr <= 0 is rejected (zero price would imply infinite yield).
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlYieldNegativePriceIsNum) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, -50, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialOddl, OddlYieldZeroRedemptionIsNum) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.5, 0, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// ODDLPRICE / ODDLYIELD -- dispatcher / propagation behaviour.
// ---------------------------------------------------------------------------

TEST(FinancialOddl, OddlPriceErrorInArgPropagates) {
  // `propagate_errors = true` short-circuits before the impl runs;
  // the first `#N/A` argument flows through unchanged.
  const Value v = EvalSource("=ODDLPRICE(NA(), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(FinancialOddl, OddlYieldErrorInArgPropagates) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), NA(), 99.5, 100, 2, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(FinancialOddl, OddlPriceTooFewArgsIsValue) {
  // arity 6 is below the min_arity of 7.
  const Value v = EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialOddl, OddlPriceTooManyArgsIsValue) {
  // arity 9 exceeds max_arity of 8.
  const Value v =
      EvalSource("=ODDLPRICE(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 0.0405, 100, 2, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialOddl, OddlYieldTooFewArgsIsValue) {
  const Value v = EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.5, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialOddl, OddlYieldTooManyArgsIsValue) {
  const Value v =
      EvalSource("=ODDLYIELD(DATE(2008,2,7), DATE(2008,6,15), DATE(2007,10,15), 0.0375, 99.5, 100, 2, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
