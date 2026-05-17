// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end tests for the Macaulay-duration / modified-duration bond
// built-ins: DURATION and MDURATION. Both live in
// `eval/builtins/financial_duration.cpp` and share the coupon-schedule
// engine in `eval/coupon_schedule.h`.

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
// DURATION
// ---------------------------------------------------------------------------

TEST(FinancialDuration, MicrosoftDocCanonicalCase) {
  // Microsoft's canonical example (Excel docs, "DURATION function"):
  //   settlement = 2008-01-01, maturity = 2016-01-01,
  //   coupon = 8%, yld = 9%, frequency = 2 (semi-annual),
  //   basis = 1 (actual/actual)
  // Expected: 5.9937749555 (matches MS reference & manual closed-form).
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 5.9937749555, 1e-9);
}

TEST(FinancialDuration, DefaultBasisZero) {
  // basis omitted defaults to 0 (US 30/360).
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2)");
  ASSERT_TRUE(v.is_number());
  // Settlement is exactly on a coupon-date boundary, so basis 0 vs 1
  // both produce a full period-day count from settlement to NCD; the
  // result equals the basis-1 case to the same numeric precision.
  EXPECT_NEAR(v.as_number(), 5.9937749555, 1e-9);
}

TEST(FinancialDuration, SettlementEqualMaturityIsNum) {
  const Value v = EvalSource("=DURATION(DATE(2016,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDuration, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=DURATION(DATE(2016,1,1), DATE(2008,1,1), 0.08, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDuration, FrequencyThreeIsNum) {
  // Only frequencies 1, 2, 4 are valid.
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 3, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDuration, FrequencyZeroIsNum) {
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDuration, BasisFiveIsNum) {
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDuration, BasisNegativeIsNum) {
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDuration, ZeroCouponIsValid) {
  // coupon == 0 is a valid zero-coupon bond. With CF_i = 0 for i<n and
  // CF_n = 1, num/den simplifies to time_n = (n - 1) + t1; dividing by
  // frequency yields time-to-maturity in years (~8 for our shape).
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_number());
  // n=16, t1=1.0 -> time_n = 16, duration = 16/2 = 8.0 years.
  EXPECT_NEAR(v.as_number(), 8.0, 1e-9);
}

TEST(FinancialDuration, NegativeCouponIsNum) {
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), -0.01, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDuration, ZeroYieldIsValid) {
  // yld == 0 collapses v to 1; the discounted cashflows are just the
  // raw cashflows, and the result is the cashflow-weighted average
  // time-to-payment. Just assert it's finite and positive.
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0, 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_LT(v.as_number(), 8.0);  // less than time-to-maturity for non-zero coupon
}

TEST(FinancialDuration, NegativeYieldIsNum) {
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, -0.01, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDuration, ErrorInArgPropagates) {
  const Value v = EvalSource("=DURATION(NA(), DATE(2016,1,1), 0.08, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(FinancialDuration, TooFewArgsIsValue) {
  // arity 4 is below the min_arity of 5.
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialDuration, TooManyArgsIsValue) {
  // arity 7 exceeds max_arity of 6.
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialDuration, NonNumericTextCouponIsValue) {
  // A bare unrecognised text argument coerces to #VALUE!.
  const Value v = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), \"abc\", 0.09, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// MDURATION
// ---------------------------------------------------------------------------

TEST(FinancialMDuration, MicrosoftDocCanonicalCase) {
  // Same shape as DURATION's canonical case. Modified duration:
  //   MDURATION = DURATION / (1 + yld/frequency)
  //             = 5.9937749555 / 1.045
  //             = 5.7356698139
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 5.7356698139, 1e-9);
}

TEST(FinancialMDuration, DefaultBasisZero) {
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 5.7356698139, 1e-9);
}

TEST(FinancialMDuration, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=MDURATION(DATE(2016,1,1), DATE(2008,1,1), 0.08, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialMDuration, FrequencyThreeIsNum) {
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 3, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialMDuration, BasisFiveIsNum) {
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialMDuration, NegativeCouponIsNum) {
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), -0.01, 0.09, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialMDuration, NegativeYieldIsNum) {
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, -0.01, 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialMDuration, ZeroYieldEqualsDuration) {
  // With yld == 0, MDURATION = DURATION / 1 = DURATION exactly.
  const Value d = EvalSource("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0, 2, 1)");
  const Value m = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0, 2, 1)");
  ASSERT_TRUE(d.is_number());
  ASSERT_TRUE(m.is_number());
  EXPECT_DOUBLE_EQ(m.as_number(), d.as_number());
}

TEST(FinancialMDuration, ErrorInArgPropagates) {
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, NA(), 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(FinancialMDuration, TooFewArgsIsValue) {
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialMDuration, TooManyArgsIsValue) {
  const Value v = EvalSource("=MDURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
