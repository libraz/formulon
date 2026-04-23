// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the coupon-date financial built-ins:
// COUPPCD, COUPNCD, COUPNUM, COUPDAYBS, COUPDAYSNC, COUPDAYS. All six
// share the shared coupon-schedule engine in `eval/coupon_schedule.h`
// and implementations in `eval/builtins/financial_coupon.cpp`.

#include <cmath>
#include <string>
#include <string_view>

#include "eval/coupon_schedule.h"
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

// Parses `src` and evaluates it through the default function registry
// with no bound workbook.
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
// Canonical bond example — 2011-01-25 settle, 2011-11-15 mature, semi-annual
// ---------------------------------------------------------------------------
// Expected schedule: PCD = 2010-11-15 (serial 40497), NCD = 2011-05-15
// (serial 40678), coupons remaining = 2. Verified by hand + Python.

TEST(FinancialCoupPcd, CanonicalBondSemiAnnualBasis1) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 40497.0);
}

TEST(FinancialCoupNcd, CanonicalBondSemiAnnualBasis1) {
  const Value v = EvalSource("=COUPNCD(DATE(2011,1,25), DATE(2011,11,15), 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 40678.0);
}

TEST(FinancialCoupNum, CanonicalBondSemiAnnualBasis1) {
  // 2011-05-15 and 2011-11-15 are strictly after settlement 2011-01-25
  // and <= maturity: 2 coupons remaining.
  const Value v = EvalSource("=COUPNUM(DATE(2011,1,25), DATE(2011,11,15), 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(FinancialCoupDayBs, CanonicalBondSemiAnnualBasis1) {
  // Actual days from 2010-11-15 to 2011-01-25 = 71.
  const Value v = EvalSource("=COUPDAYBS(DATE(2011,1,25), DATE(2011,11,15), 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 71.0);
}

TEST(FinancialCoupDaysNc, CanonicalBondSemiAnnualBasis1) {
  // Actual days from 2011-01-25 to 2011-05-15 = 110.
  const Value v = EvalSource("=COUPDAYSNC(DATE(2011,1,25), DATE(2011,11,15), 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 110.0);
}

TEST(FinancialCoupDays, CanonicalBondSemiAnnualBasis1) {
  // Basis 1 (actual/actual): NCD - PCD = 40678 - 40497 = 181.
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 181.0);
}

// ---------------------------------------------------------------------------
// Basis variations for COUPDAYS — period length by frequency / basis
// ---------------------------------------------------------------------------

TEST(FinancialCoupDays, SemiAnnualBasis0Is180) {
  // Basis 0 (US 30/360) semi-annual: 360 / 2 = 180.
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 180.0);
}

TEST(FinancialCoupDays, QuarterlyBasis0Is90) {
  // Basis 0 quarterly: 360 / 4 = 90.
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 4, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 90.0);
}

TEST(FinancialCoupDays, AnnualBasis0Is360) {
  // Basis 0 annual: 360 / 1 = 360.
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 1, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 360.0);
}

TEST(FinancialCoupDays, SemiAnnualBasis2Is180) {
  // Basis 2 (actual/360) semi-annual: 360 / 2 = 180.
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 180.0);
}

TEST(FinancialCoupDays, SemiAnnualBasis3Is182Point5) {
  // Basis 3 (actual/365) semi-annual: 365 / 2 = 182.5.
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 182.5, 1e-9);
}

TEST(FinancialCoupDays, SemiAnnualBasis4Is180) {
  // Basis 4 (EU 30/360) semi-annual: 360 / 2 = 180.
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 180.0);
}

TEST(FinancialCoupDays, QuarterlyBasis3Is91Point25) {
  // Basis 3 quarterly: 365 / 4 = 91.25.
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 4, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 91.25, 1e-9);
}

// ---------------------------------------------------------------------------
// COUPDAYBS / COUPDAYSNC with 30/360 bases (0 and 4) on the canonical example
// ---------------------------------------------------------------------------
// US 30/360 from 2010-11-15 to 2011-01-25: d1=15, d2=25, no adjustment;
// days = 360*1 + 30*(1-11) + (25-15) = 70.

TEST(FinancialCoupDayBs, Basis0Us30_360Is70) {
  const Value v = EvalSource("=COUPDAYBS(DATE(2011,1,25), DATE(2011,11,15), 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 70.0);
}

TEST(FinancialCoupDayBs, Basis4Eu30_360Is70) {
  const Value v = EvalSource("=COUPDAYBS(DATE(2011,1,25), DATE(2011,11,15), 2, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 70.0);
}

// 30/360 from 2011-01-25 to 2011-05-15: days = 30*(5-1) + (15-25) = 120 - 10 = 110.
TEST(FinancialCoupDaysNc, Basis0Us30_360Is110) {
  const Value v = EvalSource("=COUPDAYSNC(DATE(2011,1,25), DATE(2011,11,15), 2, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 110.0);
}

TEST(FinancialCoupDaysNc, Basis4Eu30_360Is110) {
  const Value v = EvalSource("=COUPDAYSNC(DATE(2011,1,25), DATE(2011,11,15), 2, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 110.0);
}

// ---------------------------------------------------------------------------
// Default basis (omitted) == 0
// ---------------------------------------------------------------------------

TEST(FinancialCoupPcd, DefaultBasisZero) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 40497.0);
}

TEST(FinancialCoupDays, DefaultBasisIs180) {
  const Value v = EvalSource("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 180.0);
}

// ---------------------------------------------------------------------------
// End-of-month clamp — maturity 2020-08-31, quarterly, settlement 2020-06-01
// ---------------------------------------------------------------------------
// Stepping back 3 months from 2020-08-31 lands on 2020-05-31 (May has 31
// days, no clamp); settlement 2020-06-01 > 2020-05-31 so PCD = 2020-05-31
// and NCD = 2020-08-31.

TEST(FinancialCoupPcd, EndOfMonthMay31NotMay30) {
  const Value v = EvalSource("=COUPPCD(DATE(2020,6,1), DATE(2020,8,31), 4, 1)");
  ASSERT_TRUE(v.is_number());
  // Serial for 2020-05-31: days since 1899-12-30 = 43982.
  EXPECT_EQ(v.as_number(), 43982.0);
}

TEST(FinancialCoupNcd, EndOfMonthAug31) {
  const Value v = EvalSource("=COUPNCD(DATE(2020,6,1), DATE(2020,8,31), 4, 1)");
  ASSERT_TRUE(v.is_number());
  // Serial for 2020-08-31: 44074.
  EXPECT_EQ(v.as_number(), 44074.0);
}

// Maturity 2020-08-31, settlement 2020-03-01, quarterly. Stepping back
// from Aug-31: (2020, 5, 31) -> 2020-05-31; still > 2020-03-01, step back
// again: (2020, 2, 31) -> clamped to 2020-02-29 (leap year). 2020-02-29
// <= 2020-03-01, so PCD = 2020-02-29.
TEST(FinancialCoupPcd, ClampFeb29OnLeapYear) {
  const Value v = EvalSource("=COUPPCD(DATE(2020,3,1), DATE(2020,8,31), 4, 1)");
  ASSERT_TRUE(v.is_number());
  // Serial for 2020-02-29: 43890.
  EXPECT_EQ(v.as_number(), 43890.0);
}

TEST(FinancialCoupNum, ClampFeb29Quarterly) {
  // From 2020-03-01 to 2020-08-31: coupons at 2020-05-31 and 2020-08-31
  // are strictly after settlement and <= maturity: 2 remaining.
  const Value v = EvalSource("=COUPNUM(DATE(2020,3,1), DATE(2020,8,31), 4, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// Error domain
// ---------------------------------------------------------------------------

TEST(FinancialCoupPcd, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,11,15), DATE(2011,1,25), 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupPcd, SettlementEqualsMaturityIsNum) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,1,25), DATE(2011,1,25), 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupPcd, BadFrequencyThreeIsNum) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 3, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupPcd, BadFrequencyZeroIsNum) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupPcd, BadBasisFiveIsNum) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupPcd, BadBasisNegativeIsNum) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 2, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupNcd, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=COUPNCD(DATE(2011,11,15), DATE(2011,1,25), 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupNum, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=COUPNUM(DATE(2011,11,15), DATE(2011,1,25), 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupDayBs, BadFrequencyIsNum) {
  const Value v = EvalSource("=COUPDAYBS(DATE(2011,1,25), DATE(2011,11,15), 7, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupDaysNc, BadBasisIsNum) {
  const Value v = EvalSource("=COUPDAYSNC(DATE(2011,1,25), DATE(2011,11,15), 2, 99)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCoupDays, NonNumericSettlementIsValue) {
  const Value v = EvalSource("=COUPDAYS(\"oops\", DATE(2011,11,15), 2, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Annual and quarterly frequency sanity checks on the canonical example
// ---------------------------------------------------------------------------

TEST(FinancialCoupNum, AnnualFrequencyOneCoupon) {
  // 2011-01-25 settle, 2011-11-15 mature, annual: stepping back from
  // 2011-11-15 gives 2010-11-15 <= settlement; PCD = 2010-11-15, one
  // coupon left (2011-11-15 itself).
  const Value v = EvalSource("=COUPNUM(DATE(2011,1,25), DATE(2011,11,15), 1, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(FinancialCoupPcd, AnnualFrequency) {
  const Value v = EvalSource("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 1, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 40497.0);  // 2010-11-15
}

TEST(FinancialCoupNum, QuarterlyFrequency) {
  // Quarterly stepping back from 2011-11-15: 2011-08-15, 2011-05-15,
  // 2011-02-15, 2010-11-15. Of these, 2011-02-15 > 2011-01-25 (strictly
  // after settlement), 2011-05-15, 2011-08-15, 2011-11-15 likewise. So
  // 4 coupons remain.
  const Value v = EvalSource("=COUPNUM(DATE(2011,1,25), DATE(2011,11,15), 4, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

// ---------------------------------------------------------------------------
// Direct engine-internals tests — validate the shared `compute_coupon_dates`
// produces the expected PCD / NCD / period_days without going through the
// parser.
// ---------------------------------------------------------------------------

TEST(CouponScheduleEngine, CanonicalExample) {
  CouponDates ctx{};
  // DATE(2011, 1, 25) = 40568; DATE(2011, 11, 15) = 40862.
  ASSERT_TRUE(compute_coupon_dates(40568.0, 40862.0, 2, 1, &ctx));
  EXPECT_EQ(ctx.pcd, 40497.0);
  EXPECT_EQ(ctx.ncd, 40678.0);
  EXPECT_EQ(ctx.coupons_remaining, 2);
  EXPECT_EQ(ctx.days_bs, 71.0);
  EXPECT_EQ(ctx.days_nc, 110.0);
  EXPECT_EQ(ctx.period_days, 181.0);
}

TEST(CouponScheduleEngine, EndOfMonthClampLeap) {
  CouponDates ctx{};
  // DATE(2020, 3, 1) = 43891; DATE(2020, 8, 31) = 44074.
  ASSERT_TRUE(compute_coupon_dates(43891.0, 44074.0, 4, 1, &ctx));
  // PCD = 2020-02-29 = 43890.
  EXPECT_EQ(ctx.pcd, 43890.0);
  // NCD = 2020-05-31 = 43982.
  EXPECT_EQ(ctx.ncd, 43982.0);
  EXPECT_EQ(ctx.coupons_remaining, 2);
}

TEST(CouponScheduleEngine, AnnualFrequency) {
  CouponDates ctx{};
  ASSERT_TRUE(compute_coupon_dates(40568.0, 40862.0, 1, 0, &ctx));
  EXPECT_EQ(ctx.pcd, 40497.0);  // 2010-11-15
  EXPECT_EQ(ctx.ncd, 40862.0);  // 2011-11-15 (maturity itself)
  EXPECT_EQ(ctx.coupons_remaining, 1);
  EXPECT_EQ(ctx.period_days, 360.0);
}

TEST(CouponScheduleEngine, PeriodDaysByBasis) {
  CouponDates ctx{};
  // Semi-annual.
  ASSERT_TRUE(compute_coupon_dates(40568.0, 40862.0, 2, 0, &ctx));
  EXPECT_EQ(ctx.period_days, 180.0);
  ASSERT_TRUE(compute_coupon_dates(40568.0, 40862.0, 2, 2, &ctx));
  EXPECT_EQ(ctx.period_days, 180.0);
  ASSERT_TRUE(compute_coupon_dates(40568.0, 40862.0, 2, 3, &ctx));
  EXPECT_NEAR(ctx.period_days, 182.5, 1e-9);
  ASSERT_TRUE(compute_coupon_dates(40568.0, 40862.0, 2, 4, &ctx));
  EXPECT_EQ(ctx.period_days, 180.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
