// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the accrual and specialized-depreciation
// financial built-ins: ACCRINT, ACCRINTM, VDB, AMORDEGRC, AMORLINC.
// These live in `eval/builtins/financial_accrual.cpp` (ACCRINT family)
// and `eval/builtins/financial_depreciation.cpp` (VDB, AMORDEGRC,
// AMORLINC) and share the YEARFRAC helpers in `eval/date_time.h`.

#include <cmath>
#include <string>
#include <string_view>

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
// with no bound workbook. Shared across scalar-only accrual / special
// depreciation tests.
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
// ACCRINT
// ---------------------------------------------------------------------------

TEST(FinancialAccrint, MicrosoftDocExample) {
  // ACCRINT docs: issue 2008-03-01, first_interest 2008-08-31,
  // settlement 2008-05-01, rate 0.1, par 1000, freq 2, basis 0,
  // calc_method TRUE. YEARFRAC(2008-03-01, 2008-05-01, 0) = 60/360 =
  // 0.166666..., so ACCRINT = 1000 * 0.1 * 0.166666... = 16.6667.
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2, 0, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.6667, 1e-3);
}

TEST(FinancialAccrint, CalcMethodTrueDefault) {
  // Omitting calc_method (defaults to TRUE): same start date (issue),
  // so the result matches the 8-arg case above.
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.6667, 1e-3);
}

TEST(FinancialAccrint, CalcMethodFalseAccruesFromIssue) {
  // Mac Excel 365 (16.108.1, ja-JP) always accrues from issue to
  // settlement, ignoring both `first_interest` and `calc_method`. The MS
  // docs claim calc_method=FALSE switches the start to first_interest
  // when settlement > first_interest, but the actual product does not.
  // issue 2008-03-01, settlement 2008-08-31. YEARFRAC(0) = 180/360 = 0.5,
  // so result = 1000 * 0.1 * 0.5 = 50.0 regardless of first_interest.
  const Value v_far = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,5,1), DATE(2008,8,31), 0.1, 1000, 2, 0, FALSE)");
  ASSERT_TRUE(v_far.is_number());
  EXPECT_NEAR(v_far.as_number(), 50.0, 1e-3);

  // calc_method=TRUE on the same (issue, settlement) must produce the
  // same value: the calc_method argument is ignored, only validated for
  // type correctness.
  const Value v_true = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,5,1), DATE(2008,8,31), 0.1, 1000, 2, 0, TRUE)");
  ASSERT_TRUE(v_true.is_number());
  EXPECT_NEAR(v_true.as_number(), 50.0, 1e-3);
}

TEST(FinancialAccrint, CalcMethodFalseEarlySettlementMatchesIssueAccrual) {
  // Even when settlement < first_interest, Mac Excel still accrues from
  // issue (this case happens to match the MS doc result coincidentally,
  // because the doc-style calc_method=FALSE branch would also pick
  // issue here). issue 2008-03-01, settlement 2008-05-01.
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2, 0, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.6667, 1e-3);
}

TEST(FinancialAccrint, BasisThreeActual365) {
  // YEARFRAC(2008-03-01, 2008-05-01, 3) = actual days / 365.
  // 2008 is a leap year; Mar 1 -> May 1 = 61 days (Mar 31 + Apr 30).
  // result = 1000 * 0.1 * 61/365 = 16.7123...
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.7123, 1e-3);
}

TEST(FinancialAccrint, IssueGeSettlementIsNum) {
  const Value v = EvalSource("=ACCRINT(DATE(2008,5,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAccrint, NonPositiveRateIsNum) {
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0, 1000, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAccrint, NonPositiveParIsNum) {
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 0, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAccrint, InvalidFrequencyIsNum) {
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAccrint, InvalidBasisIsNum) {
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// ACCRINTM
// ---------------------------------------------------------------------------

TEST(FinancialAccrintm, MicrosoftDocExample) {
  // ACCRINTM docs: issue 2008-04-01, settlement 2008-06-15, rate 0.1,
  // par 1000, basis 3 (actual/365). Days = 75, result = 1000 * 0.1 *
  // 75 / 365 = 20.5479...
  const Value v = EvalSource("=ACCRINTM(DATE(2008,4,1), DATE(2008,6,15), 0.1, 1000, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 20.5479, 1e-3);
}

TEST(FinancialAccrintm, DefaultBasisZero) {
  // Basis 0 (US 30/360): days = 74 (30*2 + 14), result = 1000 * 0.1 *
  // 74 / 360 = 20.5555...
  const Value v = EvalSource("=ACCRINTM(DATE(2008,4,1), DATE(2008,6,15), 0.1, 1000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 20.5556, 1e-3);
}

TEST(FinancialAccrintm, IssueGeSettlementIsNum) {
  const Value v = EvalSource("=ACCRINTM(DATE(2008,6,15), DATE(2008,4,1), 0.1, 1000, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAccrintm, NonPositiveRateIsNum) {
  const Value v = EvalSource("=ACCRINTM(DATE(2008,4,1), DATE(2008,6,15), -0.1, 1000, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAccrintm, NonPositiveParIsNum) {
  const Value v = EvalSource("=ACCRINTM(DATE(2008,4,1), DATE(2008,6,15), 0.1, 0, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAccrintm, InvalidBasisIsNum) {
  const Value v = EvalSource("=ACCRINTM(DATE(2008,4,1), DATE(2008,6,15), 0.1, 1000, 7)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// VDB
// ---------------------------------------------------------------------------

TEST(FinancialVdb, SingleDayNoSwitchFalse) {
  // MS doc: VDB(2400, 300, 10*365, 0, 1, 2, FALSE) = 1.315 per day.
  // DDB charge = 2400 * 2 / 3650 = 1.31507. Straight-line candidate
  // (2400 - 300) / 3650 = 0.5753 < DDB, so no switch.
  const Value v = EvalSource("=VDB(2400, 300, 10*365, 0, 1, 2, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.315, 1e-2);
}

TEST(FinancialVdb, FractionalEndPeriodFactor15) {
  // VDB(2400, 300, 10, 0, 0.875, 1.5, FALSE). factor/life = 0.15,
  // DDB for period 1 = 2400 * 0.15 = 360, clipped to [0, 0.875] =
  // 0.875 * 360 = 315.
  const Value v = EvalSource("=VDB(2400, 300, 10, 0, 0.875, 1.5, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 315.0, 1e-2);
}

TEST(FinancialVdb, LaterPeriodsSumToEndOfLife) {
  // VDB(2400, 300, 10, 6, 10, 2, FALSE) sums the last four periods.
  // By period 6 the book value is ~629; DDB still beats straight-line
  // until the salvage floor caps the last period. Periods 7..10 sum
  // to cost - salvage - (dep over 0..6), giving ~329 — i.e. the
  // entire remaining depreciable amount ~cost-salvage-(depreciation
  // through period 6). Anchored to the computed sum 329.14.
  const Value v = EvalSource("=VDB(2400, 300, 10, 6, 10, 2, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 329.14, 1.0);
}

TEST(FinancialVdb, FullLifeSumsToCostMinusSalvage) {
  // Integrating over the whole [0, life] range must equal the full
  // depreciable amount: cost - salvage = 2100.
  const Value v = EvalSource("=VDB(2400, 300, 10, 0, 10, 2, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2100.0, 1e-6);
}

TEST(FinancialVdb, NoSwitchTrueFullDdb) {
  // With no_switch=TRUE, declining-balance never switches to
  // straight-line; charges decrease geometrically. Should be strictly
  // less than the no_switch=FALSE value at the tail (where SL wins).
  const Value v = EvalSource("=VDB(2400, 300, 10, 6, 10, 2, TRUE)");
  ASSERT_TRUE(v.is_number());
  // Just verify positive, finite, and less than full asset value.
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_LT(v.as_number(), 2100.0);
}

TEST(FinancialVdb, DefaultFactorIsTwo) {
  // VDB default factor = 2. VDB(1000, 100, 5, 0, 1) should match
  // VDB(1000, 100, 5, 0, 1, 2).
  const Value a = EvalSource("=VDB(1000, 100, 5, 0, 1)");
  const Value b = EvalSource("=VDB(1000, 100, 5, 0, 1, 2)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), b.as_number(), 1e-10);
}

TEST(FinancialVdb, StartPeriodEqEndPeriodIsZero) {
  // start_period == end_period is a zero-length interval: every clipped
  // segment collapses to width 0 and the loop accumulates nothing. Mac
  // Excel 365 returns 0.0 (no period of depreciation), not #NUM!.
  const Value v = EvalSource("=VDB(1000, 100, 5, 3, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(FinancialVdb, StartPeriodGtEndPeriodIsNum) {
  // start_period > end_period is still rejected.
  const Value v = EvalSource("=VDB(1000, 100, 5, 4, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialVdb, EndPeriodExceedsLifeIsNum) {
  const Value v = EvalSource("=VDB(1000, 100, 5, 0, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialVdb, NegativeStartPeriodIsNum) {
  const Value v = EvalSource("=VDB(1000, 100, 5, -1, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialVdb, NegativeCostIsNum) {
  const Value v = EvalSource("=VDB(-1000, 100, 5, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialVdb, NegativeFactorIsNum) {
  // factor < 0 is rejected with #NUM!.
  const Value v = EvalSource("=VDB(1000, 100, 5, 0, 1, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialVdb, ZeroFactorFallsThroughToStraightLine) {
  // factor == 0 makes the DDB rate 0, so straight-line wins on every
  // period. For (cost=1000, salvage=100, life=5) the SL charge per full
  // period is (1000 - 100) / 5 = 180. Mac Excel 365 returns this finite
  // value rather than #NUM!.
  const Value v = EvalSource("=VDB(1000, 100, 5, 0, 1, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 180.0, 1e-9);
}

TEST(FinancialVdb, NonPositiveLifeIsNum) {
  const Value v = EvalSource("=VDB(1000, 100, 0, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// AMORDEGRC
// ---------------------------------------------------------------------------

TEST(FinancialAmordegrc, MicrosoftDocExample) {
  // MS doc: AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1,
  // 0.15, 1) = 776. life = 1/0.15 = 6.67, coefficient = 2.5, applied
  // rate = 0.375. Period 0 (partial): 2400 * 0.375 * YEARFRAC(2008-08-19,
  // 2008-12-31, actual/actual) = 2400 * 0.375 * 134/366 ~ 330.
  // Period 1 (full): round((2400 - 330) * 0.375) = 776.
  const Value v = EvalSource("=AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 776.0, 1.5);
}

TEST(FinancialAmordegrc, PeriodZeroPartial) {
  // Period 0 is the partial first period. With applied_rate = 0.375
  // and YEARFRAC ~ 0.366, result ~ 329 (rounded).
  const Value v = EvalSource("=AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 0, 0.15, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 330.0, 1.5);
}

TEST(FinancialAmordegrc, CostLeZeroIsNum) {
  const Value v = EvalSource("=AMORDEGRC(0, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmordegrc, SalvageGeCostIsNum) {
  const Value v = EvalSource("=AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 2400, 1, 0.15, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmordegrc, NegativePeriodIsNum) {
  const Value v = EvalSource("=AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, -1, 0.15, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmordegrc, NonPositiveRateIsNum) {
  const Value v = EvalSource("=AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmordegrc, InvalidBasisIsNum) {
  const Value v = EvalSource("=AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15, 7)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmordegrc, LifeBucketRejection) {
  // life = 1/rate = 1/0.4 = 2.5, which is below the minimum bucket
  // (life >= 3). Excel returns #NUM!.
  const Value v = EvalSource("=AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.4, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// AMORLINC
// ---------------------------------------------------------------------------

TEST(FinancialAmorlinc, MicrosoftDocExample) {
  // MS doc: AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1,
  // 0.15, 1) = 360. Period 1 (second period, flat) = 2400 * 0.15 = 360.
  const Value v = EvalSource("=AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 360.0, 1e-2);
}

TEST(FinancialAmorlinc, PeriodZeroPartial) {
  // Period 0 is the prorated first period: 2400 * 0.15 * YEARFRAC ~
  // 2400 * 0.15 * 134/366 = 131.80.
  const Value v = EvalSource("=AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 0, 0.15, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 131.80, 0.5);
}

TEST(FinancialAmorlinc, CostLeZeroIsNum) {
  const Value v = EvalSource("=AMORLINC(-10, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmorlinc, SalvageGeCostIsNum) {
  const Value v = EvalSource("=AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 2400, 1, 0.15, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmorlinc, NegativePeriodIsNum) {
  const Value v = EvalSource("=AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, -1, 0.15, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmorlinc, NonPositiveRateIsNum) {
  const Value v = EvalSource("=AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, -0.1, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmorlinc, InvalidBasisIsNum) {
  const Value v = EvalSource("=AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15, 9)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialAmorlinc, DefaultBasisZero) {
  // Basis 0 (US 30/360). Same period=1, flat full-period charge.
  const Value v = EvalSource("=AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1, 0.15)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 360.0, 1e-2);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
