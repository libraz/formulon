//
// End-to-end tests for the accrual and specialized-depreciation
// financial built-ins: ACCRINT, ACCRINTM, VDB, AMORDEGRC, AMORLINC.
// These live in `eval/builtins/financial_accrual.cpp` (ACCRINT family)
// and `eval/builtins/financial_depreciation.cpp` (VDB, AMORDEGRC,
// AMORLINC) and share the YEARFRAC helpers in `utils/date_time.h`.

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
// ACCRINT
// ---------------------------------------------------------------------------

TEST(FinancialAccrint, MicrosoftDocExample) {
  // ACCRINT docs: issue 2008-03-01 lies in the quasi-coupon period that
  // ends on first_interest 2008-08-31, so the accrual is the direct
  // 30/360 span 60/180 of a 50.0 coupon = 16.6667.
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2, 0, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.6667, 1e-3);
}

TEST(FinancialAccrint, CalcMethodTrueDefault) {
  // Omitting calc_method (defaults to TRUE) matches the 8-arg case above.
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.6667, 1e-3);
}

TEST(FinancialAccrint, SettlementAfterFirstInterestWithIssueInLastPeriod) {
  // Issue in the period ending at first_interest: both calc_method values
  // accrue the full 180/360 span from issue (Mac Excel: 50.0).
  const Value v_false = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,5,1), DATE(2008,8,31), 0.1, 1000, 2, 0, FALSE)");
  ASSERT_TRUE(v_false.is_number());
  EXPECT_NEAR(v_false.as_number(), 50.0, 1e-9);

  const Value v_true = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,5,1), DATE(2008,8,31), 0.1, 1000, 2, 0, TRUE)");
  ASSERT_TRUE(v_true.is_number());
  EXPECT_NEAR(v_true.as_number(), 50.0, 1e-9);
}

TEST(FinancialAccrint, CalcMethodFalseEarlySettlementWithIssueInLastPeriod) {
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2, 0, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.6667, 1e-3);
}

TEST(FinancialAccrint, IssueBeforeLastPeriodMeasuresFromPenultimateQuasiDate) {
  // Quarterly grid from 2008-08-31: 05-31, 02-29. Issue 03-01 to 05-31 is
  // 90 days, then 05-31 -> settlement 05-01 is -29 (30/360): 61/90 of 25.
  const Value v = EvalSource("=ACCRINT(DATE(2008,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.944444444444446, 1e-9);
}

TEST(FinancialAccrint, CalcMethodFalseDropsWholePeriodsAndCanGoNegative) {
  // Basis 1: 59/90 issue part, -153/91 from 2017-09-01 back to settlement,
  // plus the 2 whole periods only when calc_method is TRUE (Mac Excel).
  const Value v_true = EvalSource("=ACCRINT(DATE(2017,1,1), DATE(2017,12,1), DATE(2017,4,1), 0.33, 3000, 4, 1)");
  ASSERT_TRUE(v_true.is_number());
  EXPECT_NEAR(v_true.as_number(), 241.12362637362637, 1e-9);

  const Value v_false =
      EvalSource("=ACCRINT(DATE(2017,1,1), DATE(2017,12,1), DATE(2017,4,1), 0.33, 3000, 4, 1, FALSE)");
  ASSERT_TRUE(v_false.is_number());
  EXPECT_NEAR(v_false.as_number(), -253.8763736263736, 1e-9);
}

TEST(FinancialAccrint, SettlementAfterFirstInterestCountsWholePeriodsForTrue) {
  // TRUE counts each whole period after first_interest as 1 and the
  // settlement period as 90/182; FALSE keeps the raw 456/182 day span.
  const Value v_true = EvalSource("=ACCRINT(DATE(2020,1,1), DATE(2020,7,1), DATE(2021,4,1), 0.075, 100, 2, 1, TRUE)");
  ASSERT_TRUE(v_true.is_number());
  EXPECT_NEAR(v_true.as_number(), 9.354395604395604, 1e-9);

  const Value v_false = EvalSource("=ACCRINT(DATE(2020,1,1), DATE(2020,7,1), DATE(2021,4,1), 0.075, 100, 2, 1, FALSE)");
  ASSERT_TRUE(v_false.is_number());
  EXPECT_NEAR(v_false.as_number(), 9.395604395604396, 1e-9);
}

TEST(FinancialAccrint, MonthEndFirstInterestKeepsQuasiDatesOnMonthEnd) {
  // first_interest 2024-02-29 puts the previous quasi date on 2023-08-31,
  // not 08-29: 30/360 from 08-31 back to 2023-05-01 is -119 days.
  const Value v = EvalSource("=ACCRINT(DATE(2022,8,31), DATE(2024,2,29), DATE(2023,5,1), 0.075, 100, 2, 0, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -2.4791666666666665, 1e-9);
}

TEST(FinancialAccrint, IssueAfterFirstInterestTrueWalksForwardFalseUsesRawSpan) {
  // Quarterly grid from 2024-03-31; E = 91 days (2023-12-31 -> 2024-03-31).
  // TRUE: 29/91 issue part + 2 whole periods + 41/91 (Mac Excel 34.615385);
  // FALSE: the raw 254-day span over E (Mac Excel 34.890110).
  const Value v_true = EvalSource("=ACCRINT(DATE(2024,6,1),DATE(2024,3,31),DATE(2025,2,10),0.05,1000,4,1,TRUE)");
  ASSERT_TRUE(v_true.is_number());
  EXPECT_NEAR(v_true.as_number(), 34.61538461538461, 1e-9);

  const Value v_false = EvalSource("=ACCRINT(DATE(2024,6,1),DATE(2024,3,31),DATE(2025,2,10),0.05,1000,4,1,FALSE)");
  ASSERT_TRUE(v_false.is_number());
  EXPECT_NEAR(v_false.as_number(), 34.89010989010989, 1e-9);
}

TEST(FinancialAccrint, IssueAfterFirstInterestBasisZeroMultiPeriod) {
  // Annual 30/360: TRUE counts 351/360 + 1 whole period + 184/360, while
  // FALSE is DAYS360(issue, settlement) = 894 over 360.
  const Value v_true = EvalSource("=ACCRINT(DATE(2024,1,10),DATE(2023,12,31),DATE(2026,7,4),0.05,1000,1,0,TRUE)");
  ASSERT_TRUE(v_true.is_number());
  EXPECT_NEAR(v_true.as_number(), 124.30555555555556, 1e-9);

  const Value v_false = EvalSource("=ACCRINT(DATE(2024,1,10),DATE(2023,12,31),DATE(2026,7,4),0.05,1000,1,0,FALSE)");
  ASSERT_TRUE(v_false.is_number());
  EXPECT_NEAR(v_false.as_number(), 124.16666666666667, 1e-9);
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

TEST(FinancialAmordegrc, ExcessivePeriodIsNumNotAHang) {
  // A period far beyond kMaxDepreciationPeriods must reject before ever
  // entering the per-period loop, matching the cap DB/DDB/VDB already
  // enforce in the same file -- this used to loop once per period (or hit
  // UB on the int64 cast for period > 9.2e18).
  const Value v = EvalSource("=AMORDEGRC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1E18, 0.15, 1)");
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

TEST(FinancialAmorlinc, ExcessivePeriodIsNumNotAHang) {
  const Value v = EvalSource("=AMORLINC(2400, DATE(2008,8,19), DATE(2008,12,31), 300, 1E18, 0.15, 1)");
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

// VDB walks one iteration per integer period up to `end_period`, and its
// switch-to-straight-line state machine has no closed form to fall back on,
// so a schedule longer than the Excel grid's row count is refused outright
// rather than stepping for longer than the process will live.
TEST(FinancialVdb, HugeScheduleIsNum) {
  const Value v = EvalSource("=VDB(2400, 300, 1E18, 0, 1E18)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialVdb, ScheduleWithinTheCapStillComputes) {
  const Value v = EvalSource("=VDB(2400, 300, 10, 0, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2100.0, 1e-9);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
