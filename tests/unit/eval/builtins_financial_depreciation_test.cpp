// Financial builtin tests grouped by schedule family.

#include <cmath>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {
using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

TEST(FinancialSLN, Basic) {
  const Value v = EvalSource("=SLN(10000, 1000, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1800.0);
}

TEST(FinancialSLN, ZeroSalvage) {
  const Value v = EvalSource("=SLN(10000, 0, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1000.0);
}

TEST(FinancialSLN, IdentitySlnTimesLifeEqualsDepreciableBase) {
  // SLN * life == cost - salvage for any legal inputs.
  const Value v = EvalSource("=SLN(10000, 1000, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number() * 5.0, 10000.0 - 1000.0);
}

TEST(FinancialSLN, LifeZeroIsDiv0) {
  const Value v = EvalSource("=SLN(10000, 1000, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialSLN, NegativeLifeProducesNegativeDepreciation) {
  // Oracle-verified: Excel 365 Mac returns the raw algebraic answer
  // (-1800 for these inputs) rather than #NUM!.
  const Value v = EvalSource("=SLN(10000, 1000, -5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -1800.0);
}

TEST(FinancialSLN, ArityMismatchIsValue) {
  const Value v = EvalSource("=SLN(10000, 1000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialSYD, FirstPeriod) {
  // (10000-1000)*5*2/(5*6) = 3000
  const Value v = EvalSource("=SYD(10000, 1000, 5, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3000.0);
}

TEST(FinancialSYD, LastPeriod) {
  // (10000-1000)*1*2/(5*6) = 600
  const Value v = EvalSource("=SYD(10000, 1000, 5, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 600.0);
}

TEST(FinancialSYD, SumAcrossAllPeriodsEqualsDepreciableBase) {
  // sum_{p=1..life} SYD(cost, salvage, life, p) == cost - salvage.
  double total = 0.0;
  for (int p = 1; p <= 5; ++p) {
    const std::string src = "=SYD(10000, 1000, 5, " + std::to_string(p) + ")";
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_number()) << "p=" << p;
    total += v.as_number();
  }
  EXPECT_NEAR(total, 9000.0, 1e-9);
}

TEST(FinancialSYD, PeriodOutOfRangeIsNum) {
  const Value v = EvalSource("=SYD(10000, 1000, 5, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialSYD, PeriodZeroIsNum) {
  const Value v = EvalSource("=SYD(10000, 1000, 5, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialSYD, FractionalPeriodAccepted) {
  // Excel 365 accepts fractional periods: the SYD formula is a linear
  // schedule so period=0.1 evaluates naturally. Regression guard for the
  // relaxation of the `period < 1` rejection.
  //   (290-2)*(5-0.1+1)*2 / (5*6) = 288 * 5.9 * 2 / 30 = 113.28
  const Value v = EvalSource("=SYD(290, 2, 5, 0.1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 113.28);
}

TEST(FinancialSYD, LifeZeroIsNum) {
  const Value v = EvalSource("=SYD(10000, 1000, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDDB, FirstPeriod) {
  // rate = 2/5 = 0.4; dep_1 = 10000 * 0.4 = 4000
  const Value v = EvalSource("=DDB(10000, 1000, 5, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4000.0);
}

TEST(FinancialDDB, SecondPeriod) {
  // book = 6000 after period 1; dep_2 = 6000 * 0.4 = 2400
  const Value v = EvalSource("=DDB(10000, 1000, 5, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2400.0);
}

TEST(FinancialDDB, CustomFactorTriples) {
  // factor=3 -> rate = 3/5 = 0.6; dep_1 = 10000 * 0.6 = 6000
  const Value v = EvalSource("=DDB(10000, 1000, 5, 1, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6000.0);
}

TEST(FinancialDDB, SalvageFloorCaps) {
  // Running the full schedule: dep must never take book value below
  // salvage. Verify total == cost - salvage (9000) exactly — DDB
  // depreciates the full base over the asset's life, with the salvage
  // floor absorbing whatever remains in the last period. For this
  // schedule the final period is the capped one: 1296 * 0.4 = 518.4
  // would overshoot, so dep_5 = book - salvage = 296.
  double total = 0.0;
  for (int p = 1; p <= 5; ++p) {
    const std::string src = "=DDB(10000, 1000, 5, " + std::to_string(p) + ")";
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_number()) << "p=" << p;
    total += v.as_number();
  }
  EXPECT_NEAR(total, 9000.0, 1e-9);
  const Value last = EvalSource("=DDB(10000, 1000, 5, 5)");
  ASSERT_TRUE(last.is_number());
  EXPECT_NEAR(last.as_number(), 296.0, 1e-9);
}

TEST(FinancialDDB, PeriodOutOfRangeIsNum) {
  const Value v = EvalSource("=DDB(10000, 1000, 5, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDDB, NegativeCostIsNum) {
  const Value v = EvalSource("=DDB(-100, 0, 5, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDDB, FactorZeroIsNum) {
  const Value v = EvalSource("=DDB(10000, 1000, 5, 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDB, FullYearFirstPeriod) {
  // rate = round(1 - (0.1)^0.2, 3) = 0.369; dep_1 = 10000 * 0.369 = 3690.
  const Value v = EvalSource("=DB(10000, 1000, 5, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3690.0, 1e-6);
}

TEST(FinancialDB, PartialFirstYearProrated) {
  // month=6: dep_1 = 10000 * 0.369 * 6/12 = 1845.
  const Value v = EvalSource("=DB(10000, 1000, 5, 1, 6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1845.0, 1e-6);
}

TEST(FinancialDB, PartialYearSumApproxEqualsDepreciableBase) {
  // With month=6 the depreciation runs life+1 = 6 periods; the sum
  // should approximately equal cost - salvage within the 3-decimal
  // rate-rounding error.
  double total = 0.0;
  for (int p = 1; p <= 6; ++p) {
    const std::string src = "=DB(10000, 1000, 5, " + std::to_string(p) + ", 6)";
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_number()) << "p=" << p;
    total += v.as_number();
  }
  // Excel's 3-decimal rate rounding produces a small residual between
  // the sum of depreciations and the true depreciable base. Observed
  // residual for this fixture (cost=10000, salvage=1000, life=5,
  // month=6) is ~54 — tolerance of 100 leaves margin for rounding
  // accumulation without masking a real regression.
  EXPECT_NEAR(total, 9000.0, 100.0);
}

TEST(FinancialDB, PeriodLifePlusOneNeedsPartialFirstYear) {
  // month=12 (full year) with period == life+1 is invalid.
  const Value v = EvalSource("=DB(10000, 1000, 5, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDB, PeriodLifePlusOneWithPartialYearIsValid) {
  // month=6, period=6 is the partial last year; must produce a
  // positive, finite charge.
  const Value v = EvalSource("=DB(10000, 1000, 5, 6, 6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
}

TEST(FinancialDB, InvalidMonthIsNum) {
  const Value v = EvalSource("=DB(10000, 1000, 5, 1, 13)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDB, ZeroMonthIsNum) {
  const Value v = EvalSource("=DB(10000, 1000, 5, 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDb, MonthIsFlooredToInteger) {
  // Mac Excel 365 evaluates DB using INT(month), not the raw fractional
  // month argument. With month=9.2222 the formula reduces to rate=0.438,
  // dep_1=32.85, dep_2=29.4117, dep_3=(100-32.85-29.4117)*0.438 =
  // 16.5293754 — the same value Mac Excel and the IronCalc G10 oracle
  // produce. Prior to the INT(month) fix this evaluated to ~16.329735
  // because the fractional month fed directly into the rate-prorated
  // first period.
  const Value v = EvalSource("=DB(100, 10, 4, 3, 9.2222)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 16.5293754, 1e-6);
}

TEST(FinancialDb, ZeroCostReturnsZero) {
  // Excel short-circuits zero-cost assets to zero depreciation to avoid
  // the `(salvage/0)^(1/life)` blow-up. Verified against IronCalc G30.
  const Value v = EvalSource("=DB(0, 10, 4, 1, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(FinancialDDB, HugeScheduleFallsBackToClosedForm) {
  // Without the cap this would step 1e18 iterations. The closed form
  // `cost * ((1-rate)^(p-1) - (1-rate)^p)` is exact for integer periods,
  // so the answer is still well-defined.
  const Value v = EvalSource("=DDB(1, 0, 2000000, 2000000)");
  ASSERT_TRUE(v.is_number());
  const double rate = 2.0 / 2000000.0;
  const double expected = std::pow(1.0 - rate, 1999999.0) - std::pow(1.0 - rate, 2000000.0);
  EXPECT_NEAR(v.as_number(), expected, expected * 1e-9);
}

TEST(FinancialDDB, ClosedFormAgreesWithTheIterativeSchedule) {
  // The fallback is only sound because both branches compute the same
  // number. Pin that at the largest period the iterative branch still
  // takes, so a change to either branch is caught here.
  const Value v = EvalSource("=DDB(1, 0, 2000000, 1048576)");
  ASSERT_TRUE(v.is_number());
  const double rate = 2.0 / 2000000.0;
  const double expected = std::pow(1.0 - rate, 1048575.0) - std::pow(1.0 - rate, 1048576.0);
  EXPECT_NEAR(v.as_number(), expected, expected * 1e-8);
}

TEST(FinancialDb, HugeScheduleIsNum) {
  // DB's schedule carries a rounded rate and a partial-first-year term
  // with no closed form to fall back on, so an over-long schedule is
  // refused rather than approximated.
  const Value v = EvalSource("=DB(10000, 1000, 1E18, 1E18, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDb, ScheduleWithinTheCapStillComputes) {
  const Value v = EvalSource("=DB(10000, 1000, 5, 1, 6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1845.0, 1e-9);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
