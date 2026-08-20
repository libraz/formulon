//
// Shape-acceptance tests for the lazy families whose argument slots are
// documented as taking a range: SERIESSUM, the IRR / MIRR / XIRR / XNPV
// financial group, RANK / PERCENTRANK, the regression group
// (CORREL / SLOPE / FORECAST), and the holidays slot of the workday
// group.
//
// Every one of those slots resolves through `resolve_range_arg`, so the
// same rectangle must be produced whether it is written as an inline
// array literal, produced by a dynamic-array function such as
// `SEQUENCE`, or reached through a spilled-range reference (`A1#`). A
// slot that cannot resolve a shape returns an error; it never computes a
// number from the empty set.

#include <string>
#include <string_view>

#include "cell.h"
#include "eval/cell_evaluator.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
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

// Commits `formula` as a spilling formula at A1 so `A1#` resolves to its
// region. Mirrors the spill fixture used by the COUNT tests.
void SpillAtA1(Workbook& wb, std::string_view formula) {
  Sheet& sheet = wb.sheet(0);
  Cell cell;
  cell.formula_text = std::string(formula);
  Arena arena;
  const Value anchor = evaluate_cell_for_recalc(wb, sheet, cell, 0U, 0U, default_registry(), arena);
  ASSERT_TRUE(anchor.is_number());
  ASSERT_NE(sheet.spill_region_at_anchor(0U, 0U), nullptr);
}

// `A1#` == {1;2;3}.
void SpillSequenceAtA1(Workbook& wb) {
  SpillAtA1(wb, "=SEQUENCE(3)");
}

// ---------------------------------------------------------------------------
// SERIESSUM — the slot that used to answer 0 instead of an error
// ---------------------------------------------------------------------------

TEST(LazyRangeArgShapes, SeriesSumMatchesAcrossCoefficientShapes) {
  // 1·2^0 + 2·2^1 + 3·2^2 = 1 + 4 + 12 = 17.
  const Value literal = EvalSource("=SERIESSUM(2, 0, 1, {1;2;3})");
  ASSERT_TRUE(literal.is_number());
  EXPECT_DOUBLE_EQ(literal.as_number(), 17.0);

  const Value dynamic = EvalSource("=SERIESSUM(2, 0, 1, SEQUENCE(3))");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, SeriesSumAcceptsSpilledRange) {
  Workbook wb = Workbook::create();
  SpillSequenceAtA1(wb);
  const Value v = EvalSourceIn("=SERIESSUM(2, 0, 1, A1#)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), 17.0);
}

TEST(LazyRangeArgShapes, SeriesSumPropagatesCoefficientError) {
  // An unresolvable coefficients slot must surface the error rather than
  // fall back to a sum over no terms.
  const Value v = EvalSource("=SERIESSUM(2, 0, 1, SEQUENCE(0))");
  ASSERT_TRUE(v.is_error());
}

// ---------------------------------------------------------------------------
// IRR / MIRR / XIRR / XNPV
// ---------------------------------------------------------------------------

TEST(LazyRangeArgShapes, IrrMatchesAcrossValueShapes) {
  const Value literal = EvalSource("=IRR({-100;-50;0;50})");
  ASSERT_TRUE(literal.is_number());

  const Value dynamic = EvalSource("=IRR(SEQUENCE(4,1,-100,50))");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, IrrAcceptsSpilledRange) {
  Workbook wb = Workbook::create();
  SpillAtA1(wb, "=SEQUENCE(4,1,-100,50)");
  const Value spilled = EvalSourceIn("=IRR(A1#)", wb, wb.sheet(0));
  ASSERT_TRUE(spilled.is_number()) << "kind=" << static_cast<int>(spilled.kind());
  const Value literal = EvalSource("=IRR({-100;-50;0;50})");
  ASSERT_TRUE(literal.is_number());
  EXPECT_DOUBLE_EQ(spilled.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, MirrAcceptsSpilledRange) {
  Workbook wb = Workbook::create();
  SpillAtA1(wb, "=SEQUENCE(4,1,-100,50)");
  const Value spilled = EvalSourceIn("=MIRR(A1#, 0.1, 0.1)", wb, wb.sheet(0));
  ASSERT_TRUE(spilled.is_number()) << "kind=" << static_cast<int>(spilled.kind());
  const Value literal = EvalSource("=MIRR({-100;-50;0;50}, 0.1, 0.1)");
  ASSERT_TRUE(literal.is_number());
  EXPECT_DOUBLE_EQ(spilled.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, IrrStillRejectsBareScalar) {
  // A lone value is not a cash-flow sequence; the scalar collapse must
  // not turn it into a 1-element series.
  const Value v = EvalSource("=IRR(5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(LazyRangeArgShapes, MirrMatchesAcrossValueShapes) {
  const Value literal = EvalSource("=MIRR({-100;-50;0;50}, 0.1, 0.1)");
  ASSERT_TRUE(literal.is_number());

  const Value dynamic = EvalSource("=MIRR(SEQUENCE(4,1,-100,50), 0.1, 0.1)");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

// SEQUENCE(3,1,-100,80) == {-100;-20;60}: both XIRR / XNPV argument slots
// are exercised, the values slot as well as the parallel dates slot.
TEST(LazyRangeArgShapes, XnpvMatchesAcrossValueAndDateShapes) {
  const Value literal = EvalSource("=XNPV(0.09, {-100;-20;60}, {40000;40100;40200})");
  ASSERT_TRUE(literal.is_number());

  const Value dynamic_dates = EvalSource("=XNPV(0.09, {-100;-20;60}, SEQUENCE(3,1,40000,100))");
  ASSERT_TRUE(dynamic_dates.is_number()) << "kind=" << static_cast<int>(dynamic_dates.kind());
  EXPECT_DOUBLE_EQ(dynamic_dates.as_number(), literal.as_number());

  const Value dynamic_values = EvalSource("=XNPV(0.09, SEQUENCE(3,1,-100,80), {40000;40100;40200})");
  ASSERT_TRUE(dynamic_values.is_number()) << "kind=" << static_cast<int>(dynamic_values.kind());
  EXPECT_DOUBLE_EQ(dynamic_values.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, XnpvAcceptsSpilledRange) {
  Workbook wb = Workbook::create();
  SpillAtA1(wb, "=SEQUENCE(3,1,-100,80)");
  const Value spilled = EvalSourceIn("=XNPV(0.09, A1#, {40000;40100;40200})", wb, wb.sheet(0));
  ASSERT_TRUE(spilled.is_number()) << "kind=" << static_cast<int>(spilled.kind());
  const Value literal = EvalSource("=XNPV(0.09, {-100;-20;60}, {40000;40100;40200})");
  ASSERT_TRUE(literal.is_number());
  EXPECT_DOUBLE_EQ(spilled.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, XirrMatchesAcrossValueAndDateShapes) {
  const Value literal = EvalSource("=XIRR({-100;-20;60}, {40000;40100;40200}, 0.1)");
  ASSERT_TRUE(literal.is_number());

  const Value dynamic_dates = EvalSource("=XIRR({-100;-20;60}, SEQUENCE(3,1,40000,100), 0.1)");
  ASSERT_TRUE(dynamic_dates.is_number()) << "kind=" << static_cast<int>(dynamic_dates.kind());
  EXPECT_DOUBLE_EQ(dynamic_dates.as_number(), literal.as_number());

  const Value dynamic_values = EvalSource("=XIRR(SEQUENCE(3,1,-100,80), {40000;40100;40200}, 0.1)");
  ASSERT_TRUE(dynamic_values.is_number()) << "kind=" << static_cast<int>(dynamic_values.kind());
  EXPECT_DOUBLE_EQ(dynamic_values.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, XirrAcceptsSpilledRange) {
  Workbook wb = Workbook::create();
  SpillAtA1(wb, "=SEQUENCE(3,1,-100,80)");
  const Value spilled = EvalSourceIn("=XIRR(A1#, {40000;40100;40200}, 0.1)", wb, wb.sheet(0));
  ASSERT_TRUE(spilled.is_number()) << "kind=" << static_cast<int>(spilled.kind());
  const Value literal = EvalSource("=XIRR({-100;-20;60}, {40000;40100;40200}, 0.1)");
  ASSERT_TRUE(literal.is_number());
  EXPECT_DOUBLE_EQ(spilled.as_number(), literal.as_number());
}

// ---------------------------------------------------------------------------
// RANK / PERCENTRANK
// ---------------------------------------------------------------------------

TEST(LazyRangeArgShapes, RankMatchesAcrossRefShapes) {
  const Value literal = EvalSource("=RANK(3, {1;2;3;4;5})");
  ASSERT_TRUE(literal.is_number());
  EXPECT_DOUBLE_EQ(literal.as_number(), 3.0);

  const Value dynamic = EvalSource("=RANK(3, SEQUENCE(5))");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, RankAcceptsSpilledRange) {
  Workbook wb = Workbook::create();
  SpillSequenceAtA1(wb);
  const Value v = EvalSourceIn("=RANK(2, A1#)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(LazyRangeArgShapes, RankEqAndAvgMatchAcrossRefShapes) {
  const Value eq_literal = EvalSource("=RANK.EQ(3, {1;2;3;4;5})");
  const Value eq_dynamic = EvalSource("=RANK.EQ(3, SEQUENCE(5))");
  ASSERT_TRUE(eq_literal.is_number());
  ASSERT_TRUE(eq_dynamic.is_number()) << "kind=" << static_cast<int>(eq_dynamic.kind());
  EXPECT_DOUBLE_EQ(eq_dynamic.as_number(), eq_literal.as_number());

  const Value avg_literal = EvalSource("=RANK.AVG(3, {1;2;3;4;5})");
  const Value avg_dynamic = EvalSource("=RANK.AVG(3, SEQUENCE(5))");
  ASSERT_TRUE(avg_literal.is_number());
  ASSERT_TRUE(avg_dynamic.is_number()) << "kind=" << static_cast<int>(avg_dynamic.kind());
  EXPECT_DOUBLE_EQ(avg_dynamic.as_number(), avg_literal.as_number());
}

TEST(LazyRangeArgShapes, PercentRankMatchesAcrossRefShapes) {
  for (const char* fn : {"PERCENTRANK", "PERCENTRANK.INC", "PERCENTRANK.EXC"}) {
    const std::string literal_src = std::string("=") + fn + "({1;2;3;4;5}, 3)";
    const Value literal = EvalSource(literal_src);
    ASSERT_TRUE(literal.is_number()) << literal_src;
    const double expected = literal.as_number();

    const std::string dynamic_src = std::string("=") + fn + "(SEQUENCE(5), 3)";
    const Value dynamic = EvalSource(dynamic_src);
    ASSERT_TRUE(dynamic.is_number()) << dynamic_src;
    EXPECT_DOUBLE_EQ(dynamic.as_number(), expected) << dynamic_src;
  }
}

TEST(LazyRangeArgShapes, PercentRankAcceptsSpilledRange) {
  Workbook wb = Workbook::create();
  SpillSequenceAtA1(wb);
  const Value spilled = EvalSourceIn("=PERCENTRANK.INC(A1#, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(spilled.is_number()) << "kind=" << static_cast<int>(spilled.kind());
  const Value literal = EvalSource("=PERCENTRANK.INC({1;2;3}, 2)");
  ASSERT_TRUE(literal.is_number());
  EXPECT_DOUBLE_EQ(spilled.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, RankStillRejectsBareScalar) {
  const Value v = EvalSource("=RANK(3, 1+2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Regression family (`#N/A` shape vocabulary)
// ---------------------------------------------------------------------------

TEST(LazyRangeArgShapes, CorrelMatchesAcrossArrayShapes) {
  const Value literal = EvalSource("=CORREL({1;2;3}, {1;2;3})");
  ASSERT_TRUE(literal.is_number()) << "kind=" << static_cast<int>(literal.kind());
  EXPECT_NEAR(literal.as_number(), 1.0, 1e-12);

  const Value dynamic = EvalSource("=CORREL(SEQUENCE(3), SEQUENCE(3))");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, SlopeMatchesAcrossArrayShapes) {
  const Value literal = EvalSource("=SLOPE({2;4;6}, {1;2;3})");
  ASSERT_TRUE(literal.is_number());

  const Value dynamic = EvalSource("=SLOPE(SEQUENCE(3,1,2,2), SEQUENCE(3))");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, ForecastMatchesAcrossArrayShapes) {
  const Value literal = EvalSource("=FORECAST.LINEAR(4, {2;4;6}, {1;2;3})");
  ASSERT_TRUE(literal.is_number());

  const Value dynamic = EvalSource("=FORECAST.LINEAR(4, SEQUENCE(3,1,2,2), SEQUENCE(3))");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, RegressionFamilyAcceptsSpilledRange) {
  Workbook wb = Workbook::create();
  SpillSequenceAtA1(wb);
  const Sheet& sheet = wb.sheet(0);

  const Value correl = EvalSourceIn("=CORREL(A1#, A1#)", wb, sheet);
  ASSERT_TRUE(correl.is_number()) << "kind=" << static_cast<int>(correl.kind());
  EXPECT_NEAR(correl.as_number(), 1.0, 1e-12);

  const Value slope = EvalSourceIn("=SLOPE(A1#, A1#)", wb, sheet);
  ASSERT_TRUE(slope.is_number()) << "kind=" << static_cast<int>(slope.kind());
  EXPECT_NEAR(slope.as_number(), 1.0, 1e-12);

  const Value forecast = EvalSourceIn("=FORECAST.LINEAR(4, A1#, A1#)", wb, sheet);
  ASSERT_TRUE(forecast.is_number()) << "kind=" << static_cast<int>(forecast.kind());
  EXPECT_NEAR(forecast.as_number(), 4.0, 1e-12);
}

TEST(LazyRangeArgShapes, CorrelStillRejectsBareScalar) {
  const Value v = EvalSource("=CORREL(5, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Workday holidays slot
// ---------------------------------------------------------------------------

TEST(LazyRangeArgShapes, WorkdayIntlHolidaysMatchAcrossShapes) {
  // 2024-01-01 .. 2024-01-03 as holidays: three consecutive Mondays-to-
  // Wednesday serials 45292, 45293, 45294.
  const Value literal = EvalSource("=WORKDAY.INTL(DATE(2024,1,1), 3, 1, {45292;45293;45294})");
  ASSERT_TRUE(literal.is_number());

  const Value dynamic = EvalSource("=WORKDAY.INTL(DATE(2024,1,1), 3, 1, SEQUENCE(3,1,45292,1))");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, NetworkdaysHolidaysMatchAcrossShapes) {
  const Value literal = EvalSource("=NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,12), {45292;45293;45294})");
  ASSERT_TRUE(literal.is_number());

  const Value dynamic = EvalSource("=NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,12), SEQUENCE(3,1,45292,1))");
  ASSERT_TRUE(dynamic.is_number()) << "kind=" << static_cast<int>(dynamic.kind());
  EXPECT_DOUBLE_EQ(dynamic.as_number(), literal.as_number());
}

TEST(LazyRangeArgShapes, WorkdayHolidaysAcceptSpilledRange) {
  Workbook wb = Workbook::create();
  SpillAtA1(wb, "=SEQUENCE(3,1,45292,1)");
  const Sheet& sheet = wb.sheet(0);

  const Value workday = EvalSourceIn("=WORKDAY.INTL(DATE(2024,1,1), 3, 1, A1#)", wb, sheet);
  ASSERT_TRUE(workday.is_number()) << "kind=" << static_cast<int>(workday.kind());
  const Value workday_literal = EvalSource("=WORKDAY.INTL(DATE(2024,1,1), 3, 1, {45292;45293;45294})");
  ASSERT_TRUE(workday_literal.is_number());
  EXPECT_DOUBLE_EQ(workday.as_number(), workday_literal.as_number());

  const Value networkdays = EvalSourceIn("=NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,12), A1#)", wb, sheet);
  ASSERT_TRUE(networkdays.is_number()) << "kind=" << static_cast<int>(networkdays.kind());
  const Value networkdays_literal = EvalSource("=NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,12), {45292;45293;45294})");
  ASSERT_TRUE(networkdays_literal.is_number());
  EXPECT_DOUBLE_EQ(networkdays.as_number(), networkdays_literal.as_number());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
