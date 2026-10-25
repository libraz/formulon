//
// End-to-end tests for the FORECAST.ETS family: FORECAST.ETS,
// FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, and
// FORECAST.ETS.STAT.
//
// All four functions ride the lazy dispatch table because the values
// and timeline arguments may be Range refs whose `(rows, cols)` shape
// must reach the impl for pairing. They share a Holt-Winters additive
// triple-exponential smoothing core implemented in
// `src/eval/forecast_ets_lazy.cpp`.
//
// SCOPE: the ForecastEtsExact suite pins exact values; the remaining suites
// assert structural behaviour -- error codes, monotonicity (larger confidence
// -> larger half-width), finiteness of stats, and that detection /
// resampling pipelines fire at all -- with generous `EXPECT_NEAR` tolerances.

#include <cfloat>
#include <cmath>
#include <cstddef>
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

// ---------------------------------------------------------------------------
// Test harness
// ---------------------------------------------------------------------------

// Populates a workbook with a linear trend `y = 2*t` for t = 1..10 in
// columns A (timeline) and B (values).
Workbook MakeLinearWorkbook() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (int i = 0; i < 10; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(2.0 * static_cast<double>(i + 1)));
  }
  return wb;
}

// Populates a workbook with 16 quarterly observations exhibiting a
// strict period-4 seasonal pattern on top of a linear trend.
//   t = 1..16, baseline = t, seasonal additive = {+5, -3, -2, 0}.
Workbook MakeQuarterlyWorkbook() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const double pat[4] = {5.0, -3.0, -2.0, 0.0};
  for (int i = 0; i < 16; ++i) {
    const double y = static_cast<double>(i + 1) + pat[i % 4];
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(y));
  }
  return wb;
}

// Workbook with duplicate timeline entries (used to exercise the
// aggregation modes). Two duplicates per timeline value.
Workbook MakeDuplicateTimelineWorkbook() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  // Pairs: (1, 10), (1, 20), (2, 30), (2, 40), (3, 50), (3, 60),
  //        (4, 70), (4, 80), (5, 90), (5, 100).
  for (int i = 0; i < 5; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(2 * i), 0, Value::number(static_cast<double>(i + 1)));
    s.set_cell_value(static_cast<std::uint32_t>(2 * i), 1, Value::number(static_cast<double>(10 * (2 * i + 1))));
    s.set_cell_value(static_cast<std::uint32_t>(2 * i + 1), 0, Value::number(static_cast<double>(i + 1)));
    s.set_cell_value(static_cast<std::uint32_t>(2 * i + 1), 1, Value::number(static_cast<double>(10 * (2 * i + 2))));
  }
  return wb;
}

// ===========================================================================
// FORECAST.ETS basic
// ===========================================================================

TEST(ForecastEts, LinearTrendNoSeasonalityIsFinite) {
  // Force seasonality = 0 (non-seasonal) on a perfect linear series.
  // The forecast should be a finite number very close to the linear
  // extrapolation y = 2*t (because the residuals are all near zero).
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  // Generous tolerance: a damped Holt may not nail the exact 22.0
  // extrapolation but it should be close.
  EXPECT_NEAR(v.as_number(), 22.0, 5.0);
}

TEST(ForecastEts, LinearTrendForecastIsPositive) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(15, B1:B10, A1:A10, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
}

TEST(ForecastEts, ManualSeasonalityFour) {
  const Workbook wb = MakeQuarterlyWorkbook();
  // Forecast at t = 17 with manual seasonality = 4. Should produce a
  // finite number; the expected ballpark is `~17 + 5 = 22` because t=17
  // aligns with the +5 seasonal slot.
  const Value v = EvalSourceIn("=FORECAST.ETS(17, B1:B16, A1:A16, 4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_TRUE(std::isfinite(v.as_number()));
}

TEST(ForecastEts, AutoDetectSeasonalityRunsToCompletion) {
  const Workbook wb = MakeQuarterlyWorkbook();
  // No seasonality argument -> auto-detect; should pick m = 4.
  const Value v = EvalSourceIn("=FORECAST.ETS(17, B1:B16, A1:A16)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
}

TEST(ForecastEts, NonSeasonalOnNoisyLinearReasonableForecast) {
  // y_i = 3*i + small jitter. Forecast at t = 11 should be near 33.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const double jitter[10] = {0.1, -0.1, 0.05, -0.05, 0.0, 0.1, -0.1, 0.05, -0.05, 0.0};
  for (int i = 0; i < 10; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(3.0 * (i + 1) + jitter[i]));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 33.0, 5.0);
}

TEST(ForecastEts, TargetBeforeTimelineStartIsNum) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(0, B1:B10, A1:A10, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEts, LengthMismatchIsNA) {
  Workbook wb = Workbook::create();
  for (int i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    wb.sheet(0).set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(static_cast<double>(2 * (i + 1))));
  }
  // values has 5 rows, timeline 4 rows -> #N/A.
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B5, A1:A4, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(ForecastEts, SinglePointSeriesIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=FORECAST.ETS(2, B1:B1, A1:A1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(ForecastEts, ErrorInValuesPropagates) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (int i = 0; i < 5; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
  }
  s.set_cell_value(0, 1, Value::number(10.0));
  s.set_cell_value(1, 1, Value::error(ErrorCode::Div0));
  s.set_cell_value(2, 1, Value::number(30.0));
  s.set_cell_value(3, 1, Value::number(40.0));
  s.set_cell_value(4, 1, Value::number(50.0));
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B5, A1:A5, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(ForecastEts, ErrorInTimelinePropagates) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::error(ErrorCode::Ref));
  s.set_cell_value(2, 0, Value::number(3.0));
  s.set_cell_value(3, 0, Value::number(4.0));
  s.set_cell_value(4, 0, Value::number(5.0));
  for (int i = 0; i < 5; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(static_cast<double>(2 * (i + 1))));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B5, A1:A5, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(ForecastEts, TextInTimelineDropsThePair) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::text("foo"));
  s.set_cell_value(2, 0, Value::number(3.0));
  s.set_cell_value(3, 0, Value::number(4.0));
  s.set_cell_value(4, 0, Value::number(5.0));
  for (int i = 0; i < 5; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(static_cast<double>(2 * (i + 1))));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B5, A1:A5, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
}

TEST(ForecastEts, ArityUnderIsValueError) {
  const Value v = EvalSource("=FORECAST.ETS(1, {1,2,3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ForecastEts, ArityOverIsValueError) {
  const Value v = EvalSource("=FORECAST.ETS(1, {1,2,3}, {1,2,3}, 0, 1, 1, 999)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ForecastEts, SeasonalityOutOfDomainIsNum) {
  const Workbook wb = MakeLinearWorkbook();
  // Seasonality = 9000 exceeds the 8760 cap -> #NUM!.
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 9000)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEts, NegativeSeasonalityIsNum) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ===========================================================================
// FORECAST.ETS.CONFINT
// ===========================================================================

TEST(ForecastEtsConfint, DefaultConfidenceIsPositive) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.CONFINT(11, B1:B10, A1:A10)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GE(v.as_number(), 0.0);
}

TEST(ForecastEtsConfint, ConfidenceZeroIsZero) {
  // Mac Excel 365 accepts confidence == 0 as a degenerate CI: z =
  // InverseStandardNormal(0.5) = 0, so the half-width = 0 * RMSE * sqrt(h) = 0.
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.CONFINT(11, B1:B10, A1:A10, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(ForecastEtsConfint, TargetInsideTrainingWindowIsNum) {
  // h < 1 (target inside or before the last training point) -> #NUM! per
  // Mac Excel 365.
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.CONFINT(5, B1:B10, A1:A10, 0.95)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsConfint, ConfidenceOneIsNum) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.CONFINT(11, B1:B10, A1:A10, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsConfint, ConfidenceNegativeIsNum) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.CONFINT(11, B1:B10, A1:A10, -0.5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsConfint, LargerConfidenceLargerHalfWidth) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v_low = EvalSourceIn("=FORECAST.ETS.CONFINT(11, B1:B10, A1:A10, 0.5, 0)", wb, wb.sheet(0));
  const Value v_high = EvalSourceIn("=FORECAST.ETS.CONFINT(11, B1:B10, A1:A10, 0.99, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v_low.is_number());
  ASSERT_TRUE(v_high.is_number());
  // Same RMSE / horizon -> z(0.99) > z(0.5), so half-width grows.
  // (Both could be ~0 if RMSE is exactly 0; clamp the comparison
  // accordingly with a strict inequality only when the low value is
  // non-trivial.)
  if (v_low.as_number() > 1e-12) {
    EXPECT_GT(v_high.as_number(), v_low.as_number());
  } else {
    EXPECT_GE(v_high.as_number(), v_low.as_number());
  }
}

TEST(ForecastEtsConfint, LargerHorizonLargerHalfWidth) {
  // y = 3*t + small jitter so RMSE > 0.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const double jitter[10] = {0.5, -0.4, 0.3, -0.2, 0.1, 0.4, -0.3, 0.2, -0.1, 0.0};
  for (int i = 0; i < 10; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(3.0 * (i + 1) + jitter[i]));
  }
  const Value v_h1 = EvalSourceIn("=FORECAST.ETS.CONFINT(11, B1:B10, A1:A10, 0.95, 0)", wb, wb.sheet(0));
  const Value v_h5 = EvalSourceIn("=FORECAST.ETS.CONFINT(15, B1:B10, A1:A10, 0.95, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v_h1.is_number());
  ASSERT_TRUE(v_h5.is_number());
  // sqrt(5) > sqrt(1) so half-width must grow.
  EXPECT_GT(v_h5.as_number(), v_h1.as_number());
}

TEST(ForecastEtsConfint, LengthMismatchIsNA) {
  Workbook wb = Workbook::create();
  for (int i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    wb.sheet(0).set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(static_cast<double>(2 * (i + 1))));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS.CONFINT(6, B1:B5, A1:A4, 0.95, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ===========================================================================
// FORECAST.ETS.SEASONALITY
// ===========================================================================

TEST(ForecastEtsSeasonality, QuarterlyPatternReturnsFour) {
  const Workbook wb = MakeQuarterlyWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.SEASONALITY(B1:B16, A1:A16)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(ForecastEtsSeasonality, FlatLineReturnsZero) {
  // Constant series: ACF on a flat line is 0 (zero variance), so no
  // period is detected -> Mac Excel 365 returns 0.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (int i = 0; i < 10; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(7.0));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS.SEASONALITY(B1:B10, A1:A10)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(ForecastEtsSeasonality, PerfectlyLinearSeriesReturnsZero) {
  // A straight ramp is non-stationary: every lag of the raw series
  // autocorrelates near 1, so scanning it undetrended would report a
  // period. Removing the linear trend first leaves only rounding noise
  // and no period is detected.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (int i = 0; i < 10; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(10.0 + 2.0 * static_cast<double>(i)));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS.SEASONALITY(B1:B10, A1:A10)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(ForecastEtsSeasonality, SeasonalPatternRidingATrendStillReportsItsPeriod) {
  // Detrending must not swallow a real period: a period-4 cycle laid over
  // a rising ramp still has to surface as 4.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  static constexpr double kCycle[] = {0.0, 6.0, 0.0, -6.0};
  for (int i = 0; i < 16; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
    const double level = 100.0 + 3.0 * static_cast<double>(i);
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(level + kCycle[i % 4]));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS.SEASONALITY(B1:B16, A1:A16)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(ForecastEtsSeasonality, ShortSeriesReturnsZero) {
  // n = 3 -> below the n >= 4 threshold; detector reports "no period"
  // which Mac Excel 365 surfaces as 0.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::number(2.0));
  s.set_cell_value(2, 0, Value::number(3.0));
  s.set_cell_value(0, 1, Value::number(10.0));
  s.set_cell_value(1, 1, Value::number(20.0));
  s.set_cell_value(2, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=FORECAST.ETS.SEASONALITY(B1:B3, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(ForecastEtsSeasonality, ErrorPropagates) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (int i = 0; i < 5; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
  }
  s.set_cell_value(0, 1, Value::number(10.0));
  s.set_cell_value(1, 1, Value::error(ErrorCode::NA));
  s.set_cell_value(2, 1, Value::number(30.0));
  s.set_cell_value(3, 1, Value::number(40.0));
  s.set_cell_value(4, 1, Value::number(50.0));
  const Value v = EvalSourceIn("=FORECAST.ETS.SEASONALITY(B1:B5, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(ForecastEtsSeasonality, LengthMismatchIsNA) {
  Workbook wb = Workbook::create();
  for (int i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(i + 1)));
  }
  for (int i = 0; i < 4; ++i) {
    wb.sheet(0).set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS.SEASONALITY(B1:B4, A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(ForecastEtsSeasonality, ArityUnderIsValueError) {
  const Value v = EvalSource("=FORECAST.ETS.SEASONALITY({1,2,3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ===========================================================================
// FORECAST.ETS.STAT
// ===========================================================================

TEST(ForecastEtsStat, AlphaIsFiniteInUnitInterval) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_LT(v.as_number(), 1.0);
}

TEST(ForecastEtsStat, BetaIsFiniteInUnitInterval) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_LT(v.as_number(), 1.0);
}

TEST(ForecastEtsStat, GammaIsMachineEpsilonForNonSeasonal) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), DBL_EPSILON);
}

TEST(ForecastEtsStat, GammaInUnitIntervalForSeasonal) {
  const Workbook wb = MakeQuarterlyWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B16, A1:A16, 3, 4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_LT(v.as_number(), 1.0);
}

TEST(ForecastEtsStat, MaseFinite) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 4, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GE(v.as_number(), 0.0);
}

TEST(ForecastEtsStat, SmapeFinite) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 5, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GE(v.as_number(), 0.0);
}

TEST(ForecastEtsStat, MaeNonNegative) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 6, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GE(v.as_number(), 0.0);
}

TEST(ForecastEtsStat, RmseNonNegative) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 7, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
  EXPECT_GE(v.as_number(), 0.0);
}

TEST(ForecastEtsStat, StepSizeMatchesTimelineDelta) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 8, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(ForecastEtsStat, StatTypeZeroIsNum) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 0, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsStat, StatTypeNineIsNum) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 9, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsStat, StatTypeNegativeIsNum) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, -1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsStat, StatTypeFractionalTruncatesIntoOne) {
  // 1.5 truncates to 1 -> alpha. Should be valid.
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 1.5, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
}

TEST(ForecastEtsStat, StatTypeTextIsValueError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, \"foo\", 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ===========================================================================
// Cross-cutting: aggregation, data_completion
// ===========================================================================

TEST(ForecastEtsCrossCut, AggregationAverageRunsToCompletion) {
  const Workbook wb = MakeDuplicateTimelineWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B10, A1:A10, 0, 1, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
}

TEST(ForecastEtsCrossCut, AggregationCountRunsToCompletion) {
  const Workbook wb = MakeDuplicateTimelineWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B10, A1:A10, 0, 1, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
}

TEST(ForecastEtsCrossCut, AggregationCountARunsToCompletion) {
  const Workbook wb = MakeDuplicateTimelineWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B10, A1:A10, 0, 1, 3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
}

TEST(ForecastEtsCrossCut, AggregationMaxRunsToCompletion) {
  const Workbook wb = MakeDuplicateTimelineWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B10, A1:A10, 0, 1, 4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
}

TEST(ForecastEtsCrossCut, AggregationMedianRunsToCompletion) {
  const Workbook wb = MakeDuplicateTimelineWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B10, A1:A10, 0, 1, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
}

TEST(ForecastEtsCrossCut, AggregationMinRunsToCompletion) {
  const Workbook wb = MakeDuplicateTimelineWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B10, A1:A10, 0, 1, 6)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
}

TEST(ForecastEtsCrossCut, AggregationSumRunsToCompletion) {
  const Workbook wb = MakeDuplicateTimelineWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(6, B1:B10, A1:A10, 0, 1, 7)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
}

TEST(ForecastEtsCrossCut, AggregationOutOfDomainHighIsNumError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0, 1, 8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsCrossCut, AggregationOutOfDomainLowIsNumError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0, 1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsCrossCut, DataCompletionInterpolateVsZeroFillDiffer) {
  // Build a linear-ish series with ONE missing slot (gap at t=5).
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  // Pairs: (1,1), (2,2), (3,3), (4,4), -- gap at 5 --, (6,6), (7,7), (8,8),
  //        (9,9), (10,10).
  const int rows[9] = {1, 2, 3, 4, 6, 7, 8, 9, 10};
  for (int i = 0; i < 9; ++i) {
    s.set_cell_value(static_cast<std::uint32_t>(i), 0, Value::number(static_cast<double>(rows[i])));
    s.set_cell_value(static_cast<std::uint32_t>(i), 1, Value::number(static_cast<double>(rows[i])));
  }
  const Value v_interp = EvalSourceIn("=FORECAST.ETS(11, B1:B9, A1:A9, 0, 1)", wb, wb.sheet(0));
  const Value v_zero = EvalSourceIn("=FORECAST.ETS(11, B1:B9, A1:A9, 0, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v_interp.is_number());
  ASSERT_TRUE(v_zero.is_number());
  // With the interpolated gap the trend stays continuous; with zero-fill
  // the y at t=5 is 0 which drags the level down. The two forecasts
  // should differ.
  EXPECT_NE(v_interp.as_number(), v_zero.as_number());
}

TEST(ForecastEtsCrossCut, DataCompletionTwoIsNumError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ForecastEtsCrossCut, DataCompletionNegativeIsNumError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// Exact values: fixed candidate table, holdout scoring and short-series padding
// ---------------------------------------------------------------------------

double EvalExact(const std::string& formula) {
  const Value v = EvalSource(formula);
  EXPECT_TRUE(v.is_number()) << formula;
  return v.is_number() ? v.as_number() : 0.0;
}

struct ExactSeries {
  std::string values;
  std::string timeline;
  std::string season;
  std::size_t n;
};

ExactSeries MakeExactSeries(const char* values, std::size_t n, int season) {
  ExactSeries s;
  s.values = values;
  s.n = n;
  s.season = std::to_string(season);
  s.timeline = "{";
  for (std::size_t i = 1; i <= n; ++i) {
    s.timeline += (i == 1 ? "" : ";") + std::to_string(i);
  }
  s.timeline += "}";
  return s;
}

std::string Args(const ExactSeries& s) {
  return s.values + "," + s.timeline;
}

TEST(ForecastEtsExact, PeriodFourSeries) {
  const ExactSeries s = MakeExactSeries("{10;12;14;11;12;14;16;13;14;16;18;15;16;18;20;17}", 16, 4);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",1," + s.season + ")"), 0.751);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",2," + s.season + ")"), 0.001);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",3," + s.season + ")"), 0.001);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",4," + s.season + ")"), 0.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",5," + s.season + ")"), 0.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",6," + s.season + ")"), 0.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",7," + s.season + ")"), 0.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(17," + Args(s) + "," + s.season + ")"), 18.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(18," + Args(s) + "," + s.season + ")"), 20.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(21," + Args(s) + "," + s.season + ")"), 20.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(17," + Args(s) + ",0.95," + s.season + ")"), 0.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(19," + Args(s) + ",0.95," + s.season + ")"), 0.0);
}

TEST(ForecastEtsExact, SixPhaseSeries) {
  const ExactSeries s = MakeExactSeries(
      "{50.38;61.96;50.78;51.37;45.24;52.03;52.62;63.9;53.16;52.93;47.52;55.28;53.73;65.14;54.95;55.82;49.48;56.48;56."
      "4;67.0;57.8;57.46;50.88;57.03}",
      24, 6);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",1," + s.season + ")"), 0.126);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",2," + s.season + ")"), 0.001);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",3," + s.season + ")"), 0.001);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",4," + s.season + ")"), 0.0861558112623309);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",5," + s.season + ")"), 0.010308666468835702);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",6," + s.season + ")"), 0.5872749334831597);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",7," + s.season + ")"), 0.7355763696204374);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(25," + Args(s) + "," + s.season + ")"), 57.85321655080211);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(26," + Args(s) + "," + s.season + ")"), 69.19397936760204);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(29," + Args(s) + "," + s.season + ")"), 53.05027053900617);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(25," + Args(s) + ",0.95," + s.season + ")"), 1.0536375883485813);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(27," + Args(s) + ",0.95," + s.season + ")"), 1.070629041734613);
}

TEST(ForecastEtsExact, PaddedShortSeries) {
  const ExactSeries s = MakeExactSeries("{94.61;48.87;19.39;94.6;57.9;72.89;88.1;28.57;35.67}", 9, 3);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",1," + s.season + ")"), 0.002);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",2," + s.season + ")"), 0.001);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",3," + s.season + ")"), 0.25);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",4," + s.season + ")"), 0.3136113631213259);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",5," + s.season + ")"), 0.25103157995311975);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",6," + s.season + ")"), 11.131635333991463);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + Args(s) + ",7," + s.season + ")"), 19.664579855001158);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(10," + Args(s) + "," + s.season + ")"), 89.02829274003653);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(11," + Args(s) + "," + s.season + ")"), 43.627416296623636);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(14," + Args(s) + "," + s.season + ")"), 41.937864018855905);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(10," + Args(s) + ",0.95," + s.season + ")"), 38.54186828691414);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(12," + Args(s) + ",0.95," + s.season + ")"), 38.542350057256684);
}

// ---------------------------------------------------------------------------
// Exact values with the seasonality chosen by the detector, and short series
// ---------------------------------------------------------------------------

const std::string kYnoisy12 = "{20.68;23.387;21.379;23.614;22.256;23.242;25.997;24.128;25.926;25.877;26.359;26.985}";
const std::string kTnoisy12 = "{1;2;3;4;5;6;7;8;9;10;11;12}";
const std::string kYrand30 =
    "{87.8;13.5;76.4;9.8;69.0;70.2;95.0;84.3;50.4;19.8;15.0;52.9;51.0;7.1;90.3;50.7;70.1;22.0;24.4;1.2;34.3;26.7;42.3;"
    "37.7;83.4;89.1;17.5;39.6;16.6;66.3}";
const std::string kTrand30 = "{1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30}";
const std::string kYrand9 = "{94.61;48.87;19.39;94.6;57.9;72.89;88.1;28.57;35.67}";
const std::string kTrand9 = "{1;2;3;4;5;6;7;8;9}";

TEST(ForecastEtsExact, NoisyAutoDetect) {
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYnoisy12 + "," + kTnoisy12 + ",1)"), 0.002);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYnoisy12 + "," + kTnoisy12 + ",4)"), 0.1301356352922273);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYnoisy12 + "," + kTnoisy12 + ",7)"), 0.41936220161199983);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(13," + kYnoisy12 + "," + kTnoisy12 + ")"), 27.29427295578445);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(13," + kYnoisy12 + "," + kTnoisy12 + ",0.95)"), 1.9047510699056267);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYnoisy12 + "," + kTnoisy12 + ",1,1)"), 0.002);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYnoisy12 + "," + kTnoisy12 + ",4,1)"), 0.1301356352922273);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYnoisy12 + "," + kTnoisy12 + ",7,1)"), 0.41936220161199983);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(13," + kYnoisy12 + "," + kTnoisy12 + ",1)"), 27.29427295578445);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(13," + kYnoisy12 + "," + kTnoisy12 + ",0.95,1)"), 1.9047510699056267);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.SEASONALITY(" + kYnoisy12 + "," + kTnoisy12 + ")"), 2.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand9 + "," + kTrand9 + ",1)"), 0.002);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand9 + "," + kTrand9 + ",4)"), 0.3136113631213259);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand9 + "," + kTrand9 + ",7)"), 19.664579855001158);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(10," + kYrand9 + "," + kTrand9 + ")"), 89.02829274003653);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(10," + kYrand9 + "," + kTrand9 + ",0.95)"), 38.54186828691414);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand9 + "," + kTrand9 + ",1,1)"), 0.002);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand9 + "," + kTrand9 + ",4,1)"), 0.3136113631213259);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand9 + "," + kTrand9 + ",7,1)"), 19.664579855001158);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(10," + kYrand9 + "," + kTrand9 + ",1)"), 89.02829274003653);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(10," + kYrand9 + "," + kTrand9 + ",0.95,1)"), 38.54186828691414);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.SEASONALITY(" + kYrand9 + "," + kTrand9 + ")"), 3.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand30 + "," + kTrand30 + ",1)"), 0.501);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand30 + "," + kTrand30 + ",4)"), 0.6834746824287844);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYrand30 + "," + kTrand30 + ",7)"), 37.05470096447273);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(31," + kYrand30 + "," + kTrand30 + ")"), 24.533673084551403);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(31," + kYrand30 + "," + kTrand30 + ",0.95)"), 64.5715935421549);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.SEASONALITY(" + kYrand30 + "," + kTrand30 + ")"), 10.0);
}

const std::string kYp7_28 =
    "{49.225;56.546;59.162;55.085;44.201;43.854;43.843;52.649;60.164;63.224;56.97;49.659;44.543;46.637;53.714;63.522;"
    "64.664;59.804;51.428;47.146;46.938;57.348;64.757;65.672;62.848;54.104;48.086;49.278}";
const std::string kTp7_28 = "{1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28}";
const std::string kYspike_30 =
    "{0.0;0.0;0.0;0.0;0.0;10.0;0.0;0.0;0.0;0.0;0.0;0.0;10.0;0.0;0.0;0.0;0.0;0.0;0.0;10.0;0.0;0.0;0.0;0.0;0.0;0.0;10.0;"
    "0.0;0.0;0.0}";
const std::string kTspike_30 = "{1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30}";
const std::string kYtwo_peaks_48 =
    "{50.374;55.809;61.589;59.159;56.418;51.596;48.682;49.732;53.083;54.854;54.747;52.795;47.852;45.434;45.427;48.353;"
    "54.909;60.608;63.877;64.177;61.354;56.261;54.51;54.57;57.126;59.529;60.164;57.214;53.434;50.373;50.204;53.561;59."
    "173;65.135;69.881;68.713;66.205;61.962;58.973;60.332;61.243;64.511;64.629;61.881;57.915;54.373;54.385;58.358}";
const std::string kTtwo_peaks_48 =
    "{1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30;31;32;33;34;35;36;37;38;39;40;"
    "41;42;43;44;45;46;47;48}";

TEST(ForecastEtsExact, DetectedPeriod) {
  EXPECT_EQ(EvalExact("=FORECAST.ETS.SEASONALITY(" + kYp7_28 + "," + kTp7_28 + ")"), 7.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(29," + kYp7_28 + "," + kTp7_28 + ")"), 58.49976744730409);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYp7_28 + "," + kTp7_28 + ",1)"), 0.126);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.SEASONALITY(" + kYspike_30 + "," + kTspike_30 + ")"), 7.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(31," + kYspike_30 + "," + kTspike_30 + ")"), 0.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYspike_30 + "," + kTspike_30 + ",1)"), 0.751);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.SEASONALITY(" + kYtwo_peaks_48 + "," + kTtwo_peaks_48 + ")"), 16.0);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(49," + kYtwo_peaks_48 + "," + kTtwo_peaks_48 + ")"), 64.28188450452633);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYtwo_peaks_48 + "," + kTtwo_peaks_48 + ",1)"), 0.002);
}

const std::string kYsh_3 = "{6.71;6.54;1.35}";
const std::string kTsh_3 = "{1;2;3}";
const std::string kYsh_7 = "{7.69;6.36;1.15;2.17;8.99;8.74;7.28}";
const std::string kTsh_7 = "{1;2;3;4;5;6;7}";
const std::string kYshort_5 = "{5.11;7.67;0.84;9.2;4.34}";
const std::string kTshort_5 = "{1;2;3;4;5}";

TEST(ForecastEtsExact, ShortSeries) {
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYshort_5 + "," + kTshort_5 + ",1,0)"), 0.002);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYshort_5 + "," + kTshort_5 + ",4,0)"), 0.4229819009585671);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYshort_5 + "," + kTshort_5 + ",6,0)"), 2.3909051951683002);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(6," + kYshort_5 + "," + kTshort_5 + ",0)"), 5.733417349250267);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(6," + kYshort_5 + "," + kTshort_5 + ",0.95,0)"), 5.672440470422527);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYshort_5 + "," + kTshort_5 + ",1,2)"), 0.501);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYshort_5 + "," + kTshort_5 + ",4,2)"), 0.28097426346041565);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYshort_5 + "," + kTshort_5 + ",6,2)"), 1.5882070242099995);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(6," + kYshort_5 + "," + kTshort_5 + ",2)"), 9.514455932931389);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(6," + kYshort_5 + "," + kTshort_5 + ",0.95,2)"), 4.170048436094756);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",1,0)"), 0.1);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",4,0)"), 0.37431467661691564);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",6,0)"), 1.0031633333333339);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(4," + kYsh_3 + "," + kTsh_3 + ",0)"), -0.6814185099999984);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(4," + kYsh_3 + "," + kTsh_3 + ",0.95,0)"), 2.8959731396425554);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",1,2)"), 0.1);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",4,2)"), 0.37431467661691564);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",6,2)"), 1.0031633333333339);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(4," + kYsh_3 + "," + kTsh_3 + ",2)"), -0.6814185099999984);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(4," + kYsh_3 + "," + kTsh_3 + ",0.95,2)"), 2.8959731396425554);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",1,3)"), 0.1);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",4,3)"), 0.37431467661691564);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_3 + "," + kTsh_3 + ",6,3)"), 1.0031633333333339);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(4," + kYsh_3 + "," + kTsh_3 + ",3)"), -0.6814185099999984);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(4," + kYsh_3 + "," + kTsh_3 + ",0.95,3)"), 2.8959731396425554);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",1,0)"), 0.9);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",4,0)"), 0.8454786991922306);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",6,0)"), 2.2672920450004983);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(8," + kYsh_7 + "," + kTsh_7 + ",0)"), 7.869761864231694);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(8," + kYsh_7 + "," + kTsh_7 + ",0.95,0)"), 6.67902896569191);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",1,2)"), 0.501);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",4,2)"), 1.1066441576936334);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",6,2)"), 2.9676507495484272);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(8," + kYsh_7 + "," + kTsh_7 + ",2)"), 9.52261072238843);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(8," + kYsh_7 + "," + kTsh_7 + ",0.95,2)"), 7.52462699197498);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",1,3)"), 0.501);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",4,3)"), 1.2048543075640552);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT(" + kYsh_7 + "," + kTsh_7 + ",6,3)"), 3.231017634784275);
  EXPECT_EQ(EvalExact("=FORECAST.ETS(8," + kYsh_7 + "," + kTsh_7 + ",3)"), 11.92090460974675);
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(8," + kYsh_7 + "," + kTsh_7 + ",0.95,3)"), 7.193962107036191);
}

// ---------------------------------------------------------------------------
// Timeline normalisation: ordering, grouping, gaps, calendar steps, targets
// ---------------------------------------------------------------------------

ErrorCode EvalError(const std::string& formula) {
  const Value v = EvalSource(formula);
  EXPECT_TRUE(v.is_error()) << formula;
  return v.is_error() ? v.as_error() : ErrorCode::Null;
}

TEST(ForecastEtsTimeline, GroupedDuplicates) {
  EXPECT_EQ(EvalExact("=FORECAST.ETS(11.0,{14.5;12.25;10.5;11.0;13.75;12.5;11.5;16.0;18.0;15.5;19.0;21.5},{5.0;3.0;1.0;"
                      "2.0;3.0;4.0;3.0;6.0;7.0;8.0;9.0;10.0},0,1,1)"),
            21.54142795852091);  // agg1
  EXPECT_EQ(EvalExact("=FORECAST.ETS(11.0,{14.5;12.25;10.5;11.0;13.75;12.5;11.5;16.0;18.0;15.5;19.0;21.5},{5.0;3.0;1.0;"
                      "2.0;3.0;4.0;3.0;6.0;7.0;8.0;9.0;10.0},0,1,2)"),
            0.7544725192564488);  // agg2
  EXPECT_EQ(EvalExact("=FORECAST.ETS(11.0,{14.5;12.25;10.5;11.0;13.75;12.5;11.5;16.0;18.0;15.5;19.0;21.5},{5.0;3.0;1.0;"
                      "2.0;3.0;4.0;3.0;6.0;7.0;8.0;9.0;10.0},0,1,3)"),
            0.7544725192564488);  // agg3
  EXPECT_EQ(EvalExact("=FORECAST.ETS(11.0,{14.5;12.25;10.5;11.0;13.75;12.5;11.5;16.0;18.0;15.5;19.0;21.5},{5.0;3.0;1.0;"
                      "2.0;3.0;4.0;3.0;6.0;7.0;8.0;9.0;10.0},0,1,4)"),
            21.497995309190685);  // agg4
  EXPECT_EQ(EvalExact("=FORECAST.ETS(11.0,{14.5;12.25;10.5;11.0;13.75;12.5;11.5;16.0;18.0;15.5;19.0;21.5},{5.0;3.0;1.0;"
                      "2.0;3.0;4.0;3.0;6.0;7.0;8.0;9.0;10.0},0,1,5)"),
            21.57703172041051);  // agg5
  EXPECT_EQ(EvalExact("=FORECAST.ETS(11.0,{14.5;12.25;10.5;11.0;13.75;12.5;11.5;16.0;18.0;15.5;19.0;21.5},{5.0;3.0;1.0;"
                      "2.0;3.0;4.0;3.0;6.0;7.0;8.0;9.0;10.0},0,1,6)"),
            21.626440799021236);  // agg6
  EXPECT_EQ(EvalExact("=FORECAST.ETS(11.0,{14.5;12.25;10.5;11.0;13.75;12.5;11.5;16.0;18.0;15.5;19.0;21.5},{5.0;3.0;1.0;"
                      "2.0;3.0;4.0;3.0;6.0;7.0;8.0;9.0;10.0},0,1,7)"),
            18.434583549355068);  // agg7
  EXPECT_EQ(EvalExact("=FORECAST.ETS.STAT({14.5;12.25;10.5;11.0;13.75;12.5;11.5;16.0;18.0;15.5;19.0;21.5},{5.0;3.0;1.0;"
                      "2.0;3.0;4.0;3.0;6.0;7.0;8.0;9.0;10.0},8,0,1,1)"),
            1.0);  // unsorted_stat8
}

TEST(ForecastEtsTimeline, GapCompletion) {
  EXPECT_EQ(EvalExact("=FORECAST.ETS(13.0,{10.0;12.0;14.0;13.0;17.0;19.0;18.0;21.0;23.0;22.0},{1.0;2.0;3.0;4.0;7.0;8.0;"
                      "9.0;10.0;11.0;12.0},0,1,1)"),
            23.797998489528826);  // gap_c1
  EXPECT_EQ(EvalExact("=FORECAST.ETS(13.0,{10.0;12.0;14.0;13.0;17.0;19.0;18.0;21.0;23.0;22.0},{1.0;2.0;3.0;4.0;7.0;8.0;"
                      "9.0;10.0;11.0;12.0},0,0,1)"),
            23.546862593009674);  // gap_c0
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(13.0,{10.0;12.0;14.0;13.0;17.0;19.0;18.0;21.0;23.0;22.0},{1.0;2.0;3.0;4.0;"
                      "7.0;8.0;9.0;10.0;11.0;12.0},0.95,0,1,1)"),
            1.8127789621549355);  // gap_conf
}

TEST(ForecastEtsTimeline, CalendarSteps) {
  EXPECT_EQ(
      EvalExact(
          "=FORECAST.ETS(45292.0,{100.0;104.0;109.0;103.0;108.0;114.0;107.0;113.0;119.0;111.0;118.0;125.0},{44927.0;"
          "44958.0;44986.0;45017.0;45047.0;45078.0;45108.0;45139.0;45170.0;45200.0;45231.0;45261.0},0,1,1)"),
      122.06854734505528);  // monthly_next
  EXPECT_EQ(
      EvalExact("=FORECAST.ETS.STAT({100.0;104.0;109.0;103.0;108.0;114.0;107.0;113.0;119.0;111.0;118.0;125.0},{44927.0;"
                "44958.0;44986.0;45017.0;45047.0;45078.0;45108.0;45139.0;45170.0;45200.0;45231.0;45261.0},8,0,1,1)"),
      31.0);  // monthly_step
  EXPECT_EQ(
      EvalExact(
          "=FORECAST.ETS(45307.0,{100.0;104.0;109.0;103.0;108.0;114.0;107.0;113.0;119.0;111.0;118.0;125.0},{44927.0;"
          "44958.0;44986.0;45017.0;45047.0;45078.0;45108.0;45139.0;45170.0;45200.0;45231.0;45261.0},0,1,1)"),
      122.92389789279682);  // monthly_mid
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(45323.0,{100.0;104.0;109.0;103.0;108.0;114.0;107.0;113.0;119.0;111.0;118."
                      "0;125.0},{44927.0;44958.0;44986.0;45017.0;45047.0;45078.0;45108.0;45139.0;45170.0;45200.0;45231."
                      "0;45261.0},0.95,0,1,1)"),
            7.309925665120638);  // monthly_conf
  EXPECT_EQ(
      EvalExact(
          "=FORECAST.ETS(45322.0,{100.0;104.0;109.0;103.0;108.0;114.0;107.0;113.0;119.0;111.0;118.0;125.0},{44957.0;"
          "44985.0;45016.0;45046.0;45077.0;45107.0;45138.0;45169.0;45199.0;45230.0;45260.0;45291.0},0,1,1)"),
      122.06854734505528);  // monthend_next
  EXPECT_EQ(
      EvalExact("=FORECAST.ETS(45322.0,{100.0;104.0;109.0;103.0;114.0;107.0;113.0;119.0;111.0;118.0;125.0},{44957.0;"
                "44985.0;45016.0;45046.0;45107.0;45138.0;45169.0;45199.0;45230.0;45260.0;45291.0},0,1,1)"),
      122.05762809693425);  // monthend_gap
}

TEST(ForecastEtsTimeline, FractionalTargets) {
  EXPECT_EQ(EvalExact("=FORECAST.ETS(13.5,{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{1.0;2.0;3.0;4."
                      "0;5.0;6.0;7.0;8.0;9.0;10.0;11.0;12.0},0,1,1)"),
            22.56794634344145);  // frac_ets
  EXPECT_EQ(EvalExact("=FORECAST.ETS.CONFINT(13.5,{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{1.0;2."
                      "0;3.0;4.0;5.0;6.0;7.0;8.0;9.0;10.0;11.0;12.0},0.95,0,1,1)"),
            1.9490820308059253);  // frac_conf
  EXPECT_EQ(EvalExact("=FORECAST.ETS(3.0,{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{1.0;2.0;3.0;4."
                      "0;5.0;6.0;7.0;8.0;9.0;10.0;11.0;12.0},0,1,1)"),
            12.003240904940714);  // in_window_target
}

TEST(ForecastEtsTimeline, RejectedInputs) {
  EXPECT_EQ(EvalError("=FORECAST.ETS(13.0,{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{1.0;2.0;3.0;4."
                      "0;5.0;6.0;7.0;8.0;9.0;10.0;11.0},0,1,1)"),
            ErrorCode::NA);                                                 // mismatch
  EXPECT_EQ(EvalError("=FORECAST.ETS(2.0,{1.0},{1.0})"), ErrorCode::Div0);  // one_point
  EXPECT_EQ(EvalError("=FORECAST.ETS(13.0,{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{1.0;2.0;3.0;4."
                      "0;5.0;6.0;7.0;8.0;9.0;10.0;11.0;12.0},0,2,1)"),
            ErrorCode::Num);  // bad_completion
  EXPECT_EQ(EvalError("=FORECAST.ETS(13.0,{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{1.0;2.0;3.0;4."
                      "0;5.0;6.0;7.0;8.0;9.0;10.0;11.0;12.0},0,1,8)"),
            ErrorCode::Num);  // bad_agg
  EXPECT_EQ(EvalError("=FORECAST.ETS(13.0,{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{5.0;5.0;5.0;5."
                      "0;5.0;5.0;5.0;5.0;5.0;5.0;5.0;5.0},0,1,1)"),
            ErrorCode::Num);  // same_times
}

TEST(ForecastEtsTimeline, BlankAndTextValueCells) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const double ys[10] = {10.0, 12.0, 0.0, 14.0, 0.0, 16.0, 15.0, 18.0, 17.0, 20.0};
  for (std::uint32_t i = 0; i < 10; ++i) {
    s.set_cell_value(i, 1, Value::number(static_cast<double>(i + 1)));
    if (i != 2U && i != 4U) {
      s.set_cell_value(i, 0, Value::number(ys[i]));
    }
  }
  s.set_cell_value(4, 0, Value::text("x"));
  const auto eval = [&](const char* f) {
    const Value v = EvalSourceIn(f, wb, wb.sheet(0));
    EXPECT_TRUE(v.is_number()) << f;
    return v.is_number() ? v.as_number() : 0.0;
  };
  EXPECT_EQ(eval("=FORECAST.ETS(11, A1:A10, B1:B10, 0, 1)"), 20.09714968702844);
  EXPECT_EQ(eval("=FORECAST.ETS(11, A1:A10, B1:B10, 0, 0)"), 21.054355325337056);
  EXPECT_EQ(eval("=FORECAST.ETS(11, A1:A10, B1:B10, 0, 1, 3)"), 1.0);
}

TEST(ForecastEtsSeasonality, DetectorWindowIsLargestPowerOfTwoTail) {
  // Period 7 for the first 1500 points, period 5 after. Only the last 2048
  // points are analysed, so the period is 7; a 1024-point tail would give none.
  constexpr std::uint32_t kPoints = 2100;
  constexpr std::uint32_t kSplit = 1500;
  const double heavy[7] = {3.0, -1.0, 2.0, -4.0, 0.0, 5.0, -3.0};
  const double light[5] = {4.0, -2.0, 1.0, -3.0, 0.0};
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  std::uint32_t state = 12345U;
  for (std::uint32_t i = 0; i < kPoints; ++i) {
    state = state * 1664525U + 1013904223U;
    const double noise = 0.3 * static_cast<double>(state >> 8U) / 16777216.0;
    const double pattern = i < kSplit ? heavy[i % 7U] : light[i % 5U];
    s.set_cell_value(i, 0, Value::number(50.0 + pattern + noise));
    s.set_cell_value(i, 1, Value::number(static_cast<double>(i + 1U)));
  }
  const Value v = EvalSourceIn("=FORECAST.ETS.SEASONALITY(A1:A2100, B1:B2100)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

// ---------------------------------------------------------------------------
// Far horizons: more than 2^31 - 1 steps past the last node is #NUM!
// ---------------------------------------------------------------------------

const std::string kMonthlyArgs =
    "{100.0;104.0;109.0;103.0;108.0;114.0;107.0;113.0;119.0;111.0;118.0;125.0},{44927.0;44958.0;44986.0;45017.0;"
    "45047.0;45078.0;45108.0;45139.0;45170.0;45200.0;45231.0;45261.0}";
const std::string kYearlyArgs =
    "{100.0;104.0;109.0;103.0;108.0;114.0;107.0;113.0;119.0;111.0;118.0;125.0},{40179.0;40544.0;40909.0;41275.0;"
    "41640.0;42005.0;42370.0;42736.0;43101.0;43466.0;43831.0;44197.0}";
const std::string kDailyArgs =
    "{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{1.0;2.0;3.0;4.0;5.0;6.0;7.0;8.0;9.0;10.0;11.0;"
    "12.0}";

TEST(ForecastEtsFarTarget, CalendarTimelinesTerminate) {
  for (const std::string& args : {kMonthlyArgs, kYearlyArgs}) {
    for (const char* target : {"3E+6"}) {
      const Value point = EvalSource("=FORECAST.ETS(" + std::string(target) + "," + args + ",0)");
      EXPECT_TRUE(point.is_number()) << target;
      const Value conf = EvalSource("=FORECAST.ETS.CONFINT(" + std::string(target) + "," + args + ",0.95,0)");
      ASSERT_TRUE(conf.is_number()) << target;
      EXPECT_GT(conf.as_number(), 0.0) << target;
    }
    // 1E+15 used to walk the calendar one node at a time and never return.
    for (const char* target : {"1E+15", "1E+300", "-1E+300", "2^53", "-1E+15"}) {
      EXPECT_EQ(EvalError("=FORECAST.ETS(" + std::string(target) + "," + args + ",0)"), ErrorCode::Num) << target;
      EXPECT_EQ(EvalError("=FORECAST.ETS.CONFINT(" + std::string(target) + "," + args + ",0.95,0)"), ErrorCode::Num)
          << target;
    }
  }
}

TEST(ForecastEtsFarTarget, DailyHorizonBoundary) {
  // Last node 12, step 1: 2147483659 is 2^31 - 1 steps out and still answers.
  EXPECT_TRUE(EvalSource("=FORECAST.ETS(2147483659," + kDailyArgs + ",0)").is_number());
  const Value conf = EvalSource("=FORECAST.ETS.CONFINT(2147483659," + kDailyArgs + ",0.95,0)");
  ASSERT_TRUE(conf.is_number());
  EXPECT_GT(conf.as_number(), 0.0);
  // The bound counts steps from the last node, not date serials.
  const std::string negative =
      "{10.0;12.5;11.0;14.0;13.5;16.0;15.0;18.5;17.0;20.0;19.5;22.0},{-12.0;-11.0;-10.0;-9.0;-8.0;-7.0;-6.0;-5.0;-4.0;"
      "-3.0;-2.0;-1.0}";
  EXPECT_TRUE(EvalSource("=FORECAST.ETS(1E+9," + negative + ",0)").is_number());
  EXPECT_EQ(EvalError("=FORECAST.ETS(3E+9," + negative + ",0)"), ErrorCode::Num);
  // Last node -1: 2147483646 is 2^31 - 1 steps out, 2147483647 is one past.
  EXPECT_TRUE(EvalSource("=FORECAST.ETS(2147483646," + negative + ",0)").is_number());
  EXPECT_EQ(EvalError("=FORECAST.ETS(2147483647," + negative + ",0)"), ErrorCode::Num);
  for (const char* target : {"2147483660", "3E+9", "1E+15", "2^53", "1E+18", "5E+18", "1E+300"}) {
    EXPECT_EQ(EvalError("=FORECAST.ETS(" + std::string(target) + "," + kDailyArgs + ",0)"), ErrorCode::Num) << target;
    EXPECT_EQ(EvalError("=FORECAST.ETS.CONFINT(" + std::string(target) + "," + kDailyArgs + ",0.95,0)"), ErrorCode::Num)
        << target;
  }
}

TEST(ForecastEtsFarTarget, StatTypeOutOfRangeIsNum) {
  for (const char* stat : {"1E+300", "-1E+300", "2^63", "-2^63"}) {
    EXPECT_EQ(EvalError("=FORECAST.ETS.STAT(" + kDailyArgs + "," + std::string(stat) + ",0)"), ErrorCode::Num) << stat;
  }
}

// ---------------------------------------------------------------------------
// Array target_date: one result per element, each equal to the scalar call
// ---------------------------------------------------------------------------

TEST(ForecastEtsArrayTarget, ForecastLiftsOverTargets) {
  const double at13 = EvalExact("=FORECAST.ETS(13.0," + kDailyArgs + ",0)");
  const double at14 = EvalExact("=FORECAST.ETS(14.5," + kDailyArgs + ",0)");
  const Value v = EvalSource("=FORECAST.ETS({13.0,14.5,\"x\",0.0}," + kDailyArgs + ",0)");
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array()->rows, 1U);
  ASSERT_EQ(v.as_array()->cols, 4U);
  const Value* c = v.as_array()->cells;
  ASSERT_TRUE(c[0].is_number());
  EXPECT_EQ(c[0].as_number(), at13);
  ASSERT_TRUE(c[1].is_number());
  EXPECT_EQ(c[1].as_number(), at14);
  ASSERT_TRUE(c[2].is_error());
  EXPECT_EQ(c[2].as_error(), ErrorCode::Value);
  ASSERT_TRUE(c[3].is_error());
  EXPECT_EQ(c[3].as_error(), ErrorCode::Num);
}

TEST(ForecastEtsArrayTarget, ConfintLiftsOverTargets) {
  const double at13 = EvalExact("=FORECAST.ETS.CONFINT(13.0," + kDailyArgs + ",0.9,0)");
  const double at14 = EvalExact("=FORECAST.ETS.CONFINT(14.5," + kDailyArgs + ",0.9,0)");
  const Value v = EvalSource("=FORECAST.ETS.CONFINT({13.0;14.5;12.0}," + kDailyArgs + ",0.9,0)");
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array()->rows, 3U);
  ASSERT_EQ(v.as_array()->cols, 1U);
  const Value* c = v.as_array()->cells;
  ASSERT_TRUE(c[0].is_number());
  EXPECT_EQ(c[0].as_number(), at13);
  ASSERT_TRUE(c[1].is_number());
  EXPECT_EQ(c[1].as_number(), at14);
  ASSERT_TRUE(c[2].is_error());
  EXPECT_EQ(c[2].as_error(), ErrorCode::Num);
}

TEST(ForecastEtsArrayTarget, RangeTargetSpills) {
  Workbook wb = MakeLinearWorkbook();
  wb.sheet(0).set_cell_value(0, 2, Value::number(11.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(12.0));
  const Value s11 = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(s11.is_number());
  const double at11 = s11.as_number();
  const Value s12 = EvalSourceIn("=FORECAST.ETS(12, B1:B10, A1:A10, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(s12.is_number());
  const double at12 = s12.as_number();
  const Value v = EvalSourceIn("=FORECAST.ETS(C1:C2, B1:B10, A1:A10, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array()->rows, 2U);
  ASSERT_EQ(v.as_array()->cols, 1U);
  EXPECT_EQ(v.as_array()->cells[0].as_number(), at11);
  EXPECT_EQ(v.as_array()->cells[1].as_number(), at12);
}

TEST(ForecastEtsArrayTarget, TargetErrorWinsOverSeriesError) {
  // The series fails with #N/A (length mismatch); a target that is itself an
  // error or non-numeric reports its own error first, as the scalar call does.
  const std::string bad = "{1.0;2.0;3.0;4.0;5.0},{1.0;2.0;3.0;4.0}";
  EXPECT_EQ(EvalError("=FORECAST.ETS(6.0," + bad + ",0)"), ErrorCode::NA);
  EXPECT_EQ(EvalError("=FORECAST.ETS(\"x\"," + bad + ",0)"), ErrorCode::Value);
  const Value v = EvalSource("=FORECAST.ETS({6.0,\"x\"}," + bad + ",0)");
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array()->cols, 2U);
  EXPECT_EQ(v.as_array()->cells[0].as_error(), ErrorCode::NA);
  EXPECT_EQ(v.as_array()->cells[1].as_error(), ErrorCode::Value);
  const Value conf = EvalSource("=FORECAST.ETS.CONFINT({6.0,\"x\"}," + kDailyArgs + ",1.5,0)");
  ASSERT_TRUE(conf.is_array());
  EXPECT_EQ(conf.as_array()->cells[0].as_error(), ErrorCode::Num);
  EXPECT_EQ(conf.as_array()->cells[1].as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
