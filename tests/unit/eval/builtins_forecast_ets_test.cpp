// Copyright 2026 libraz. Licensed under the MIT License.
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
// SCOPE: this file asserts STRUCTURAL behaviour only -- error codes,
// monotonicity (larger confidence -> larger half-width, larger horizon
// -> larger half-width via sqrt(h) growth), finiteness of stats, and
// that detection / resampling pipelines fire at all. The exact numeric
// outputs are NOT asserted here; oracle calibration against macOS
// Excel 365 will arrive in a follow-up commit. Tolerances on
// `EXPECT_NEAR` are deliberately generous.

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
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Test harness
// ---------------------------------------------------------------------------

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

Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

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

TEST(ForecastEts, TextInTimelineIsValueError) {
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
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
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

TEST(ForecastEtsStat, GammaZeroForNonSeasonal) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS.STAT(B1:B10, A1:A10, 3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
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

TEST(ForecastEtsCrossCut, AggregationOutOfDomainHighIsValueError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0, 1, 8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ForecastEtsCrossCut, AggregationOutOfDomainLowIsValueError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0, 1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
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

TEST(ForecastEtsCrossCut, DataCompletionTwoIsValueError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ForecastEtsCrossCut, DataCompletionNegativeIsValueError) {
  const Workbook wb = MakeLinearWorkbook();
  const Value v = EvalSourceIn("=FORECAST.ETS(11, B1:B10, A1:A10, 0, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
