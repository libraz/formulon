//
// Unit tests for the visual conditional-format rule kinds: colour scales,
// data bars and icon sets.

#include "cf/scale_evaluator.h"

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_helpers.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "cf_evaluator_fixtures.h"
#include "gtest/gtest.h"
#include "value.h"

namespace formulon::cf {
namespace {

using test::At;
using test::CFEvalHarness;
using test::Cfvo;
using test::MakeRange;
using test::MakeRule;
using test::PopulateLinearPopulation;
using test::RGB;
using test::SqrefContext;
using test::TwoStopMinMax;

// ---------------------------------------------------------------------------
// ColorScale
// ---------------------------------------------------------------------------

ColorScaleSpec ThreeStopMinMidMax(Color lo, Color mid, Color hi, std::string mid_percent = "50") {
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Min), Cfvo(CfvoType::Percentile, std::move(mid_percent)), Cfvo(CfvoType::Max)};
  spec.colors = {lo, mid, hi};
  return spec;
}

TEST(CFEvaluator, ColorScaleWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ColorScaleTwoStopMinMaxInterpolatesBetweenStops) {
  // Population [10, 20, 30, 40, 50]; min = 10 (red), max = 50 (green).
  // Cell 30 is the midpoint → R=128, G=128, B=0 (linear RGB blend).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(30.0), ctx);
  EXPECT_EQ(match.kind, CFMatchKind::ColorScale);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  EXPECT_EQ(match.resolved_fill_color->r, 128);
  EXPECT_EQ(match.resolved_fill_color->g, 128);
  EXPECT_EQ(match.resolved_fill_color->b, 0);
}

TEST(CFEvaluator, ColorScaleClampsToOuterStops) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  // Below min: clamps to red.
  CFMatch lo = make_match(r, Value::number(-100.0), ctx);
  ASSERT_TRUE(lo.resolved_fill_color.has_value());
  EXPECT_EQ(*lo.resolved_fill_color, RGB(255, 0, 0));
  // Above max: clamps to green.
  CFMatch hi = make_match(r, Value::number(1000.0), ctx);
  ASSERT_TRUE(hi.resolved_fill_color.has_value());
  EXPECT_EQ(*hi.resolved_fill_color, RGB(0, 255, 0));
}

TEST(CFEvaluator, ColorScaleAtMinAndMaxReturnsExactStopColors) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch low = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(low.resolved_fill_color.has_value());
  EXPECT_EQ(*low.resolved_fill_color, RGB(255, 0, 0));
  CFMatch high = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(high.resolved_fill_color.has_value());
  EXPECT_EQ(*high.resolved_fill_color, RGB(0, 255, 0));
}

TEST(CFEvaluator, ColorScaleThreeStopUsesMiddleStopAtMedian) {
  // Population [10, 20, 30, 40, 50]; mid = percentile(50) = 30 (white).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = ThreeStopMinMidMax(RGB(255, 0, 0), RGB(255, 255, 255), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.resolved_fill_color.has_value());
  EXPECT_EQ(*mid.resolved_fill_color, RGB(255, 255, 255));
}

TEST(CFEvaluator, ColorScaleThreeStopInterpolatesWithinSegment) {
  // Population [10, 20, 30, 40, 50]; segments 10..30 (red→white) and
  // 30..50 (white→green). Cell 20 is the midpoint of the lower segment.
  // Lower segment blend: (255, 0, 0) → (255, 255, 255), fraction 0.5
  // → (255, 128, 128).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = ThreeStopMinMidMax(RGB(255, 0, 0), RGB(255, 255, 255), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(20.0), ctx);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  EXPECT_EQ(match.resolved_fill_color->r, 255);
  EXPECT_EQ(match.resolved_fill_color->g, 128);
  EXPECT_EQ(match.resolved_fill_color->b, 128);
}

TEST(CFEvaluator, ColorScaleNumberCfvoUsesLiteralThreshold) {
  // Force min=0, max=100 via Number CFVOs irrespective of population.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Number, "0"), Cfvo(CfvoType::Number, "100")};
  spec.colors = {RGB(0, 0, 0), RGB(255, 255, 255)};
  r.color_scale = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  // Halfway between black and white = (128, 128, 128).
  EXPECT_EQ(*match.resolved_fill_color, RGB(128, 128, 128));
}

TEST(CFEvaluator, ColorScalePercentCfvoUsesPopulationRangeFraction) {
  // Population [10, 20, 30, 40, 50]; min=10, max=50 → range=40.
  // Percent 25 → 10 + 0.25*40 = 20. Percent 75 → 10 + 0.75*40 = 40.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Percent, "25"), Cfvo(CfvoType::Percent, "75")};
  spec.colors = {RGB(0, 0, 0), RGB(255, 255, 255)};
  r.color_scale = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  // Cell at lower-stop position (20) → black.
  CFMatch lo = make_match(r, Value::number(20.0), ctx);
  ASSERT_TRUE(lo.resolved_fill_color.has_value());
  EXPECT_EQ(*lo.resolved_fill_color, RGB(0, 0, 0));
  // Cell at upper-stop position (40) → white.
  CFMatch hi = make_match(r, Value::number(40.0), ctx);
  ASSERT_TRUE(hi.resolved_fill_color.has_value());
  EXPECT_EQ(*hi.resolved_fill_color, RGB(255, 255, 255));
}

TEST(CFEvaluator, ColorScaleNonNumericCellDoesNotResolve) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::text("middle"), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
  CFMatch match = make_match(r, Value::text("middle"), ctx);
  EXPECT_FALSE(match.resolved_fill_color.has_value());
}

TEST(CFEvaluator, ColorScaleEmptyPopulationDoesNotResolve) {
  CFEvalHarness harness;  // No values populated; sqref is all-blank.
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::number(0.0), ctx));
  CFMatch match = make_match(r, Value::number(0.0), ctx);
  EXPECT_FALSE(match.resolved_fill_color.has_value());
}

TEST(CFEvaluator, ColorScaleDegeneratePopulationCollapsesToFirstStop) {
  // Population is [7, 7, 7, 7]; min == max. The cell value 7 hits the
  // first stop's clamp branch → returns colors[0].
  CFEvalHarness harness;
  for (std::uint32_t row = 0; row < 4; ++row) {
    harness.sheet.set_cell_value(row, 0, Value::number(7.0));
  }
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 3, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(7.0), ctx);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  EXPECT_EQ(*match.resolved_fill_color, RGB(255, 0, 0));
}

TEST(CFEvaluator, ColorScaleFormulaCfvoEvaluatesAtAnchor) {
  // CFVO formulas reference cells; value comes from the formula evaluator.
  // A1 = 0 (min anchor), B1 = 100 (max anchor); cell 50 → mid grey.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(0.0));
  harness.sheet.set_cell_value(0, 1, Value::number(100.0));
  // Population is the union: still 0..100 across [A1, B1].
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 0, 1)};
  CFRule r = MakeRule(RuleType::ColorScale);
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Formula, "A1"), Cfvo(CfvoType::Formula, "B1")};
  spec.colors = {RGB(0, 0, 0), RGB(255, 255, 255)};
  r.color_scale = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  EXPECT_EQ(*match.resolved_fill_color, RGB(128, 128, 128));
}

// A scale threshold belongs to the rule, not to the cell being rendered.
// Excel resolves a Formula CFVO once, at the authoring cell, so the whole
// sqref block shares one gradient; resolving it per target turned a
// relative-reference CFVO into a gradient that slid row by row. The
// population is already block-shared, so only the thresholds could drift.
TEST(CFEvaluator, ColorScaleFormulaCfvoResolvesAtAnchorForEveryTargetInTheBlock) {
  CFEvalHarness harness;
  // Threshold sources: only A1 / C1 are the rule's real endpoints. The
  // rows below them hold decoys that a per-target shift would pick up.
  harness.sheet.set_cell_value(0, 0, Value::number(0.0));
  harness.sheet.set_cell_value(1, 0, Value::number(1000.0));
  harness.sheet.set_cell_value(2, 0, Value::number(2000.0));
  harness.sheet.set_cell_value(0, 2, Value::number(100.0));
  harness.sheet.set_cell_value(1, 2, Value::number(3000.0));
  harness.sheet.set_cell_value(2, 2, Value::number(4000.0));
  // The scaled block itself: B1:B3, all holding the same value so any
  // difference in the rendered colour comes from the thresholds alone.
  harness.sheet.set_cell_value(0, 1, Value::number(50.0));
  harness.sheet.set_cell_value(1, 1, Value::number(50.0));
  harness.sheet.set_cell_value(2, 1, Value::number(50.0));

  const std::vector<CFCellRange> sqref{MakeRange(0, 1, 2, 1)};
  CFRule r = MakeRule(RuleType::ColorScale);
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Formula, "A1"), Cfvo(CfvoType::Formula, "C1")};
  spec.colors = {RGB(0, 0, 0), RGB(255, 255, 255)};
  r.color_scale = std::move(spec);

  // Thresholds resolve to [A1, C1] = [0, 100] for every row, so a value
  // of 50 renders mid grey throughout.
  for (std::uint32_t row = 0; row <= 2U; ++row) {
    CFEvalContext ctx = harness.context(At(0, 1), At(row, 1));
    ctx.sqref = &sqref;
    CFMatch match = make_match(r, Value::number(50.0), ctx);
    ASSERT_TRUE(match.resolved_fill_color.has_value()) << "row=" << row;
    EXPECT_EQ(*match.resolved_fill_color, RGB(128, 128, 128)) << "row=" << row;
  }
}

TEST(CFEvaluator, ColorScaleValueOnlyOverloadStillReturnsFalse) {
  // Pin the staging contract: ColorScale returns false on the
  // value-only overload (no sqref / population access). The value-only
  // make_match still returns a DifferentialFormat match.
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  EXPECT_FALSE(match_rule(r, Value::number(30.0)));

  CFMatch match = make_match(r);
  EXPECT_EQ(match.kind, CFMatchKind::DifferentialFormat);
  EXPECT_FALSE(match.resolved_fill_color.has_value());
}

// ---------------------------------------------------------------------------
// DataBar
// ---------------------------------------------------------------------------

TEST(CFEvaluator, DataBarConstantPopulationIsFullLength) {
  CFRule rule = MakeRule(RuleType::DataBar);
  DataBarSpec bar;
  bar.min.type = CfvoType::Min;
  bar.max.type = CfvoType::Max;
  bar.min_length_pct = 10;
  bar.max_length_pct = 90;
  rule.data_bar = bar;
  ColorScalePopulation population;
  population.sorted = {42.0, 42.0, 42.0};
  population.min = 42.0;
  population.max = 42.0;
  CFEvalHarness harness;
  std::vector<CFCellRange> sqref{{At(0, 0), At(2, 0)}};
  CFEvalContext context = harness.context(At(0, 0), At(0, 0));
  context.sqref = &sqref;
  context.cached_population = &population;

  auto render = scales::resolve_data_bar(rule, Value::number(42.0), context);
  ASSERT_TRUE(render.has_value());
  EXPECT_DOUBLE_EQ(render->length_pct, 90.0);
}

DataBarSpec MakeDataBarSpec(CfValueObject min_cfvo, CfValueObject max_cfvo, Color fill = RGB(0, 128, 255),
                            DataBarAxisPosition axis = DataBarAxisPosition::Automatic) {
  DataBarSpec spec;
  spec.min = std::move(min_cfvo);
  spec.max = std::move(max_cfvo);
  spec.fill = fill;
  spec.negative_fill = RGB(255, 0, 0);
  spec.axis_position = axis;
  spec.min_length_pct = 0;
  spec.max_length_pct = 100;
  return spec;
}

TEST(CFEvaluator, DataBarWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, DataBarLengthIsLinearBetweenMinAndMax) {
  // Population [10, 20, 30, 40, 50]; min=10, max=50 → range=40.
  // Cell at min → length 0%. Cell at max → length 100%. Cell at 30 → 50%.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch lo = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(lo.data_bar_render.has_value());
  EXPECT_EQ(lo.kind, CFMatchKind::DataBar);
  EXPECT_DOUBLE_EQ(lo.data_bar_render->length_pct, 0.0);

  CFMatch hi = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(hi.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(hi.data_bar_render->length_pct, 100.0);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(mid.data_bar_render->length_pct, 50.0);
}

TEST(CFEvaluator, DataBarClampsValuesOutsideThresholdRange) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch below = make_match(r, Value::number(-100.0), ctx);
  ASSERT_TRUE(below.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(below.data_bar_render->length_pct, 0.0);

  CFMatch above = make_match(r, Value::number(1000.0), ctx);
  ASSERT_TRUE(above.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(above.data_bar_render->length_pct, 100.0);
}

TEST(CFEvaluator, DataBarMinAndMaxLengthAreApplied) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  DataBarSpec spec = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  spec.min_length_pct = 10;
  spec.max_length_pct = 90;
  r.data_bar = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch lo = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(lo.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(lo.data_bar_render->length_pct, 10.0);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(mid.data_bar_render->length_pct, 50.0);

  CFMatch hi = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(hi.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(hi.data_bar_render->length_pct, 90.0);
}

TEST(CFEvaluator, DataBarAutomaticAxisAtZeroForAllNonNegativePopulation) {
  // Population [10, 20, 30, 40, 50] is all >= 0 → axis at left edge.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 0.0);
  EXPECT_FALSE(match.data_bar_render->is_negative);
}

TEST(CFEvaluator, DataBarAutomaticAxisAtHundredForAllNegativePopulation) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(-50.0));
  harness.sheet.set_cell_value(1, 0, Value::number(-30.0));
  harness.sheet.set_cell_value(2, 0, Value::number(-10.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(-30.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 100.0);
  EXPECT_TRUE(match.data_bar_render->is_negative);
}

TEST(CFEvaluator, DataBarAutomaticAxisProportionalForMixedSignPopulation) {
  // Population [-30, 10, 70] → min = -30, max = 70. Negative span = 30,
  // total span = 100 → axis at 30%.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(-30.0));
  harness.sheet.set_cell_value(1, 0, Value::number(10.0));
  harness.sheet.set_cell_value(2, 0, Value::number(70.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 30.0);
}

TEST(CFEvaluator, DataBarMixedSignLengthSplitsAtAxisInsteadOfWholeRangeLinear) {
  // Population [-50, -10, 20, 100] → min = -50, max = 100. Axis at
  // |min| / (max - min) = 50 / 150 ≈ 33.33%.
  //
  // Regression: length used to be the whole-range linear map
  // `(cell - min) / (max - min)`, which for e.g. cell=-10 would give
  // `(-10 - -50) / 150 = 26.67%` -- a bar nearly as long as the true
  // min. The correct OOXML semantics split at the axis: positive bars
  // scale by `value / max`, negative bars by `value / min` (both
  // negative, so a positive fraction) -- cell=-10 should be a *short*
  // bar (20% of the negative side), not a long one.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(-50.0));
  harness.sheet.set_cell_value(1, 0, Value::number(-10.0));
  harness.sheet.set_cell_value(2, 0, Value::number(20.0));
  harness.sheet.set_cell_value(3, 0, Value::number(100.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 3, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch axis_probe = make_match(r, Value::number(20.0), ctx);
  ASSERT_TRUE(axis_probe.data_bar_render.has_value());
  EXPECT_NEAR(axis_probe.data_bar_render->axis_position_pct, 33.333333333333336, 1e-9);

  // Most-negative value: full-length bar on the negative side.
  CFMatch most_negative = make_match(r, Value::number(-50.0), ctx);
  ASSERT_TRUE(most_negative.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(most_negative.data_bar_render->length_pct, 100.0);
  EXPECT_TRUE(most_negative.data_bar_render->is_negative);

  // -10 is 20% of the way from 0 to min (-50): a short negative bar.
  CFMatch small_negative = make_match(r, Value::number(-10.0), ctx);
  ASSERT_TRUE(small_negative.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(small_negative.data_bar_render->length_pct, 20.0);
  EXPECT_TRUE(small_negative.data_bar_render->is_negative);

  // 20 is 20% of the way from 0 to max (100): a short positive bar.
  CFMatch small_positive = make_match(r, Value::number(20.0), ctx);
  ASSERT_TRUE(small_positive.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(small_positive.data_bar_render->length_pct, 20.0);
  EXPECT_FALSE(small_positive.data_bar_render->is_negative);

  // Most-positive value (= max): full-length bar on the positive side.
  CFMatch most_positive = make_match(r, Value::number(100.0), ctx);
  ASSERT_TRUE(most_positive.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(most_positive.data_bar_render->length_pct, 100.0);
  EXPECT_FALSE(most_positive.data_bar_render->is_negative);
}

TEST(CFEvaluator, DataBarAllPositiveDataKeepsWholeRangeLinearLength) {
  // Non-regression: same-sign data must keep the original whole-range
  // linear map even though it also uses Automatic axis mode.
  // Population [10, 20, 30, 40, 50] -- min=10, max=50, mirrors
  // `DataBarLengthIsLinearBetweenMinAndMax` but pinned to the mixed-sign
  // code path's guard condition (all non-negative here).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(mid.data_bar_render->length_pct, 50.0);
}

TEST(CFEvaluator, DataBarMiddleAxisIsAlwaysFifty) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max), RGB(0, 0, 255), DataBarAxisPosition::Middle);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 50.0);
}

TEST(CFEvaluator, DataBarNoneAxisIsZero) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max), RGB(0, 0, 255), DataBarAxisPosition::None);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 0.0);
}

TEST(CFEvaluator, DataBarSelectsNegativeFillForNegativeValues) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(-30.0));
  harness.sheet.set_cell_value(1, 0, Value::number(70.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 1, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  DataBarSpec spec = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  spec.fill = RGB(0, 200, 0);
  spec.negative_fill = RGB(200, 0, 0);
  r.data_bar = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch positive = make_match(r, Value::number(70.0), ctx);
  ASSERT_TRUE(positive.data_bar_render.has_value());
  EXPECT_EQ(positive.data_bar_render->fill, RGB(0, 200, 0));
  EXPECT_FALSE(positive.data_bar_render->is_negative);

  CFMatch negative = make_match(r, Value::number(-30.0), ctx);
  ASSERT_TRUE(negative.data_bar_render.has_value());
  EXPECT_EQ(negative.data_bar_render->fill, RGB(200, 0, 0));
  EXPECT_TRUE(negative.data_bar_render->is_negative);
}

TEST(CFEvaluator, DataBarNumberCfvosUseLiteralThresholds) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  // Force min=0, max=100 regardless of population.
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Number, "0"), Cfvo(CfvoType::Number, "100"));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->length_pct, 50.0);
}

TEST(CFEvaluator, DataBarNonNumericCellDoesNotResolve) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::text("30"), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
  CFMatch match = make_match(r, Value::text("30"), ctx);
  EXPECT_FALSE(match.data_bar_render.has_value());
}

TEST(CFEvaluator, DataBarDegenerateRangeUsesFullLength) {
  // Population is all 7s → Excel draws a full-length bar for every cell.
  CFEvalHarness harness;
  for (std::uint32_t row = 0; row < 4; ++row) {
    harness.sheet.set_cell_value(row, 0, Value::number(7.0));
  }
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 3, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_TRUE(match_rule(r, Value::number(7.0), ctx));
  CFMatch match = make_match(r, Value::number(7.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->length_pct, 100.0);
}

TEST(CFEvaluator, DataBarValueOnlyOverloadStillReturnsFalse) {
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  EXPECT_FALSE(match_rule(r, Value::number(30.0)));
  CFMatch match = make_match(r);
  EXPECT_EQ(match.kind, CFMatchKind::DifferentialFormat);
  EXPECT_FALSE(match.data_bar_render.has_value());
}

// ---------------------------------------------------------------------------
// IconSet
// ---------------------------------------------------------------------------

CfValueObject IconCfvo(CfvoType type, std::string value, bool gte = true) {
  CfValueObject cfvo = Cfvo(type, std::move(value));
  cfvo.gte = gte;
  return cfvo;
}

IconSetSpec ThreeIconNumberSet(std::string lo, std::string hi, IconSetName name = IconSetName::Three_TrafficLights1) {
  IconSetSpec spec;
  spec.name = name;
  spec.thresholds = {IconCfvo(CfvoType::Number, std::move(lo)), IconCfvo(CfvoType::Number, std::move(hi))};
  return spec;
}

TEST(CFEvaluator, IconSetWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  EXPECT_FALSE(match_rule(r, Value::number(30.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, IconSetThreeIconAssignsBucketByThreshold) {
  // Thresholds 20 / 40 → bucket 0: cell < 20; bucket 1: 20 <= cell < 40;
  // bucket 2: cell >= 40.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch low = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(low.icon_render.has_value());
  EXPECT_EQ(low.kind, CFMatchKind::IconSet);
  EXPECT_EQ(low.icon_render->icon_index, 0);
  EXPECT_EQ(low.icon_render->set_name, IconSetName::Three_TrafficLights1);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.icon_render.has_value());
  EXPECT_EQ(mid.icon_render->icon_index, 1);

  CFMatch high = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(high.icon_render.has_value());
  EXPECT_EQ(high.icon_render->icon_index, 2);
}

TEST(CFEvaluator, IconSetExactlyAtThresholdHonoursGteFlag) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  // gte=true on threshold 20: cell == 20 belongs to the upper bucket.
  IconSetSpec spec_gte;
  spec_gte.name = IconSetName::Three_TrafficLights1;
  spec_gte.thresholds = {IconCfvo(CfvoType::Number, "20", true), IconCfvo(CfvoType::Number, "40", true)};
  r.icon_set = spec_gte;
  const auto ctx_gte = SqrefContext(harness, sqref);
  CFMatch at_threshold_gte = make_match(r, Value::number(20.0), ctx_gte);
  ASSERT_TRUE(at_threshold_gte.icon_render.has_value());
  EXPECT_EQ(at_threshold_gte.icon_render->icon_index, 1);

  // gte=false on threshold 20: cell == 20 belongs to the lower bucket.
  IconSetSpec spec_gt;
  spec_gt.name = IconSetName::Three_TrafficLights1;
  spec_gt.thresholds = {IconCfvo(CfvoType::Number, "20", false), IconCfvo(CfvoType::Number, "40", false)};
  r.icon_set = spec_gt;
  const auto ctx_gt = SqrefContext(harness, sqref);
  CFMatch at_threshold_gt = make_match(r, Value::number(20.0), ctx_gt);
  ASSERT_TRUE(at_threshold_gt.icon_render.has_value());
  EXPECT_EQ(at_threshold_gt.icon_render->icon_index, 0);
}

TEST(CFEvaluator, IconSetReverseFlipsBucketIndex) {
  // Without reverse: 10 → 0, 50 → 2. With reverse: 10 → 2, 50 → 0.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  IconSetSpec spec = ThreeIconNumberSet("20", "40");
  spec.reverse = true;
  r.icon_set = spec;
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch low = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(low.icon_render.has_value());
  EXPECT_EQ(low.icon_render->icon_index, 2);

  CFMatch high = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(high.icon_render.has_value());
  EXPECT_EQ(high.icon_render->icon_index, 0);
}

TEST(CFEvaluator, IconSetFiveIconAssignsAcrossFourThresholds) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  IconSetSpec spec;
  spec.name = IconSetName::Five_Arrows;
  spec.thresholds = {IconCfvo(CfvoType::Number, "15"), IconCfvo(CfvoType::Number, "25"),
                     IconCfvo(CfvoType::Number, "35"), IconCfvo(CfvoType::Number, "45")};
  r.icon_set = spec;
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_EQ(make_match(r, Value::number(10.0), ctx).icon_render->icon_index, 0);
  EXPECT_EQ(make_match(r, Value::number(20.0), ctx).icon_render->icon_index, 1);
  EXPECT_EQ(make_match(r, Value::number(30.0), ctx).icon_render->icon_index, 2);
  EXPECT_EQ(make_match(r, Value::number(40.0), ctx).icon_render->icon_index, 3);
  EXPECT_EQ(make_match(r, Value::number(50.0), ctx).icon_render->icon_index, 4);
}

TEST(CFEvaluator, IconSetPercentCfvoUsesPopulationRangeFraction) {
  // Population [10, 20, 30, 40, 50]; min=10, max=50 → range=40.
  // Thresholds at 25% (=20) and 75% (=40).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  IconSetSpec spec;
  spec.name = IconSetName::Three_TrafficLights1;
  spec.thresholds = {IconCfvo(CfvoType::Percent, "25"), IconCfvo(CfvoType::Percent, "75")};
  r.icon_set = spec;
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_EQ(make_match(r, Value::number(10.0), ctx).icon_render->icon_index, 0);
  EXPECT_EQ(make_match(r, Value::number(30.0), ctx).icon_render->icon_index, 1);
  EXPECT_EQ(make_match(r, Value::number(50.0), ctx).icon_render->icon_index, 2);
}

TEST(CFEvaluator, IconSetNonNumericCellDoesNotResolve) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::text("30"), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
  CFMatch match = make_match(r, Value::text("30"), ctx);
  EXPECT_FALSE(match.icon_render.has_value());
}

TEST(CFEvaluator, IconSetEmptyPopulationDoesNotResolve) {
  CFEvalHarness harness;
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  CFMatch match = make_match(r, Value::number(30.0), ctx);
  EXPECT_FALSE(match.icon_render.has_value());
}

TEST(CFEvaluator, IconSetEmptyThresholdsDoesNotResolve) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  IconSetSpec spec;
  spec.name = IconSetName::Three_TrafficLights1;
  // No thresholds.
  r.icon_set = spec;
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
}

TEST(CFEvaluator, IconSetValueOnlyOverloadStillReturnsFalse) {
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  EXPECT_FALSE(match_rule(r, Value::number(30.0)));
  CFMatch match = make_match(r);
  EXPECT_EQ(match.kind, CFMatchKind::DifferentialFormat);
  EXPECT_FALSE(match.icon_render.has_value());
}

}  // namespace
}  // namespace formulon::cf
