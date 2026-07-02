// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// ColorScale / DataBar / IconSet resolution. See scale_evaluator.h for
// the contract and cf/cf_evaluator.h for the public-facing rule kinds.

#include "cf/scale_evaluator.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_helpers.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon::cf::scales {

namespace {

constexpr double kPercentDivisor = 100.0;
constexpr double kColorChannelMax = 255.0;

// ---------------------------------------------------------------------------
// CFVO resolution — shared by ColorScale / DataBar / IconSet.
// ---------------------------------------------------------------------------

// Resolves a single `<cfvo>` to its threshold value. `Formula` CFVOs
// run through the same parse-shift-evaluate path the rest of the
// context-aware evaluator uses, anchored at the rule's anchor (the
// formula authoring cell). Returns nullopt when the CFVO cannot be
// resolved (e.g. malformed literal, formula evaluation error).
std::optional<double> resolve_cfvo(const CfValueObject& cfvo, const ColorScalePopulation& pop,
                                   const CFEvalContext& ctx) {
  switch (cfvo.type) {
    case CfvoType::Number:
      return helpers::parse_double(cfvo.value);
    case CfvoType::Percent: {
      auto pct = helpers::parse_double(cfvo.value);
      if (!pct.has_value()) {
        return std::nullopt;
      }
      return pop.min + (*pct / kPercentDivisor) * (pop.max - pop.min);
    }
    case CfvoType::Percentile: {
      auto pct = helpers::parse_double(cfvo.value);
      if (!pct.has_value() || pop.sorted.empty()) {
        return std::nullopt;
      }
      return helpers::percentile_inc(pop.sorted, *pct / kPercentDivisor);
    }
    case CfvoType::Min:
    case CfvoType::AutoMin:
      return pop.sorted.empty() ? std::optional<double>() : std::optional<double>(pop.min);
    case CfvoType::Max:
    case CfvoType::AutoMax:
      return pop.sorted.empty() ? std::optional<double>() : std::optional<double>(pop.max);
    case CfvoType::Formula: {
      const Value evaluated = helpers::parse_shift_evaluate(cfvo.value, ctx);
      if (evaluated.is_number()) {
        return evaluated.as_number();
      }
      if (evaluated.is_boolean()) {
        return evaluated.as_boolean() ? 1.0 : 0.0;
      }
      return std::nullopt;
    }
  }
  return std::nullopt;
}

std::optional<std::vector<double>> resolve_cfvo_list(const std::vector<CfValueObject>& cfvos,
                                                     const ColorScalePopulation& pop, const CFEvalContext& ctx) {
  std::vector<double> resolved;
  resolved.reserve(cfvos.size());
  for (const CfValueObject& cfvo : cfvos) {
    auto value = resolve_cfvo(cfvo, pop, ctx);
    if (!value.has_value()) {
      return std::nullopt;
    }
    resolved.push_back(*value);
  }
  return resolved;
}

// ---------------------------------------------------------------------------
// ColorScale — resolve `<cfvo>` thresholds against the sqref population
// and linearly interpolate the bounding stop colours in RGB space.
// ---------------------------------------------------------------------------

// Linear interpolation between two sRGB colours. `fraction` is clamped
// to [0, 1] by the caller. Alpha is interpolated alongside RGB so
// stops with transparent components blend correctly.
Color interpolate_color(Color start, Color end, double fraction) {
  const auto blend = [fraction](std::uint8_t low_channel, std::uint8_t high_channel) {
    const double mixed = static_cast<double>(low_channel) +
                         fraction * (static_cast<double>(high_channel) - static_cast<double>(low_channel));
    const double clamped = std::max(0.0, std::min(kColorChannelMax, mixed));
    return static_cast<std::uint8_t>(std::lround(clamped));
  };
  Color out;
  out.r = blend(start.r, end.r);
  out.g = blend(start.g, end.g);
  out.b = blend(start.b, end.b);
  out.a = blend(start.a, end.a);
  return out;
}

// ---------------------------------------------------------------------------
// DataBar axis placement.
//
// `length_pct` is the bar length expressed as a 0..100 percent of the
// cell width: cells at `min_threshold` produce `min_length_pct`; cells
// at `max_threshold` produce `max_length_pct`; cells outside clamp.
// `axis_position_pct` follows OOXML semantics: `Automatic` splits at the
// proportional negative offset, `Middle` pins to 50, `None` pins to 0.
// `is_negative` is set when the cell value is strictly negative so the
// host can flip the fill side.
// ---------------------------------------------------------------------------

constexpr double kAxisMid = 50.0;
constexpr double kAxisLeft = 0.0;
constexpr double kAxisRight = 100.0;

double automatic_axis_position(double threshold_min, double threshold_max) {
  // All non-negative → bar grows from the left edge.
  if (threshold_min >= 0.0) {
    return kAxisLeft;
  }
  // All non-positive → bar grows from the right edge.
  if (threshold_max <= 0.0) {
    return kAxisRight;
  }
  // Mixed sign: split proportionally so equal-magnitude positive and
  // negative bars meet at the same axis. Rare degenerate case
  // (threshold_min == 0 == threshold_max) is handled by the branches
  // above.
  const double negative_span = -threshold_min;
  const double total_span = negative_span + threshold_max;
  return (negative_span / total_span) * kAxisRight;
}

}  // namespace

// ---------------------------------------------------------------------------
// Public resolvers.
// ---------------------------------------------------------------------------

std::optional<Color> resolve_color_scale(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.color_scale.has_value()) {
    return std::nullopt;
  }
  const ColorScaleSpec& spec = *rule.color_scale;
  if (spec.thresholds.size() != spec.colors.size() || spec.thresholds.size() < 2) {
    return std::nullopt;
  }

  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = helpers::numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return std::nullopt;
  }

  auto resolved_thresholds = resolve_cfvo_list(spec.thresholds, *pop, ctx);
  if (!resolved_thresholds.has_value()) {
    return std::nullopt;
  }

  const double cell = cell_value.as_number();
  // Locate the segment that contains the cell value. Cells outside the
  // outermost stops clamp to the boundary colour.
  if (cell <= resolved_thresholds->front()) {
    return spec.colors.front();
  }
  if (cell >= resolved_thresholds->back()) {
    return spec.colors.back();
  }
  for (std::size_t i = 0; i + 1 < resolved_thresholds->size(); ++i) {
    const double lower_bound = (*resolved_thresholds)[i];
    const double upper_bound = (*resolved_thresholds)[i + 1];
    if (cell >= lower_bound && cell <= upper_bound) {
      // When the segment collapses (lower == upper), pick the upper
      // colour; the cell is exactly at a stop so either end is correct.
      const double span = upper_bound - lower_bound;
      const double fraction = span == 0.0 ? 1.0 : (cell - lower_bound) / span;
      return interpolate_color(spec.colors[i], spec.colors[i + 1], fraction);
    }
  }
  return spec.colors.back();
}

bool match_color_scale(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  // ColorScale "matches" any cell for which a fill colour can be
  // computed. Non-numeric cells, empty populations, and malformed
  // specs short-circuit upstream and return nullopt.
  return resolve_color_scale(rule, cell_value, ctx).has_value();
}

std::optional<DataBarRender> resolve_data_bar(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.data_bar.has_value()) {
    return std::nullopt;
  }
  const DataBarSpec& spec = *rule.data_bar;

  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = helpers::numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return std::nullopt;
  }

  auto threshold_min = resolve_cfvo(spec.min, *pop, ctx);
  auto threshold_max = resolve_cfvo(spec.max, *pop, ctx);
  if (!threshold_min.has_value() || !threshold_max.has_value()) {
    return std::nullopt;
  }
  // Degenerate threshold range collapses the bar — no meaningful length
  // can be produced. Fall through to nullopt so the caller can decide.
  if (*threshold_min == *threshold_max) {
    return std::nullopt;
  }

  const double cell = cell_value.as_number();
  const auto min_len = static_cast<double>(spec.min_length_pct);
  const auto max_len = static_cast<double>(spec.max_length_pct);

  // OOXML's "automatic axis" semantics split mixed-sign data (min < 0 <
  // max) at the axis instead of interpolating linearly across the whole
  // [min, max] span: a positive value's bar grows rightward from the
  // axis proportional to `value / max`, a negative value's bar grows
  // leftward proportional to `value / min` (both negative, so this is a
  // positive fraction). Using the plain whole-range linear map here
  // would draw negative bars as if they were mirrored positive ones
  // instead of shrinking toward the axis as their magnitude drops.
  //
  // Same-sign data (min >= 0 or max <= 0) keeps the original whole-range
  // linear map: `automatic_axis_position` already pins the axis to an
  // edge in that case, so the two formulas would only coincide when the
  // pinned-edge threshold is exactly zero, and same-sign data has no
  // "other side" for the split formula to describe anyway.
  const bool mixed_sign =
      spec.axis_position == DataBarAxisPosition::Automatic && *threshold_min < 0.0 && *threshold_max > 0.0;
  double clamped_fraction;
  if (mixed_sign) {
    const double raw_fraction = (cell >= 0.0) ? cell / *threshold_max : cell / *threshold_min;
    clamped_fraction = std::max(0.0, std::min(1.0, raw_fraction));
  } else {
    const double range = *threshold_max - *threshold_min;
    const double raw_fraction = (cell - *threshold_min) / range;
    clamped_fraction = std::max(0.0, std::min(1.0, raw_fraction));
  }

  DataBarRender render;
  render.length_pct = min_len + clamped_fraction * (max_len - min_len);
  render.is_negative = cell < 0.0;
  render.fill = render.is_negative ? spec.negative_fill : spec.fill;
  render.border = render.is_negative ? spec.negative_border : spec.border;
  render.gradient = spec.gradient;

  switch (spec.axis_position) {
    case DataBarAxisPosition::None:
      render.axis_position_pct = kAxisLeft;
      break;
    case DataBarAxisPosition::Middle:
      render.axis_position_pct = kAxisMid;
      break;
    case DataBarAxisPosition::Automatic:
      render.axis_position_pct = automatic_axis_position(*threshold_min, *threshold_max);
      break;
  }
  return render;
}

bool match_data_bar(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  return resolve_data_bar(rule, cell_value, ctx).has_value();
}

// ---------------------------------------------------------------------------
// IconSet — bucket the cell value across `N - 1` thresholds and assign
// an icon index in `[0, N - 1]` for an N-icon set.
//
// Each threshold's `gte` flag toggles `>=` vs. `>` at that boundary. The
// loop walks every threshold and bumps the index for each one the cell
// passes — assuming the OOXML reader populated thresholds in ascending
// order, which Excel always emits. `reverse` flips the index so the
// default-up direction can be inverted without re-sorting the colour /
// icon resources.
// ---------------------------------------------------------------------------

std::optional<IconRender> resolve_icon_set(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.icon_set.has_value()) {
    return std::nullopt;
  }
  const IconSetSpec& spec = *rule.icon_set;
  if (spec.thresholds.empty()) {
    return std::nullopt;  // A 1-icon "set" has no boundaries to match against.
  }

  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = helpers::numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return std::nullopt;
  }

  auto resolved = resolve_cfvo_list(spec.thresholds, *pop, ctx);
  if (!resolved.has_value()) {
    return std::nullopt;
  }

  const double cell = cell_value.as_number();
  std::uint8_t icon_index = 0;
  for (std::size_t i = 0; i < resolved->size(); ++i) {
    const bool above = spec.thresholds[i].gte ? (cell >= (*resolved)[i]) : (cell > (*resolved)[i]);
    if (above) {
      icon_index = static_cast<std::uint8_t>(i + 1);
    }
  }

  if (spec.reverse) {
    const auto bucket_count = static_cast<std::uint8_t>(resolved->size() + 1);
    icon_index = static_cast<std::uint8_t>(bucket_count - 1 - icon_index);
  }

  IconRender render;
  render.set_name = spec.name;
  render.icon_index = icon_index;
  return render;
}

bool match_icon_set(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  return resolve_icon_set(rule, cell_value, ctx).has_value();
}

}  // namespace formulon::cf::scales
