// Copyright 2026 libraz. Licensed under the MIT License.
//
// Match-result types for the conditional-format evaluator. The
// evaluator returns one `CFMatch` per rule that fires for a given
// cell; visual rule kinds (`ColorScale`, `DataBar`, `IconSet`)
// additionally carry the resolved render payload. The UI does not
// re-implement CF semantics — it consumes `CFMatch` and either
// looks up the engine-supplied dxf id or paints the resolved
// colour / data bar / icon.
//
// See backup/plans/20-conditional-format-deep.md §20.5.

#ifndef FORMULON_CF_CF_MATCH_H_
#define FORMULON_CF_CF_MATCH_H_

#include <cstdint>
#include <optional>
#include <string>

#include "cf/cf_types.h"

namespace formulon::cf {

/// Distinguishes how a UI should consume the match. Mutually exclusive:
/// each `CFMatch` carries exactly one engaged payload optional.
enum class CFMatchKind : std::uint8_t {
  /// `dxf_id` references `styles.dxfs[i]`; the UI merges that
  /// differential-format record into the cell's base style.
  DifferentialFormat = 0,
  /// `resolved_fill_color` is the interpolated cell-fill colour for a
  /// 2- or 3-stop colour scale.
  ColorScale = 1,
  /// `data_bar_render` carries the bar length and axis position.
  DataBar = 2,
  /// `icon_render` carries the bucket index (after `reverse` flip).
  IconSet = 3,
};

/// Resolved data-bar render payload.
struct DataBarRender {
  /// 0..100 — length of the bar relative to the cell width.
  double length_pct = 0.0;
  /// 0 (left edge) or 50 (centred axis for mixed-sign sets).
  double axis_position_pct = 0.0;
  /// True when the cell value is negative; the UI flips fill side.
  bool is_negative = false;
  Color fill{};
  std::optional<Color> border;
  bool gradient = true;
};

/// Resolved icon-set render payload.
struct IconRender {
  IconSetName set_name = IconSetName::Three_Arrows;
  /// 0..N-1, after `reverse` has been applied.
  std::uint8_t icon_index = 0;
};

/// One rule's match result for one cell. The engaged payload optional
/// matches `kind`; the others are empty.
struct CFMatch {
  /// Lifted from `CFRule::id`; empty when the rule has no extLst id.
  std::string rule_id;
  /// Workbook-global priority, lifted from `CFRule::priority`. Smaller
  /// numbers ranked higher; the evaluator returns the match list in
  /// priority-ascending order so a UI can fold them in order.
  std::int32_t priority = 0;
  CFMatchKind kind = CFMatchKind::DifferentialFormat;

  /// Engaged for `DifferentialFormat` matches. References
  /// `styles.dxfs[i]`.
  std::optional<std::uint32_t> dxf_id;
  /// Engaged for `ColorScale` matches.
  std::optional<Color> resolved_fill_color;
  /// Engaged for `DataBar` matches.
  std::optional<DataBarRender> data_bar_render;
  /// Engaged for `IconSet` matches.
  std::optional<IconRender> icon_render;
};

}  // namespace formulon::cf

#endif  // FORMULON_CF_CF_MATCH_H_
