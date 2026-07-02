// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Conditional-format data-model primitives shared by the workbook layer,
// the OOXML reader/writer, and the CF evaluator.
//
// This header is intentionally type-only: no behaviour, no allocation
// strategy, no evaluator hooks. Subsequent PRs build the evaluator,
// reader/writer, and sheet integration on top of these structures.

#ifndef FORMULON_CF_CF_TYPES_H_
#define FORMULON_CF_CF_TYPES_H_

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "cell.h"

namespace formulon::cf {

/// One of the seventeen `<cfRule>` types defined by ECMA-376
/// (`type` attribute on `<cfRule>`).
enum class RuleType : std::uint8_t {
  Expression = 0,
  CellIs = 1,
  ColorScale = 2,
  DataBar = 3,
  IconSet = 4,
  Top10 = 5,
  AboveAverage = 6,
  ContainsText = 7,
  NotContainsText = 8,
  BeginsWith = 9,
  EndsWith = 10,
  ContainsBlanks = 11,
  NotContainsBlanks = 12,
  ContainsErrors = 13,
  NotContainsErrors = 14,
  TimePeriod = 15,
  DuplicateValues = 16,
  UniqueValues = 17,
};

/// Comparison operator carried by `cellIs` rules.
enum class CellIsOperator : std::uint8_t {
  LessThan = 0,
  LessThanOrEqual = 1,
  Equal = 2,
  NotEqual = 3,
  GreaterThanOrEqual = 4,
  GreaterThan = 5,
  Between = 6,
  NotBetween = 7,
};

/// `<cfvo>` (conditional format value object) `type` attribute.
enum class CfvoType : std::uint8_t {
  Number = 0,
  Percent = 1,
  Percentile = 2,
  Min = 3,
  Max = 4,
  Formula = 5,
  AutoMin = 6,
  AutoMax = 7,
};

/// `<dataBar>` axis-position attribute (Excel 2010+ extension).
enum class DataBarAxisPosition : std::uint8_t {
  Automatic = 0,
  Middle = 1,
  None = 2,
};

/// Built-in icon set. Names follow ECMA-376 / OOXML
/// (`iconSet` element `iconSet` attribute).
enum class IconSetName : std::uint8_t {
  Three_Arrows = 0,
  Three_ArrowsGray = 1,
  Three_Flags = 2,
  Three_TrafficLights1 = 3,
  Three_TrafficLights2 = 4,
  Three_Signs = 5,
  Three_Symbols = 6,
  Three_Symbols2 = 7,
  Four_Arrows = 8,
  Four_ArrowsGray = 9,
  Four_RedToBlack = 10,
  Four_Rating = 11,
  Four_TrafficLights = 12,
  Five_Arrows = 13,
  Five_ArrowsGray = 14,
  Five_Rating = 15,
  Five_Quarters = 16,
};

/// `<cfRule type="timePeriod">` value bucket.
enum class TimePeriod : std::uint8_t {
  Today = 0,
  Yesterday = 1,
  Tomorrow = 2,
  Last7Days = 3,
  ThisWeek = 4,
  LastWeek = 5,
  NextWeek = 6,
  ThisMonth = 7,
  LastMonth = 8,
  NextMonth = 9,
};

/// Plain RGBA colour. Channels are 0-255 (sRGB). Alpha is opaque (255)
/// unless explicitly set; OOXML `tint` adjustments are resolved by the
/// reader before this struct is populated.
struct Color {
  std::uint8_t r = 0;
  std::uint8_t g = 0;
  std::uint8_t b = 0;
  std::uint8_t a = 255;

  friend bool operator==(Color x, Color y) noexcept { return x.r == y.r && x.g == y.g && x.b == y.b && x.a == y.a; }
  friend bool operator!=(Color x, Color y) noexcept { return !(x == y); }
};

/// Excel grid extent, mirrored from `Sheet::kMaxRows` / `Sheet::kMaxCols`
/// so `CFCellRange` can classify whole-column / whole-row ranges without a
/// heavy `sheet.h` include here. A `static_assert` in `cf_helpers.cpp`
/// keeps these in sync with the canonical `Sheet` constants.
inline constexpr std::uint32_t kCfMaxRows = 1048576U;
inline constexpr std::uint32_t kCfMaxCols = 16384U;

/// Inclusive cell-range expressed as two `CellAddress` corners. `first`
/// is the top-left, `last` is the bottom-right; single-cell ranges have
/// `first == last`. `sqref` lists are `std::vector<CFCellRange>`; the
/// OOXML reader splits whitespace-separated A1 tokens into one entry
/// each.
///
/// Whole-column (`A:A`, `A:C`) and whole-row (`1:1`, `1:3`) sqref tokens are
/// stored at their full logical extent — a whole column covers every row
/// (0..kCfMaxRows-1), a whole row every column (0..kCfMaxCols-1) — so
/// membership tests work unchanged. `is_full_col()` / `is_full_row()`
/// recover the classification from that extent (matching Excel, which
/// treats `A:A` and `A1:A1048576` as identical); range-aware evaluation
/// clamps the unbounded axis to the sheet's used range, and the writer
/// re-emits the compact `A:A` / `1:1` form. This keeps the struct layout
/// (and its C-ABI mirror `fm_cf_cell_range_t`) unchanged.
struct CFCellRange {
  CellAddress first{};
  CellAddress last{};

  /// True when the range spans every row of its column span (`A:A`, `A:C`).
  bool is_full_col() const noexcept { return first.row == 0 && last.row == kCfMaxRows - 1U; }
  /// True when the range spans every column of its row span (`1:1`, `1:3`).
  bool is_full_row() const noexcept { return first.col == 0 && last.col == kCfMaxCols - 1U; }

  friend bool operator==(CFCellRange a, CFCellRange b) noexcept { return a.first == b.first && a.last == b.last; }
  friend bool operator!=(CFCellRange a, CFCellRange b) noexcept { return !(a == b); }
};

/// One `<cfvo>` entry — a threshold either as a literal number / percent
/// / percentile, or as a formula evaluated with the rule's sqref top-left
/// as the reference cell.
struct CfValueObject {
  CfvoType type = CfvoType::Number;
  /// Source representation. For `Number`/`Percent`/`Percentile` this is
  /// the textual form (`"50"`, `"75.5"`); for `Formula` it is the raw
  /// formula source without the leading `=`. `Min`/`Max`/`AutoMin`/
  /// `AutoMax` ignore this field.
  std::string value;
  /// `gte` controls the boundary comparison used by `iconSet` bucketing
  /// (`true` => `>=`; `false` => `>`). `colorScale` rules ignore this.
  bool gte = true;
};

/// `<colorScale>` sub-element of a `cfRule`. Two-stop scales use 2
/// thresholds + 2 colours; three-stop scales use 3 of each. Lengths must
/// match in any valid CF set.
struct ColorScaleSpec {
  std::vector<CfValueObject> thresholds;
  std::vector<Color> colors;
};

/// `<dataBar>` sub-element of a `cfRule`.
struct DataBarSpec {
  CfValueObject min{};
  CfValueObject max{};
  Color fill{};
  std::optional<Color> border;
  Color negative_fill{};
  std::optional<Color> negative_border;
  DataBarAxisPosition axis_position = DataBarAxisPosition::Automatic;
  Color axis_color{0, 0, 0, 255};
  bool gradient = true;
  bool show_value = true;
  /// `minLength` / `maxLength` from OOXML, expressed as 0-100 percent.
  std::uint8_t min_length_pct = 10;
  std::uint8_t max_length_pct = 90;
};

/// `<iconSet>` sub-element of a `cfRule`.
struct IconSetSpec {
  IconSetName name = IconSetName::Three_Arrows;
  /// N-1 thresholds for an N-icon set. For example, the default
  /// `Three_Arrows` set carries two thresholds (33% / 67%).
  std::vector<CfValueObject> thresholds;
  bool reverse = false;
  bool show_value = true;
  /// `percent` attribute on `<iconSet>`: when `true`, the thresholds are
  /// interpreted as percent of (max - min); when `false`, as plain
  /// numbers. Defaulted to `true` to match Excel-emitted XML.
  bool percent = true;
};

/// One `<cfRule>` definition. Matches the union of all rule types: the
/// engaged optional fields are determined by `type`. For example,
/// `cellIs` rules carry `op` and `formula1` (+ `formula2` for `Between`/
/// `NotBetween`); `colorScale` rules carry `color_scale`; etc.
struct CFRule {
  /// Stable identifier for the rule, read from the `id` attribute on the
  /// base `<cfRule id="...">` element itself (not from any `<x14:cfRule>`
  /// extension, which this engine does not parse). Excel emits this
  /// attribute directly on the legacy `<cfRule>` when the rule has a
  /// richer `<x14:cfRule>` counterpart elsewhere in `<extLst>`, using it
  /// to cross-reference the two; the value happens to be the same
  /// GUID-shaped string `<x14:cfRule id="...">` carries. Empty for
  /// legacy rules with no such counterpart. The writer round-trips
  /// whatever value is present onto the base element.
  std::string id;
  RuleType type = RuleType::Expression;
  /// Workbook-global priority. Smaller numbers evaluate first.
  std::int32_t priority = 1;
  /// `stopIfTrue` attribute. When `true`, a successful match prevents
  /// later (higher-priority-number) rules from being evaluated for the
  /// same cell, even across sibling `<conditionalFormatting>` blocks.
  bool stop_if_true = false;
  /// Index into `styles.dxfs` populated by the OOXML reader. `nullopt`
  /// for visual rule types (colorScale / dataBar / iconSet) that
  /// describe their own rendering.
  std::optional<std::uint32_t> dxf_id;

  /// Primary formula source. Always present for `expression` and
  /// `cellIs`; optional for `containsText`/`beginsWith`/`endsWith`/
  /// `top10`/`aboveAverage` (Excel emits a derived formula but the
  /// canonical decision is on `op` + side-channel fields).
  std::optional<std::string> formula1;
  /// Secondary formula (Between / NotBetween).
  std::optional<std::string> formula2;
  /// Comparison operator for `cellIs` rules.
  std::optional<CellIsOperator> op;

  /// Visual-rule payloads. Engaged at most one at a time, matching the
  /// rule `type`.
  std::optional<ColorScaleSpec> color_scale;
  std::optional<DataBarSpec> data_bar;
  std::optional<IconSetSpec> icon_set;

  /// `top10` payload: rank, percent flag, bottom flag.
  std::optional<std::int32_t> rank;
  bool percent = false;
  bool bottom = false;

  /// `aboveAverage` payload.
  bool above_average = true;
  bool equal_average = false;
  std::optional<double> std_dev;

  /// `containsText`/`beginsWith`/`endsWith`/`notContainsText` literal.
  std::optional<std::string> text;

  /// `timePeriod` bucket.
  std::optional<TimePeriod> time_period;

  /// Verbatim, unparsed `<extLst>...</extLst>` XML captured from this
  /// rule's own `<cfRule>` element (`CT_CfRule`'s schema-trailing
  /// `extLst?`). Typically carries a 2010+ `<x14:id>` cross-reference to
  /// a richer `<x14:cfRule>` counterpart, or other forward-compat
  /// extension content this engine does not model. Round-tripped
  /// byte-for-byte by the writer; never interpreted. `nullopt` when the
  /// rule has no `<extLst>` child.
  std::optional<std::string> ext_lst_raw;
};

/// One `<conditionalFormatting>` block — the sqref union plus its rule
/// list. `pivot_scope` mirrors the OOXML attribute (rules attached to a
/// pivot table area rather than free-standing cells).
struct ConditionalFormat {
  std::vector<CFCellRange> sqref;
  std::vector<CFRule> rules;
  bool pivot_scope = false;

  /// Verbatim, unparsed `<extLst>...</extLst>` XML captured from this
  /// block's own `<conditionalFormatting>` element (`CT_ConditionalFormatting`'s
  /// schema-trailing `extLst?`, a sibling of the block's `<cfRule>`
  /// children — distinct from each rule's own `CFRule::ext_lst_raw`).
  /// Round-tripped byte-for-byte by the writer; never interpreted.
  /// `nullopt` when the block has no `<extLst>` child.
  std::optional<std::string> ext_lst_raw;
};

}  // namespace formulon::cf

#endif  // FORMULON_CF_CF_TYPES_H_
