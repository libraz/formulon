//
// Pivot data-model primitives shared by the workbook layer, the OOXML
// reader/writer, the pivot evaluator, and `GETPIVOTDATA`.
//
// This header is intentionally type-only: no behaviour, no allocation
// strategy, no evaluator hooks. Subsequent PRs build the cache, evaluator,
// and lookup paths on top of these structures.

#ifndef FORMULON_PIVOT_PIVOT_TYPES_H_
#define FORMULON_PIVOT_PIVOT_TYPES_H_

#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <variant>
#include <vector>

namespace formulon::pivot {

/// Where a pivot field is positioned within the report layout.
///
/// `None` is a field that appears in `<pivotFields>` but is placed on no
/// axis and is not a data field — an "available" field in Excel's field
/// list. It must be distinguished from `Value` so the writer does not
/// stamp `dataField="1"` onto an unused field on round trip.
enum class PivotAxis : std::uint8_t {
  Row = 0,
  Col = 1,
  Value = 2,
  Page = 3,
  None = 4,
};

/// Aggregation function applied to a `Value` field.
enum class Aggregation : std::uint8_t {
  Sum = 0,
  Count = 1,
  Average = 2,
  Max = 3,
  Min = 4,
  Product = 5,
  CountNumbers = 6,
  StdDev = 7,
  StdDevP = 8,
  Var = 9,
  VarP = 10,
};

/// Subtotal function. Currently mirrors `Aggregation` one-to-one;
/// kept distinct so the OOXML subtotal-only options (e.g. `default`)
/// can be added later without breaking the value-aggregation enum.
enum class SubtotalFn : std::uint8_t {
  Sum = 0,
  Count = 1,
  Average = 2,
  Max = 3,
  Min = 4,
  Product = 5,
  CountNumbers = 6,
  StdDev = 7,
  StdDevP = 8,
  Var = 9,
  VarP = 10,
};

/// Visual layout mode for the rendered pivot.
enum class PivotLayout : std::uint8_t {
  Compact = 0,
  Tabular = 1,
  Outline = 2,
};

/// Filter type applied to a pivot axis.
///
/// Range-typed filters (`ValueBetween`, `LabelDate`) read both
/// `PivotFilter::value` (low bound) and `PivotFilter::value_high`
/// (high bound). The other filter types only consume `value`.
enum class FilterType : std::uint8_t {
  ValueTop10 = 0,
  ValueGreaterThan = 1,
  ValueBetween = 2,
  LabelContains = 3,
  LabelBeginsWith = 4,
  LabelDate = 5,
};

/// Granularity at which a date column is grouped.
enum class DateGrouping : std::uint8_t {
  Day = 0,
  Month = 1,
  Quarter = 2,
  Year = 3,
  Week = 4,
  Hour = 5,
  Minute = 6,
  Second = 7,
};

/// Calendar system used when grouping date fields.
enum class CalendarSystem : std::uint8_t {
  Gregorian = 0,
  Japanese = 1,
};

/// One distinct value of a pivot field, with manual filter visibility.
///
/// `visible` reflects the XML-defined manual filter state authored in the
/// pivot definition. Slicer-driven (transient) selection lives on
/// `PivotTable::active_filters_`, not here.
struct PivotItem {
  std::string name;
  bool visible = true;

  /// Index into the bound cache field's `shared_items`, taken from the
  /// source `<item x="N">`. The OOXML reader captures this so it can
  /// (a) resolve `name` against the cache after load and (b) re-emit the
  /// exact same index on write. `has_cache_index` distinguishes an
  /// explicit `x` attribute from the implicit sequential order Excel uses
  /// when it omits the attribute; when false the writer falls back to the
  /// item's document-order position. Hand-built items (C API, tests) leave
  /// both at their defaults.
  bool has_cache_index = false;
  std::uint32_t cache_index = 0;
};

/// One entry of the OOXML `<pageFields>` block: a field placed on the page
/// (report-filter) axis, plus the item it is currently showing.
///
/// The block exists because `<pivotFields>` records only that a field sits
/// on the page axis, never in which order the page fields stack above the
/// report nor which of their items is selected.
struct PivotPageField {
  /// Index into `PivotTable::fields()`, from `<pageField fld="N">`.
  std::uint32_t field_index = 0;

  /// The single selected item, from `<pageField item="N">`. Absent means no
  /// single item is selected: either every item is showing, or several are.
  /// Excel expresses a multi-item selection through the field's own hidden
  /// `PivotItem`s rather than through this attribute, so the two cases are
  /// told apart by item visibility, not by presence.
  ///
  /// ECMA-376 words the index as "the item in the PivotTable field", which
  /// reads as a position in `PivotField::items` rather than a cache
  /// shared-item index; the two coincide for the ordinary field whose items
  /// are listed in shared-item order. Resolution follows the literal
  /// reading.
  std::optional<std::uint32_t> item_index;
};

/// Configuration for grouping a date-typed source column.
struct PivotDateGroup {
  DateGrouping granularity = DateGrouping::Year;
  CalendarSystem calendar = CalendarSystem::Gregorian;
  std::optional<int> start_year;
  std::optional<int> end_year;
};

/// Single filter clause attached to a pivot field.
///
/// The `value` payload covers the common Top-N, threshold, and label
/// cases via an int / double / string variant. Range-typed filters
/// (`ValueBetween`, `LabelDate`) additionally consume the optional
/// `value_high` upper-bound payload; other filter types ignore it.
struct PivotFilter {
  PivotAxis axis = PivotAxis::Row;
  std::string field_name;
  FilterType type = FilterType::ValueTop10;
  std::variant<int, double, std::string> value;
  /// Upper-bound payload for range filters:
  ///   * `ValueBetween` — `value` is the inclusive low bound, `value_high` is
  ///     the inclusive high bound (both numeric).
  ///   * `LabelDate` — `value` is the start date serial (numeric), `value_high`
  ///     is the end date serial (inclusive).
  /// Unused for the other filter types. The default `std::monostate`
  /// signals "unbounded above"; range filters with no upper bound
  /// degrade to a no-op.
  std::variant<std::monostate, int, double> value_high;
  /// Index into `PivotTable::data_fields()` whose aggregate is scored by a
  /// value filter. Label/date filters ignore this selector. Keeping the
  /// default at zero preserves the original single-measure behaviour.
  std::uint32_t data_field_index = 0;
};

/// Comparison an authored OOXML `<filters>` caption filter applies to a
/// field's rendered labels.
///
/// This mirrors the caption half of `ST_PivotFilterType` one-to-one.
/// It is deliberately separate from `FilterType`: `FilterType` is the
/// embedder-facing surface exposed through the C ABI and the bindings,
/// while these values only ever originate from a file the reader
/// decodes, so widening them costs no ABI compatibility.
///
/// Ordering comparisons (`GreaterThan` .. `NotBetween`) compare labels
/// as text, which is what Excel does for a caption filter even when the
/// labels happen to look numeric.
enum class CaptionPredicate : std::uint8_t {
  Equal = 0,
  NotEqual = 1,
  BeginsWith = 2,
  NotBeginsWith = 3,
  EndsWith = 4,
  NotEndsWith = 5,
  Contains = 6,
  NotContains = 7,
  GreaterThan = 8,
  GreaterThanOrEqual = 9,
  LessThan = 10,
  LessThanOrEqual = 11,
  Between = 12,
  NotBetween = 13,
};

/// One decoded `<filter>` entry from an authored `<filters>` block.
///
/// `field_index` is the source `fld` attribute: an index into
/// `<pivotFields>`, which OOXML keeps parallel to the bound cache's
/// fields. `value_high` carries the upper bound for `Between` /
/// `NotBetween` and is unused otherwise.
struct AuthoredCaptionFilter {
  std::uint32_t field_index = 0;
  CaptionPredicate predicate = CaptionPredicate::Equal;
  std::string value;
  std::string value_high;
};

/// Which quantity the top-N dialog counts out.
///
/// Excel's "Top 10" dialog offers three, and writes each as a different
/// `<filter type>` over an identically shaped `<top10 val="N">`: `count`,
/// `percent`, `sum`. Nothing inside `<top10>` distinguishes `count` from
/// `sum` at all, so the type attribute is the only discriminator and
/// reading one as the other silently returns different rows.
///
/// This is deliberately not folded into `FilterType`: that enum is the
/// embedder-facing surface exposed through the C ABI, and these values
/// only ever originate from a file the reader decodes.
enum class TopNBasis : std::uint8_t {
  /// `<filter type="count">` — keep the N highest-scoring leaves.
  Items = 0,
  /// `<filter type="percent">` — keep the highest-scoring leaves whose
  /// running total first reaches N percent of the axis total.
  Percent = 1,
  /// `<filter type="sum">` — the same running-total rule against an
  /// absolute target rather than a share.
  Sum = 2,
};

/// One decoded value-or-date `<filter>` entry from an authored
/// `<filters>` block.
///
/// Sibling of `AuthoredCaptionFilter`, and separate from `PivotFilter`
/// for the same reason: `PivotFilter` is the embedder-facing slicer
/// surface reached through the C ABI, and clearing a slicer selection
/// must not clear a rule that came out of the file.
///
/// It reuses `FilterType` rather than introducing another enum because
/// each decoded family already has an evaluation counterpart there, and
/// the reader only produces a member it can evaluate:
///
///   * `ValueTop10` — `<filter type="count">`, `value` is the item count
///     from the nested `<top10 val="N">`.
///   * `ValueGreaterThan` — `<filter type="valueGreaterThan">`, `value`
///     is the exclusive threshold.
///   * `ValueBetween` — `<filter type="valueBetween">`, inclusive bounds
///     in `value` / `value_high`.
///   * `LabelDate` — `<filter type="dateBetween">`, inclusive date
///     serials in `value` / `value_high`.
///
/// The axis is not stored: a file names only the field, and which axis
/// the pruning applies to follows from that field's membership in
/// `row_field_order()` / `col_field_order()`.
struct AuthoredValueFilter {
  /// Source `fld` attribute: an index into `<pivotFields>`, which OOXML
  /// keeps parallel to the bound cache's fields.
  std::uint32_t field_index = 0;
  FilterType type = FilterType::ValueTop10;
  double value = 0.0;
  /// Upper bound for the range families; absent otherwise. A range entry
  /// with no upper bound degrades to a no-op, matching `PivotFilter`.
  std::optional<double> value_high;
  /// Source `iMeasureFld`: which data field's aggregate is scored.
  std::uint32_t data_field_index = 0;
  /// Which quantity `value` counts out. Only meaningful when `type` is
  /// `ValueTop10`, which is the one family the dialog offers a choice for.
  TopNBasis top_n_basis = TopNBasis::Items;
};

/// A date window named relative to when the pivot is computed.
///
/// The `<filters>` block spells these as a bare type with no criteria —
/// `<filter type="dateThisMonth" fld="2"/>` — because the window is
/// implied by the name. Resolving one therefore needs a clock reading,
/// which is why they are held apart from `AuthoredValueFilter` (whose
/// bounds are literals in the file) and why the workbook can pin the
/// reading they resolve against.
///
/// Every family here spans one contiguous range. The recurring
/// `M1`..`M12` / `Q1`..`Q4` selectors are not part of it: they pick every
/// January rather than one January, so they are not a window at all and
/// are modelled by `AuthoredRecurringFilter`.
///
/// The week group runs Sunday through Saturday, and the three windows
/// tile without gap or overlap. Excel resolves them against the calendar
/// week rather than a rolling seven days from today, so `ThisWeek` on a
/// Friday still starts on the preceding Sunday.
enum class RelativePeriod : std::uint8_t {
  Today = 0,
  Yesterday = 1,
  Tomorrow = 2,
  ThisMonth = 3,
  LastMonth = 4,
  NextMonth = 5,
  ThisQuarter = 6,
  LastQuarter = 7,
  NextQuarter = 8,
  ThisYear = 9,
  LastYear = 10,
  NextYear = 11,
  YearToDate = 12,
  ThisWeek = 13,
  LastWeek = 14,
  NextWeek = 15,
};

/// An authored relative-period date filter.
///
/// Sibling of `AuthoredValueFilter`, split off because its window is not
/// in the file: it is resolved at evaluation time from the workbook's
/// clock. Applied pre-aggregation, exactly like the absolute `dateBetween`
/// family it degenerates to once resolved.
struct AuthoredPeriodFilter {
  /// Source `fld` attribute: an index into `<pivotFields>`.
  std::uint32_t field_index = 0;
  RelativePeriod period = RelativePeriod::Today;
};

/// An authored recurring-period date filter: the `M1`..`M12` and
/// `Q1`..`Q4` families.
///
/// Held apart from `AuthoredPeriodFilter` because these select a calendar
/// position rather than a range. `M1` keeps every January of every year,
/// so no `DateWindow` can express it and no clock reading is needed to
/// resolve it — the criterion is simply which month a record's date falls
/// in. Both families reduce to one inclusive month range because a
/// quarter is three adjacent months: `Q2` is months 4..6.
struct AuthoredRecurringFilter {
  /// Source `fld` attribute: an index into `<pivotFields>`.
  std::uint32_t field_index = 0;
  /// Inclusive 1-based calendar month bounds. `low == high` for `M<n>`.
  unsigned month_low = 1;
  unsigned month_high = 1;
};

/// Sort directive for a pivot field.
struct SortSpec {
  bool ascending = true;
  /// Empty selects display-label ordering. Otherwise this matches a data
  /// field's display name or source/custom field name and sorts sibling items
  /// by that field's aggregate.
  std::string by_field;
};

/// "Show values as" derivation applied to each cell after the raw
/// aggregation finishes. Mirrors Excel's `dataField/@showDataAs`
/// attribute. `Normal` leaves the aggregate unchanged.
///
/// Modes:
///   * `PercentOfRow` / `PercentOfCol` / `PercentOfTotal` — divide each
///     cell by its row, column, or grand total.
///   * `RunningTotalInRow` / `RunningTotalInCol` — running cumulative
///     sum along the named axis.
///   * `Index` — `(cell * grand_total) / (row_sum * col_sum)`.
///   * `DifferenceFrom` / `PercentDifferenceFrom` — absolute or relative
///     difference from a reference cell along a named base axis. The
///     reference is selected by `PivotDataField::show_as_base_field`
///     (which axis: row vs col) and `show_as_base_item` (which item
///     within that field, including the `(previous)` / `(next)`
///     sentinels).
///   * `PercentOfParentRow` / `PercentOfParentCol` / `PercentOfParent` —
///     divide each cell by the subtotal of its parent group within the
///     row hierarchy, the column hierarchy, or whichever axis hosts the
///     `show_as_base_field` (PercentOfParent only).
///
/// When the show-as mode requires the grand total (`PercentOfTotal`,
/// `Index`) but the table has both `grand_totals_rows == false` and
/// `grand_totals_cols == false`, the transform falls back to recomputing
/// the total over the surviving values matrix.
enum class ShowValuesAs : std::uint8_t {
  Normal = 0,
  PercentOfRow = 1,
  PercentOfCol = 2,
  PercentOfTotal = 3,
  RunningTotalInRow = 4,
  RunningTotalInCol = 5,
  Index = 6,
  DifferenceFrom = 7,
  PercentDifferenceFrom = 8,
  PercentOfParentRow = 9,
  PercentOfParentCol = 10,
  PercentOfParent = 11,
};

/// Sentinel `baseItem` value meaning "the previous item along the base
/// field's axis". Mirrors Excel's literal value (1048828) and round-trips
/// through OOXML verbatim.
inline constexpr std::uint32_t kShowAsBasePrev = 1048828U;

/// Sentinel `baseItem` value meaning "the next item along the base
/// field's axis". Mirrors Excel's literal value (1048829) and round-trips
/// through OOXML verbatim.
inline constexpr std::uint32_t kShowAsBaseNext = 1048829U;

/// One data-field entry from `<dataFields>/<dataField>`.
///
/// GETPIVOTDATA looks up data fields by `name` (e.g. "Sum of Amount").
/// `field_index` is the index into `PivotTable::fields()` of the
/// pivot-field this data-field aggregates from. A single source field
/// can be referenced by multiple data fields (e.g. Sum + Average of the
/// same column), each with its own display name.
struct PivotDataField {
  std::string name;  ///< Display name; key for GETPIVOTDATA lookup.
  std::uint32_t field_index = 0;
  Aggregation aggregation = Aggregation::Sum;
  std::string number_format;
  ShowValuesAs show_as = ShowValuesAs::Normal;

  /// For `DifferenceFrom` / `PercentDifferenceFrom`: index into
  /// `PivotTable::fields()` of the "base field" whose items are the
  /// reference axis. For `PercentOfParent`: the parent field in the row
  /// or column hierarchy that this aggregation is normalised against.
  /// `std::nullopt` means "use previous on the show-as axis" (the OOXML
  /// default for an unset `baseField` on Difference variants).
  std::optional<std::uint32_t> show_as_base_field;

  /// For `DifferenceFrom` / `PercentDifferenceFrom`: which item of the
  /// base field is the reference. Two reserved sentinels:
  ///   * `kShowAsBasePrev` (1048828) — "(previous)"
  ///   * `kShowAsBaseNext` (1048829) — "(next)"
  /// Any other value is an index into the base field's `items[]`.
  std::optional<std::uint32_t> show_as_base_item;
};

/// Field-level configuration as authored in the OOXML pivot definition.
///
/// `aggregations` is non-empty only for fields placed on
/// `PivotAxis::Value`; for other axes it is intentionally empty.
struct PivotField {
  std::string source_name;
  std::string custom_name;
  PivotAxis axis = PivotAxis::Row;
  std::vector<Aggregation> aggregations;
  SortSpec sort;
  std::vector<PivotItem> items;
  /// Position of the subtotal row relative to its group: true = above
  /// (the OOXML `subtotalTop` default), false = below. Layout flag only;
  /// whether a subtotal exists is governed by `default_subtotal` /
  /// `subtotal_fns`.
  bool subtotal_top = true;

  /// Custom subtotal functions selected on this field, mirroring the
  /// OOXML `<pivotField>` `*Subtotal` boolean attribute family
  /// (`sumSubtotal`, `avgSubtotal`, ...). Empty when only the implicit
  /// default subtotal applies.
  std::vector<SubtotalFn> subtotal_fns;

  /// Mirrors the OOXML `defaultSubtotal` attribute (default `true`).
  /// When `false` the field suppresses its automatic default subtotal;
  /// any entries in `subtotal_fns` then describe the explicit custom set.
  bool default_subtotal = true;

  std::string number_format;
  std::optional<PivotDateGroup> date_group;

  /// OOXML `<pivotField>` attributes the model does not represent
  /// structurally (e.g. `compact`, `outline`, `showAll`,
  /// `includeNewItemsInFilter`), captured as `(name, value)` pairs so the
  /// writer re-emits them verbatim. Rendering keys off `PivotTable::layout`,
  /// so these are preserved for round-trip only.
  std::vector<std::pair<std::string, std::string>> passthrough_attrs;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_TYPES_H_
