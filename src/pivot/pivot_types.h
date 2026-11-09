//
// Pivot data-model primitives shared by the workbook layer, the OOXML
// reader/writer, the pivot evaluator, and `GETPIVOTDATA`.
//
// This header is intentionally type-only: no behaviour, no allocation
// strategy, no evaluator hooks. Subsequent PRs build the cache, evaluator,
// and lookup paths on top of these structures.

#ifndef FORMULON_PIVOT_PIVOT_TYPES_H_
#define FORMULON_PIVOT_PIVOT_TYPES_H_

#include <cstddef>
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

/// The `Aggregation` with the same ordinal as `fn` (the two enums share their
/// layout, Sum=0 .. VarP=10).
inline Aggregation aggregation_from_subtotal_fn(SubtotalFn fn) {
  return static_cast<Aggregation>(static_cast<std::uint8_t>(fn));
}

/// Visual layout mode for the rendered pivot.
enum class PivotLayout : std::uint8_t {
  Compact = 0,
  Tabular = 1,
  Outline = 2,
};

/// Filter type applied to a pivot field: every `ST_PivotFilterType` member
/// except `unknown`.
///
/// Families and their payloads (`PivotFilter::value` / `value_high`):
///   * Value (`ValueTop10`, `ValueGreaterThan`, `ValueBetween`, 6-11, 24-28):
///     ranked on a data field's aggregate after aggregation; numeric payloads.
///     The top / bottom families keep items while the count, percent of the
///     field total or running sum is not yet reached.
///   * Caption (`LabelContains`, `LabelBeginsWith`, 12-23): matched against
///     the item label per record; text payloads as entered, `~` escapes kept.
///   * Date (`LabelDate` = `dateBetween`, 29-35): compared against the
///     record's date serial; numeric serial payloads.
///   * Relative period and calendar position (36-67): no payload.
///
/// Range types (`*Between`, `*NotBetween`, `LabelDate`) read both bounds.
enum class FilterType : std::uint8_t {
  ValueTop10 = 0,
  ValueGreaterThan = 1,
  ValueBetween = 2,
  LabelContains = 3,
  LabelBeginsWith = 4,
  LabelDate = 5,
  ValueEqual = 6,
  ValueNotEqual = 7,
  ValueGreaterThanOrEqual = 8,
  ValueLessThan = 9,
  ValueLessThanOrEqual = 10,
  ValueNotBetween = 11,
  LabelEqual = 12,
  LabelNotEqual = 13,
  LabelNotBeginsWith = 14,
  LabelEndsWith = 15,
  LabelNotEndsWith = 16,
  LabelNotContains = 17,
  LabelGreaterThan = 18,
  LabelGreaterThanOrEqual = 19,
  LabelLessThan = 20,
  LabelLessThanOrEqual = 21,
  LabelBetween = 22,
  LabelNotBetween = 23,
  ValueBottomCount = 24,
  ValueTopPercent = 25,
  ValueBottomPercent = 26,
  ValueTopSum = 27,
  ValueBottomSum = 28,
  DateEqual = 29,
  DateNotEqual = 30,
  DateOlderThan = 31,
  DateOlderThanOrEqual = 32,
  DateNewerThan = 33,
  DateNewerThanOrEqual = 34,
  DateNotBetween = 35,
  DateTomorrow = 36,
  DateToday = 37,
  DateYesterday = 38,
  DateNextWeek = 39,
  DateThisWeek = 40,
  DateLastWeek = 41,
  DateNextMonth = 42,
  DateThisMonth = 43,
  DateLastMonth = 44,
  DateNextQuarter = 45,
  DateThisQuarter = 46,
  DateLastQuarter = 47,
  DateNextYear = 48,
  DateThisYear = 49,
  DateLastYear = 50,
  DateYearToDate = 51,
  DateQ1 = 52,
  DateQ2 = 53,
  DateQ3 = 54,
  DateQ4 = 55,
  DateM1 = 56,
  DateM2 = 57,
  DateM3 = 58,
  DateM4 = 59,
  DateM5 = 60,
  DateM6 = 61,
  DateM7 = 62,
  DateM8 = 63,
  DateM9 = 64,
  DateM10 = 65,
  DateM11 = 66,
  DateM12 = 67,
};

/// Calendar a date field's years are grouped by: Gregorian years, or the
/// Japanese era years.
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

  /// `<item n>`: the caption the user gave the item, shown instead of its
  /// cache label. Empty when the item is not renamed. Kept apart from `name`,
  /// which hidden-item matching and page labels resolve against the cache.
  std::string caption{};

  /// `<item sd>`: false when the item is collapsed, hiding the items of the
  /// inner fields below it.
  bool show_detail = true;

  /// `<item f>`: the item is a calculated item.
  bool calculated = false;

  /// `<item>` attributes the model does not represent (`s`, `m`, `c`,
  /// `d`, `e`), captured as `(name, value)` pairs so the writer re-emits them
  /// verbatim.
  std::vector<std::pair<std::string, std::string>> passthrough_attrs{};
};

/// A non-data `<item t="...">` of a `<pivotField>`'s `<items>`: a subtotal
/// marker (`default`, `sum`, `countA`, ...), `grand` or `blank`, kept
/// verbatim so its attributes survive a save.
///
/// The subtotal markers restate `PivotField::default_subtotal` /
/// `subtotal_fns`; the writer re-emits them only while they still agree with
/// that selection and regenerates them otherwise.
struct PivotMarkerItem {
  std::string xml;
  /// The `t` attribute.
  std::string type;
  /// Number of real items that preceded it. Anchors `blank` among the items;
  /// subtotal markers and `grand` always close the list.
  std::size_t after_items = 0;
  /// The `sd` attribute. A `default` marker whose value no longer matches
  /// `xml` is regenerated by the writer.
  bool show_detail = true;
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

  /// `<pageField hier="N">`; -1 is the non-OLAP marker Excel writes for a
  /// field bound to a cache field. Absent when the file left it out.
  std::optional<std::int32_t> hierarchy = -1;

  /// `<pageField name>` and `<pageField cap>` (OLAP member name and caption),
  /// kept verbatim; empty when absent.
  std::string name{};
  std::string caption{};
};

/// The `<filter>` element a loaded `PivotFilter` was decoded from.
///
/// The writer re-emits `xml` verbatim while the entry still designates the
/// field and data field it was read with, so a load -> save cycle keeps
/// Excel's own bytes (`evalOrder`, `id`, `stringValue1`, ...). Once either
/// index moves the entry is re-serialised from the model instead.
struct PivotFilterSource {
  std::string xml;
  /// Source `fld`: the index into `<pivotFields>` the filter applies to.
  std::uint32_t field_index = 0;
  /// Source `iMeasureFld`, or 0 when absent.
  std::uint32_t data_field_index = 0;
  /// Source `id`, kept so generated siblings never reuse it.
  std::uint32_t id = 0;
};

/// Single filter clause attached to a pivot field.
///
/// `value` / `value_high` carry the payload the type's family takes (see
/// `FilterType`); `std::monostate` is an unset slot. A range filter with no
/// upper bound filters nothing.
struct PivotFilter {
  PivotAxis axis = PivotAxis::Row;
  std::string field_name;
  FilterType type = FilterType::ValueTop10;
  std::variant<std::monostate, int, double, std::string> value;
  std::variant<std::monostate, int, double, std::string> value_high;
  /// Index into `PivotTable::data_fields()` whose aggregate is scored by a
  /// value filter. Caption and date filters ignore it.
  std::uint32_t data_field_index = 0;
  /// Set for an entry decoded from a file's `<filters>`; empty for one added
  /// through the API.
  std::optional<PivotFilterSource> source{};
};

/// A loaded `<filter>` the `PivotFilter` model cannot express: an unknown
/// type, a member-property or OLAP measure filter (`mpFld`,
/// `iMeasureHier`), or one whose criteria do not decode. Kept verbatim so a
/// save never drops it; it filters nothing.
struct RetainedPivotFilter {
  std::string xml;
  /// Number of `PivotTable::active_filters()` entries that preceded this one
  /// in the source block, which fixes its slot on save.
  std::size_t active_before = 0;
  /// Source `id`, kept so generated siblings never reuse it.
  std::uint32_t id = 0;
};

/// Sort directive for a pivot field.
struct SortSpec {
  bool ascending = true;
  /// True when sibling order at this field's hierarchy level comes from
  /// `PivotField::items` document order rather than from label or
  /// `by_field` order. Mirrors an explicit OOXML `sortType="manual"`.
  /// `ascending` is not consulted while this is set.
  bool manual = false;
  /// Empty selects display-label ordering. Otherwise this matches a data
  /// field's display name or source/custom field name and sorts sibling items
  /// by that field's aggregate.
  std::string by_field;
  /// True when the order was stated rather than defaulted: read from an
  /// explicit `sortType`, or set through `fm_sheet_set_pivot_field_sort`.
  /// The writer spells `sortType` out for a field on an axis or one whose
  /// order was stated, so a sorted field taken off the report keeps its order.
  bool stated = false;
};

/// Sentinel `field x=` value on `<rowFields>`/`<colFields>` naming the
/// Values (Sigma) pseudo-field position -- where the header selecting
/// among two or more data fields sits in the row/column hierarchy. Never
/// a real index into `PivotTable::fields()`, so it is kept out of
/// `PivotTable::row_field_order()` / `col_field_order()` entirely; its
/// document position lives in `PivotTable::row_values_position()` /
/// `col_values_position()` instead. The rendering of multiple data
/// fields is already driven by `PivotTable::data_fields()` independent
/// of this marker, so nothing downstream of the reader needs to
/// recognise the value `-2` itself.
inline constexpr std::int32_t kValuesFieldPosition = -2;

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
///   * `PercentOf` — each cell as a fraction of the base item's cell.
///   * `PercentOfRunningTotal` — running total along the base field over
///     the total across that field.
///   * `RankAscending` / `RankDescending` — dense rank among the base
///     field's siblings.
///
/// When the show-as mode requires the grand total (`PercentOfTotal`,
/// `Index`) the denominator is the whole-table aggregate re-aggregated from
/// the records with the data field's own aggregation, whether or not the
/// grand-total strips are displayed.
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
  PercentOf = 12,
  PercentOfRunningTotal = 13,
  RankAscending = 14,
  RankDescending = 15,
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
  /// Any other value is a position in the base field's `items[]` list.
  std::optional<std::uint32_t> show_as_base_item;

  /// `<ext>` children of the element's `<extLst>` other than the show-as
  /// extension the model decodes, verbatim.
  std::string raw_ext_xml{};

  /// The rest of the show-as extension's `<x14:dataField>`: its attributes
  /// other than `pivotShowAs` (`sourceField`, `uniqueName`, ...) and its
  /// children, verbatim.
  std::vector<std::pair<std::string, std::string>> x14_attrs{};
  std::string x14_children_xml{};
};

/// Field-level configuration as authored in the OOXML pivot definition.
struct PivotField {
  std::string source_name;
  std::string custom_name;
  PivotAxis axis = PivotAxis::Row;
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

  /// `showAll`: items without data are listed. A loaded field takes the
  /// attribute (absent = true); a field created in memory starts off.
  bool show_all = false;

  /// Mirrors the OOXML `defaultSubtotal` attribute (default `true`).
  /// When `false` the field suppresses its automatic default subtotal;
  /// any entries in `subtotal_fns` then describe the explicit custom set.
  bool default_subtotal = true;

  /// `<pivotField numFmtId>` as its decimal string; empty when absent.
  std::string number_format;

  /// OOXML `<pivotField>` attributes the model does not represent
  /// structurally (e.g. `compact`, `outline`,
  /// `includeNewItemsInFilter`), captured as `(name, value)` pairs so the
  /// writer re-emits them verbatim. Rendering keys off `PivotTable::layout`,
  /// so these are preserved for round-trip only.
  ///
  /// `compact` / `outline` land here only when they depart from the form the
  /// table's `layout` implies (a field Excel's field settings gave its own
  /// form); the writer then emits them instead of the derived pair.
  std::vector<std::pair<std::string, std::string>> passthrough_attrs;

  /// `insertBlankRow`: a blank row follows each of the field's item blocks.
  bool insert_blank_row = false;

  /// `fillDownLabels` of the x14 `<pivotField>` extension: the field's label is
  /// repeated on every row of its block.
  bool fill_down_labels = false;

  /// Children after `<items>` the model does not represent (`<autoSortScope>`),
  /// verbatim and in document order.
  std::string raw_children_xml{};

  /// `<ext>` children of the element's `<extLst>` other than the x14
  /// `<pivotField>` extension, verbatim.
  std::string raw_ext_xml{};

  /// The rest of the x14 `<pivotField>` extension: its attributes other than
  /// `fillDownLabels`, and its children, verbatim.
  std::vector<std::pair<std::string, std::string>> x14_attrs{};
  std::string x14_children_xml{};

  /// The non-data entries of `<items>`, in document order.
  std::vector<PivotMarkerItem> marker_items{};
};

/// The field's custom subtotal functions in the fixed order Excel lists
/// them (enum order, duplicates dropped) rather than the order they were
/// added. Layout rows and the written `<items>` markers both follow it.
inline std::vector<SubtotalFn> ordered_subtotal_fns(const PivotField& field) {
  // A presence mask over the closed enum instead of sort + unique.
  std::uint32_t present = 0;
  for (const SubtotalFn fn : field.subtotal_fns) {
    present |= 1U << static_cast<unsigned>(fn);
  }
  std::vector<SubtotalFn> fns;
  for (unsigned i = 0; (present >> i) != 0U; ++i) {
    if ((present & (1U << i)) != 0U) {
      fns.push_back(static_cast<SubtotalFn>(i));
    }
  }
  return fns;
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_TYPES_H_
