// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include <variant>
#include <vector>

namespace formulon::pivot {

/// Where a pivot field is positioned within the report layout.
enum class PivotAxis : std::uint8_t {
  Row = 0,
  Col = 1,
  Value = 2,
  Page = 3,
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
};

/// Sort directive for a pivot field.
struct SortSpec {
  bool ascending = true;
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
  bool subtotal_top = false;
  std::vector<SubtotalFn> subtotal_fns;
  std::string number_format;
  std::optional<PivotDateGroup> date_group;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_TYPES_H_
