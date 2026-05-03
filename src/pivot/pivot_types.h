// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Pivot data-model primitives shared by the workbook layer, the OOXML
// reader/writer, the pivot evaluator, and `GETPIVOTDATA`. See
// backup/plans/15-pivot-and-advanced.md §15.1 for the authoritative
// specification.
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
/// The `value` payload is intentionally minimal at this stage: an int /
/// double / string variant covers the common Top-N, threshold, and label
/// cases. Range-typed filters (`LabelDate`, `ValueBetween`) will extend
/// this variant in a follow-up PR.
struct PivotFilter {
  PivotAxis axis = PivotAxis::Row;
  std::string field_name;
  FilterType type = FilterType::ValueTop10;
  std::variant<int, double, std::string> value;
};

/// Sort directive for a pivot field.
struct SortSpec {
  bool ascending = true;
  std::string by_field;
};

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
