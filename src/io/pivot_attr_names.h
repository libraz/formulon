//
// OOXML attribute spellings of the pivot-table and pivot-cache enums, shared by
// the readers and writers.

#ifndef FORMULON_IO_PIVOT_ATTR_NAMES_H_
#define FORMULON_IO_PIVOT_ATTR_NAMES_H_

#include <array>
#include <cstddef>
#include <string_view>

#include "io/enum_name_table.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_types.h"

namespace formulon::io {

/// `<pivotField axis="...">` spellings. `Value` and `None` have none: a Value
/// field is written as `dataField="1"` and an unused field carries no axis.
inline constexpr std::array<EnumName, 3> kPivotAxisNames = {{
    {pivot::PivotAxis::Row, "axisRow"},
    {pivot::PivotAxis::Col, "axisCol"},
    {pivot::PivotAxis::Page, "axisPage"},
}};

/// `<dataField subtotal="...">` spellings.
inline constexpr std::array<EnumName, 11> kPivotAggregationNames = {{
    {pivot::Aggregation::Sum, "sum"},
    {pivot::Aggregation::Count, "count"},
    {pivot::Aggregation::Average, "average"},
    {pivot::Aggregation::Max, "max"},
    {pivot::Aggregation::Min, "min"},
    {pivot::Aggregation::Product, "product"},
    {pivot::Aggregation::CountNumbers, "countNums"},
    {pivot::Aggregation::StdDev, "stdDev"},
    {pivot::Aggregation::StdDevP, "stdDevp"},
    {pivot::Aggregation::Var, "var"},
    {pivot::Aggregation::VarP, "varp"},
}};
static_assert(kPivotAggregationNames.size() == static_cast<std::size_t>(pivot::Aggregation::VarP) + 1U);

/// `<dataField showDataAs="...">` spellings; `Normal` is the omitted default.
/// `runTotalInCol` is the Excel-private column-direction spelling the reader
/// accepts; the writer emits `runTotal` for both directions.
inline constexpr std::array<EnumName, 9> kPivotShowDataAsNames = {{
    {pivot::ShowValuesAs::PercentOfRow, "percentOfRow"},
    {pivot::ShowValuesAs::PercentOfCol, "percentOfCol"},
    {pivot::ShowValuesAs::PercentOfTotal, "percentOfTotal"},
    {pivot::ShowValuesAs::RunningTotalInRow, "runTotal"},
    {pivot::ShowValuesAs::RunningTotalInCol, "runTotalInCol"},
    {pivot::ShowValuesAs::Index, "index"},
    {pivot::ShowValuesAs::DifferenceFrom, "difference"},
    {pivot::ShowValuesAs::PercentDifferenceFrom, "percentDiff"},
    {pivot::ShowValuesAs::PercentOf, "percent"},
}};

/// Modes Excel stores in the `<x14:dataField pivotShowAs="...">` extension of a `<dataField>`
/// instead of `showDataAs`.
inline constexpr std::array<EnumName, 6> kPivotShowAsX14Names = {{
    {pivot::ShowValuesAs::PercentOfParentRow, "percentOfParentRow"},
    {pivot::ShowValuesAs::PercentOfParentCol, "percentOfParentCol"},
    {pivot::ShowValuesAs::PercentOfParent, "percentOfParent"},
    {pivot::ShowValuesAs::PercentOfRunningTotal, "percentOfRunningTotal"},
    {pivot::ShowValuesAs::RankAscending, "rankAscending"},
    {pivot::ShowValuesAs::RankDescending, "rankDescending"},
}};

/// `<ext uri>` of the extension carrying `kPivotShowAsX14Names`.
inline constexpr std::string_view kPivotDataFieldExtUri = "{E15A36E0-9728-4e99-A89B-3F7291B0FE68}";

/// `<rangePr groupBy="...">` spellings.
inline constexpr std::array<EnumName, 8> kPivotGroupByNames = {{
    {pivot::GroupBy::Range, "range"},
    {pivot::GroupBy::Seconds, "seconds"},
    {pivot::GroupBy::Minutes, "minutes"},
    {pivot::GroupBy::Hours, "hours"},
    {pivot::GroupBy::Days, "days"},
    {pivot::GroupBy::Months, "months"},
    {pivot::GroupBy::Quarters, "quarters"},
    {pivot::GroupBy::Years, "years"},
}};
static_assert(kPivotGroupByNames.size() == static_cast<std::size_t>(pivot::GroupBy::Years) + 1U);

/// The `ST_PivotFilterType` spelling of each `FilterType`, by ordinal. The
/// top / bottom pairs share one spelling; `<top10 top="0">` tells them apart.
inline constexpr std::array<std::string_view, 68> kFilterTypeNames = {{
    "count",
    "valueGreaterThan",
    "valueBetween",
    "captionContains",
    "captionBeginsWith",
    "dateBetween",
    "valueEqual",
    "valueNotEqual",
    "valueGreaterThanOrEqual",
    "valueLessThan",
    "valueLessThanOrEqual",
    "valueNotBetween",
    "captionEqual",
    "captionNotEqual",
    "captionNotBeginsWith",
    "captionEndsWith",
    "captionNotEndsWith",
    "captionNotContains",
    "captionGreaterThan",
    "captionGreaterThanOrEqual",
    "captionLessThan",
    "captionLessThanOrEqual",
    "captionBetween",
    "captionNotBetween",
    "count",
    "percent",
    "percent",
    "sum",
    "sum",
    "dateEqual",
    "dateNotEqual",
    "dateOlderThan",
    "dateOlderThanOrEqual",
    "dateNewerThan",
    "dateNewerThanOrEqual",
    "dateNotBetween",
    "tomorrow",
    "today",
    "yesterday",
    "nextWeek",
    "thisWeek",
    "lastWeek",
    "nextMonth",
    "thisMonth",
    "lastMonth",
    "nextQuarter",
    "thisQuarter",
    "lastQuarter",
    "nextYear",
    "thisYear",
    "lastYear",
    "yearToDate",
    "Q1",
    "Q2",
    "Q3",
    "Q4",
    "M1",
    "M2",
    "M3",
    "M4",
    "M5",
    "M6",
    "M7",
    "M8",
    "M9",
    "M10",
    "M11",
    "M12",
}};
static_assert(kFilterTypeNames.size() == static_cast<std::size_t>(pivot::FilterType::DateM12) + 1U);

/// `<ext uri>` of the x14 `<pivotField>` extension, spelled as Excel writes it.
inline constexpr std::string_view kPivotFieldExtUri = "{2946ED86-A175-432a-8AC1-64E0C546D7DE}";

/// One subtotal function's `<pivotField>` boolean attribute (`sumSubtotal`) and
/// its `<item t="...">` token (`sum`) inside the field's `<items>` list.
struct PivotSubtotalName {
  pivot::SubtotalFn fn;
  std::string_view attr;
  std::string_view item;
};

/// Subtotal spellings in the canonical ECMA-376 attribute order.
inline constexpr std::array<PivotSubtotalName, 11> kPivotSubtotalNames = {{
    {pivot::SubtotalFn::Sum, "sumSubtotal", "sum"},
    {pivot::SubtotalFn::Count, "countASubtotal", "countA"},
    {pivot::SubtotalFn::Average, "avgSubtotal", "avg"},
    {pivot::SubtotalFn::Max, "maxSubtotal", "max"},
    {pivot::SubtotalFn::Min, "minSubtotal", "min"},
    {pivot::SubtotalFn::Product, "productSubtotal", "product"},
    {pivot::SubtotalFn::CountNumbers, "countSubtotal", "count"},
    {pivot::SubtotalFn::StdDev, "stdDevSubtotal", "stdDev"},
    {pivot::SubtotalFn::StdDevP, "stdDevPSubtotal", "stdDevP"},
    {pivot::SubtotalFn::Var, "varSubtotal", "var"},
    {pivot::SubtotalFn::VarP, "varPSubtotal", "varP"},
}};
static_assert(kPivotSubtotalNames.size() == static_cast<std::size_t>(pivot::SubtotalFn::VarP) + 1U);

/// Returns the spellings of `fn`, or `nullptr` when it has none.
inline const PivotSubtotalName* find_pivot_subtotal_name(pivot::SubtotalFn fn) {
  for (const PivotSubtotalName& entry : kPivotSubtotalNames) {
    if (entry.fn == fn) {
      return &entry;
    }
  }
  return nullptr;
}

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_ATTR_NAMES_H_
