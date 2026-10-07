//
// OOXML attribute spellings of the pivot-table enums, shared by the pivot table
// reader and writer.

#ifndef FORMULON_IO_PIVOT_ATTR_NAMES_H_
#define FORMULON_IO_PIVOT_ATTR_NAMES_H_

#include <array>
#include <cstddef>
#include <string_view>

#include "io/enum_name_table.h"
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
inline constexpr std::array<EnumName, 11> kPivotShowDataAsNames = {{
    {pivot::ShowValuesAs::PercentOfRow, "percentOfRow"},
    {pivot::ShowValuesAs::PercentOfCol, "percentOfCol"},
    {pivot::ShowValuesAs::PercentOfTotal, "percentOfTotal"},
    {pivot::ShowValuesAs::RunningTotalInRow, "runTotal"},
    {pivot::ShowValuesAs::RunningTotalInCol, "runTotalInCol"},
    {pivot::ShowValuesAs::Index, "index"},
    {pivot::ShowValuesAs::DifferenceFrom, "difference"},
    {pivot::ShowValuesAs::PercentDifferenceFrom, "percentDiff"},
    {pivot::ShowValuesAs::PercentOfParentRow, "percentOfParentRow"},
    {pivot::ShowValuesAs::PercentOfParentCol, "percentOfParentCol"},
    {pivot::ShowValuesAs::PercentOfParent, "percentOfParent"},
}};

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
