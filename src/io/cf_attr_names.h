//
// OOXML attribute spellings of the conditional-formatting enums, shared by the
// CF reader and writer. Entry order is the enum's ordinal order.

#ifndef FORMULON_IO_CF_ATTR_NAMES_H_
#define FORMULON_IO_CF_ATTR_NAMES_H_

#include <array>

#include "cf/cf_types.h"
#include "io/enum_name_table.h"

namespace formulon::io {

/// `<cfRule type="...">` spellings.
inline constexpr std::array<EnumName, 18> kCfRuleTypeNames = {{
    {cf::RuleType::Expression, "expression"},
    {cf::RuleType::CellIs, "cellIs"},
    {cf::RuleType::ColorScale, "colorScale"},
    {cf::RuleType::DataBar, "dataBar"},
    {cf::RuleType::IconSet, "iconSet"},
    {cf::RuleType::Top10, "top10"},
    {cf::RuleType::AboveAverage, "aboveAverage"},
    {cf::RuleType::ContainsText, "containsText"},
    {cf::RuleType::NotContainsText, "notContainsText"},
    {cf::RuleType::BeginsWith, "beginsWith"},
    {cf::RuleType::EndsWith, "endsWith"},
    {cf::RuleType::ContainsBlanks, "containsBlanks"},
    {cf::RuleType::NotContainsBlanks, "notContainsBlanks"},
    {cf::RuleType::ContainsErrors, "containsErrors"},
    {cf::RuleType::NotContainsErrors, "notContainsErrors"},
    {cf::RuleType::TimePeriod, "timePeriod"},
    {cf::RuleType::DuplicateValues, "duplicateValues"},
    {cf::RuleType::UniqueValues, "uniqueValues"},
}};
static_assert(kCfRuleTypeNames.size() == static_cast<std::size_t>(cf::RuleType::UniqueValues) + 1U);

/// `<cfRule operator="...">` spellings.
inline constexpr std::array<EnumName, 8> kCfCellIsOperatorNames = {{
    {cf::CellIsOperator::LessThan, "lessThan"},
    {cf::CellIsOperator::LessThanOrEqual, "lessThanOrEqual"},
    {cf::CellIsOperator::Equal, "equal"},
    {cf::CellIsOperator::NotEqual, "notEqual"},
    {cf::CellIsOperator::GreaterThanOrEqual, "greaterThanOrEqual"},
    {cf::CellIsOperator::GreaterThan, "greaterThan"},
    {cf::CellIsOperator::Between, "between"},
    {cf::CellIsOperator::NotBetween, "notBetween"},
}};
static_assert(kCfCellIsOperatorNames.size() == static_cast<std::size_t>(cf::CellIsOperator::NotBetween) + 1U);

/// `<cfvo type="...">` spellings.
inline constexpr std::array<EnumName, 8> kCfvoTypeNames = {{
    {cf::CfvoType::Number, "num"},
    {cf::CfvoType::Percent, "percent"},
    {cf::CfvoType::Percentile, "percentile"},
    {cf::CfvoType::Min, "min"},
    {cf::CfvoType::Max, "max"},
    {cf::CfvoType::Formula, "formula"},
    {cf::CfvoType::AutoMin, "autoMin"},
    {cf::CfvoType::AutoMax, "autoMax"},
}};
static_assert(kCfvoTypeNames.size() == static_cast<std::size_t>(cf::CfvoType::AutoMax) + 1U);

/// `<cfRule timePeriod="...">` spellings.
inline constexpr std::array<EnumName, 10> kCfTimePeriodNames = {{
    {cf::TimePeriod::Today, "today"},
    {cf::TimePeriod::Yesterday, "yesterday"},
    {cf::TimePeriod::Tomorrow, "tomorrow"},
    {cf::TimePeriod::Last7Days, "last7Days"},
    {cf::TimePeriod::ThisWeek, "thisWeek"},
    {cf::TimePeriod::LastWeek, "lastWeek"},
    {cf::TimePeriod::NextWeek, "nextWeek"},
    {cf::TimePeriod::ThisMonth, "thisMonth"},
    {cf::TimePeriod::LastMonth, "lastMonth"},
    {cf::TimePeriod::NextMonth, "nextMonth"},
}};
static_assert(kCfTimePeriodNames.size() == static_cast<std::size_t>(cf::TimePeriod::NextMonth) + 1U);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_ATTR_NAMES_H_
