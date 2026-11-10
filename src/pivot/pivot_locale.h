//
// Locale-driven labelling for the pivot-grid layout layer.
//
// The pivot projection in `src/pivot/pivot_layout.{h,cpp}` is locale-
// agnostic; it leaves the choice of placeholders ("Row Labels" /
// "Column Labels"), subtotal suffix, grand-total label, and data-field
// name template up to the caller. This translation unit centralises
// those strings per `ExcelProfile`, and exposes the helper that
// turns a profile into a populated `pivot::PivotLayoutOptions`.

#ifndef FORMULON_PIVOT_PIVOT_LOCALE_H_
#define FORMULON_PIVOT_PIVOT_LOCALE_H_

#include <cstddef>
#include <optional>
#include <string>
#include <string_view>

#include "excel_profile.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"

namespace formulon::pivot {

/// Returns the layout-option overrides appropriate for the locale of
/// the given Excel profile.
///
/// Labels ignore the host (Mac vs Windows): they follow the workbook's
/// display locale, not the running Excel binary. The profile itself is
/// carried through so item identity and ordering collate per profile.
/// Every locale renders Excel's compact / tabular / outline layout; the
/// field-name header projection of a default-constructed
/// `PivotLayoutOptions{}` is not reachable from a profile.
pivot::PivotLayoutOptions pivot_layout_options_for(ExcelProfile profile);

/// Returns the localized aggregation label used in pivot data-field
/// headers (e.g. "Sum of" / "合計"). The returned view points at
/// static storage and is safe to copy or compare.
std::string_view aggregation_label(pivot::Aggregation agg, ExcelProfile profile);

/// Returns the localized function word on a custom subtotal row ("East 最大",
/// "Somme East"). It equals `aggregation_label` except in zh-CN, whose
/// data-field captions carry a suffix the subtotal row omits.
std::string_view subtotal_function_label(pivot::Aggregation agg, ExcelProfile profile);

/// Returns the separator placed between the localized aggregation
/// label and the source field name in a data-field display name
/// ("合計 / Amount" uses " / "; "Sum of Amount" uses " of ").
std::string_view data_field_separator(ExcelProfile profile);

/// Formats the full data-field display name for the locale of the
/// given profile, e.g. "合計 / Amount" (ja-JP), "Sum of Amount" (en-US)
/// or "求和项:Amount" (zh-CN): the aggregation label, the locale's
/// separator, then the field name.
std::string data_field_display_name(pivot::Aggregation agg, std::string_view field_name, ExcelProfile profile);

/// Name Excel gives a new data field aggregating `field_index` with `agg`:
/// `data_field_display_name` over the field's display name, followed by 2, 3,
/// ... while it equals, under the profile's collation, another data field's
/// name (`exclude_data_field` aside) or a cache field's display name.
std::string data_field_default_name(const PivotTable& table, const PivotCache& cache, pivot::Aggregation agg,
                                    std::size_t field_index, ExcelProfile profile,
                                    std::optional<std::size_t> exclude_data_field = std::nullopt);

/// Name Excel gives the group field of calendar period `period` over the base
/// field `base_name` in the profile's locale, e.g. "月 (Date)" or "Months
/// (Date)".
std::string date_group_field_name(GroupBy period, std::string_view base_name, ExcelProfile profile);

/// The word a refresh names a field with blank header text by, followed by
/// the field's 1-based column in the source (`フィールド2`, `Field2`). Empty
/// for a locale whose word has not been measured.
std::optional<std::string_view> blank_header_field_word(ExcelProfile profile);

/// Label of the in-range group item `index` of a calendar-unit `period`, in
/// the order Excel lists them: second / minute 0-59, hour 0-23, day 0-365 of
/// a leap year, month 0-11, quarter 0-3. Empty for an index past the unit.
std::string date_group_unit_label(GroupBy period, unsigned index, ExcelProfile profile);

/// Label of the years group item for Gregorian `year` ("2024年", "2024").
std::string date_group_year_label(int year, ExcelProfile profile);

/// Label of the item collecting values before the group's start, given the
/// start rendered as a short date.
std::string date_group_before_label(std::string_view start_text);

/// Label of the item collecting values after the group's end, given the end
/// rendered as a short date.
std::string date_group_after_label(std::string_view end_text);

/// Label of one interval of a days group spanning several days, given its
/// first and last day rendered as short dates.
std::string date_group_interval_label(std::string_view first_text, std::string_view last_text);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_LOCALE_H_
