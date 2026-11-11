//
// Group-item numbering for pivot group fields.
//
// A group field (`databaseField="0"` with `<fieldGroup base>`) has no source
// column; each record's value is the index into its `<groupItems>` that the
// base field's value falls in. A range group reserves item 0 for values below
// its start and the last item for values above its end. The numbers are
// materialised into `PivotCacheRecord::cells` once, so every consumer reads a
// group field through `cell_value` like any database field.
//
// Grouping a date field appends one group field per calendar period, finest
// first, each labelled in the workbook's locale when it is created; the base
// field's `par` names the last one. Ungrouping removes them again, which is
// only possible while they are the trailing fields of the cache.

#ifndef FORMULON_PIVOT_FIELD_GROUPING_H_
#define FORMULON_PIVOT_FIELD_GROUPING_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "excel_profile.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
class Workbook;
}  // namespace formulon

namespace formulon::pivot {

/// Inclusive bounds of a range group: date serials for a date group, plain
/// numbers otherwise.
struct GroupBounds {
  double start = 0.0;
  double end = 0.0;
};

/// True when `field` is a group field whose cells are derived from a base.
bool is_group_field(const PivotCacheField& field) noexcept;

/// Serial of an ISO 8601 group bound (`startDate` / `endDate`) under the given
/// epoch; empty when `iso` does not parse.
std::optional<double> group_date_serial(std::string_view iso, bool date1904) noexcept;

/// Bounds `range` states: its parsed `startDate` / `endDate`, or
/// `start_num` / `end_num` for a numeric group. Empty when a date bound is
/// absent or does not parse.
std::optional<GroupBounds> stated_group_bounds(const RangeGroup& range, bool date1904) noexcept;

/// Group-item number of base value `value` under `range` with `bounds` and
/// `item_count` group items: 0 below the start, the last item above the end,
/// and 0 for a blank or a value outside the date domain. A numeric value
/// equal to the end falls in the last range, even below the start, and a
/// quotient within 15 significant digits of a range edge counts as on it.
std::uint32_t range_group_item_number(const RangeGroup& range, const GroupBounds& bounds, std::size_t item_count,
                                      const Value& value, bool date1904) noexcept;

/// Group-item number of the base field's shared item `base_item` under a
/// discrete group; empty when the map does not cover it.
std::optional<std::uint32_t> discrete_group_item_number(const FieldGroup& group, std::size_t base_item) noexcept;

/// Writes every group field's group-item number into each record's cell for
/// that field, as a non-index, non-date number. A date range group without
/// stated bounds takes the base data's first day and the day after its last.
void materialize_group_fields(PivotCache& cache, bool date1904);

/// Bit of calendar period `period` in a date-group period mask: seconds 1,
/// minutes 2, hours 4, days 8, months 16, quarters 32, years 64.
constexpr std::uint32_t date_period_bit(GroupBy period) noexcept {
  return 1U << (static_cast<unsigned>(period) - 1U);
}

/// Every calendar period a date group can take.
inline constexpr std::uint32_t kAllDatePeriods = 0x7FU;

/// A date grouping to create over a base field.
struct DateGroupSpec {
  /// `date_period_bit` of each period to group by.
  std::uint32_t periods = 0;
  /// `Japanese` groups the years period by era year.
  CalendarSystem calendar = CalendarSystem::Gregorian;
  /// Days per item of a group by days alone.
  std::uint32_t interval_days = 1;
  /// First and last date serial; empty takes the base data's.
  std::optional<double> start;
  std::optional<double> end;
};

/// The base field of cache field `field`: its `base` when it is a group
/// field, else `field` itself.
std::size_t group_base_index(const PivotCache& cache, std::size_t field) noexcept;

/// Period mask of the date group fields over base field `base`; 0 when the
/// field is not date-grouped.
std::uint32_t date_group_periods(const PivotCache& cache, std::size_t base, bool date1904);

/// Calendar of the date grouping over `base`: `Japanese` when its years are
/// grouped by era year.
CalendarSystem date_group_calendar(const PivotCache& cache, std::size_t base, bool date1904);

/// Bounds Excel proposes for base values spanning `min_serial` ..
/// `max_serial`: the minimum as it is, and the maximum when it carries a time
/// of day, else the midnight after it (no later than the last valid date).
GroupBounds auto_date_group_bounds(double min_serial, double max_serial, bool date1904) noexcept;

/// ISO 8601 body of `serial` to the second (`2024-01-15T09:30:15`), the form
/// `startDate`, `minDate` and `<d>` carry; the inverse of `group_date_serial`.
std::string group_date_iso(double serial, bool date1904);

/// Appends to `cache` one group field per period of `spec` over database
/// field `base`, finest first, named and labelled in `profile`'s locale,
/// points the base's `par` at the last one and materialises them. Returns
/// the index of the first appended field.
///
/// `kInvalidArgument` for an empty or unknown period mask, a zero interval,
/// an interval with any period but days, a bound that is not a valid date or
/// a start after the end. `kPreconditionFailed` for a base that holds a
/// blank, a value that is not a date, or no value at all, a calculated base,
/// and a base already grouped.
Expected<std::size_t, Error> group_cache_by_date(PivotCache& cache, std::size_t base, const DateGroupSpec& spec,
                                                 bool date1904, ExcelProfile profile);

/// Removes the group fields over `base` from `cache` and clears the base's
/// `par`. Returns how many were removed (0 when `base` is not grouped).
/// `kPreconditionFailed` when they are not the trailing fields of the cache.
Expected<std::size_t, Error> ungroup_cache_field(PivotCache& cache, std::size_t base);

/// Groups field `field` (a base field or one of its group fields) of pivot
/// `pivot` on sheet `sheet` by date, replacing any grouping it has. Every
/// pivot on the same cache gains the group fields. On this pivot they take
/// the base's place on its axis, coarsest outermost; on the others they are
/// inserted before the base, which stays innermost. Errors as
/// `group_cache_by_date`, plus `kSheetIndexOutOfRange` / `kIndexOutOfRange` for an index out of range.
Expected<void, Error> set_date_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field,
                                     std::uint32_t periods, CalendarSystem calendar, std::uint32_t interval_days,
                                     std::optional<double> start, std::optional<double> end);

/// Removes the date grouping of `field` (a base field or one of its group
/// fields) from every pivot on its cache; the base returns to the axis
/// position its group fields held. A field without grouping is left as is.
/// `kSheetIndexOutOfRange` / `kIndexOutOfRange` for an index out of range, `kPreconditionFailed` as
/// `ungroup_cache_field`.
Expected<void, Error> clear_date_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field);

/// A date grouping as stored, for reading back: the settings `set_date_group`
/// takes, with the stored bounds even where they are automatic.
struct DateGroupState {
  std::uint32_t periods = 0;
  CalendarSystem calendar = CalendarSystem::Gregorian;
  std::uint32_t interval_days = 1;
  double start = 0.0;
  double end = 0.0;
  bool auto_start = true;
  bool auto_end = true;
};

/// A numeric grouping as stored: `startNum`, `endNum`, `groupInterval` and
/// whether each bound follows the data.
struct NumericGroupState {
  double start = 0.0;
  double end = 0.0;
  double interval = 1.0;
  bool auto_start = true;
  bool auto_end = true;
};

/// The date grouping over cache field `field` (a base field or one of its
/// group fields); empty when it is not date-grouped.
std::optional<DateGroupState> date_group_state(const PivotCache& cache, std::size_t field, bool date1904);

/// The numeric grouping over cache field `field` (a base field or its group
/// field); empty when it is not numerically grouped.
std::optional<NumericGroupState> numeric_group_state(const PivotCache& cache, std::size_t field);

/// Groups field `field` (a base field or its group field) of pivot `pivot` on
/// sheet `sheet` into numeric ranges of `interval`, replacing any grouping it
/// has, and re-projects every pivot on its cache. An empty bound follows the
/// base's present values. The group items, as a refresh regenerates them,
/// are `<start`, one range per step from `start` up to the first edge at or
/// past `end` (at least one), then `>` and the last range's upper edge:
///
/// - With start, end and interval whole, a range whose upper edge a+interval
///   is at most 1e15 reads `a-(a+interval-1)`, the last one `a-end` when
///   `end` is its upper edge.
/// - Every other range reads `a-(a+interval)`; the forms can mix in one field.
///
/// Each number is its 15-digit decimal view, spelled in full while the text
/// of its magnitude has at most 20 characters, else in exponent form
/// (`1.3E+20`, `2.43E-17`), with the workbook locale's decimal separator
/// (`0,3-0,8`).
///
/// `kInvalidArgument` for a non-finite or non-positive interval, a non-finite
/// bound or a start after the end; `kPreconditionFailed` for a base holding a
/// blank or a value that is not a number, holding no number to take an empty
/// bound from, a calculated base, or group fields over the base that are not
/// the last cache fields; `kSecurityResourceLimit` for more group
/// items than a field holds; `kSheetIndexOutOfRange` / `kIndexOutOfRange` for
/// an index out of range.
Expected<void, Error> set_numeric_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field,
                                        std::optional<double> start, std::optional<double> end, double interval);

/// Removes the numeric grouping of `field` from every pivot on its cache, as
/// `clear_date_group` does for a date grouping. A field without one is left
/// as it is.
Expected<void, Error> clear_numeric_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_FIELD_GROUPING_H_
