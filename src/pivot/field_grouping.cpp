#include "pivot/field_grouping.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

#include "pivot/date_serial.h"
#include "pivot/pivot_cache.h"
#include "pivot/record_access.h"
#include "utils/checked_index.h"
#include "utils/date_time.h"
#include "utils/iso_date.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// Days before each month in a leap year: a days group numbers 2/29 as day 60
// in every year, so 3/1 is day 61 whether or not the year has a 2/29.
constexpr unsigned kLeapDaysBeforeMonth[12] = {0U, 31U, 60U, 91U, 121U, 152U, 182U, 213U, 244U, 274U, 305U, 335U};

date_time::YMD ymd_of_day(double day, bool date1904) noexcept {
  return date1904 ? date_time::ymd_from_serial(day, /*date1904=*/true) : date_time::legacy_1900_ymd(day);
}

// Clamps a 1-based in-range group number onto the items between the two
// out-of-range catch-alls.
std::uint32_t inner_item(double number, std::uint32_t last) noexcept {
  const double cap = last >= 2U ? static_cast<double>(last - 1U) : static_cast<double>(last);
  if (!(number >= 1.0)) {
    return 1U;
  }
  return static_cast<std::uint32_t>(std::min(number, cap));
}

// A date group's bound that is absent or unparsable falls back to the base
// data: the first day, and the day after the last (clamped to the last valid
// date), as Excel's Group dialog proposes them.
GroupBounds resolve_bounds(const PivotCache& cache, std::size_t base, const RangeGroup& range, bool date1904) {
  if (const std::optional<GroupBounds> stated = stated_group_bounds(range, date1904)) {
    return *stated;
  }
  const std::optional<double> start = group_date_serial(range.start_date, date1904);
  const std::optional<double> end = group_date_serial(range.end_date, date1904);
  bool have_data = false;
  double data_min = 0.0;
  double data_max = 0.0;
  for (const PivotCacheRecord& record : cache.records()) {
    const Value v = cell_value(cache, record, base);
    if (!v.is_number() || !is_valid_date_serial(v.as_number(), date1904)) {
      continue;
    }
    const double day = std::floor(v.as_number());
    data_min = have_data ? std::min(data_min, day) : day;
    data_max = have_data ? std::max(data_max, day) : day;
    have_data = true;
  }
  GroupBounds bounds;
  bounds.start = start.has_value() ? *start : data_min;
  if (end.has_value()) {
    bounds.end = *end;
  } else {
    const double auto_end = have_data ? data_max + 1.0 : bounds.start;
    bounds.end = std::min(auto_end, last_valid_date_serial(date1904));
  }
  return bounds;
}

// The base field's shared-item index of `record`'s value, which a discrete
// group's map is keyed by.
std::optional<std::size_t> base_item_index(const PivotCache& cache, const PivotCacheRecord& record, std::size_t base) {
  const PivotCacheField& base_field = cache.fields()[base];
  if (is_group_field(base_field)) {
    const Value v = cell_value(cache, record, base);
    return v.is_number() ? index_from_double(v.as_number(), base_field.field_group->group_items.size()) : std::nullopt;
  }
  if (base < record.cells.size()) {
    const Value& cell = record.cells[base];
    const bool is_index = base < record.cell_is_index.size() ? record.cell_is_index[base]
                                                             : (!base_field.shared_items.empty() && cell.is_number());
    if (is_index) {
      return cell.is_number() ? index_from_double(cell.as_number(), base_field.shared_items.size()) : std::nullopt;
    }
  }
  const Value v = cell_value(cache, record, base);
  for (std::size_t i = 0; i < base_field.shared_items.size(); ++i) {
    if (base_field.shared_items[i] == v) {
      return i;
    }
  }
  return std::nullopt;
}

void store_group_cell(PivotCacheRecord& record, std::size_t field_index, std::size_t field_count, Value number) {
  if (record.cells.size() < field_count) {
    record.cells.resize(field_count, Value::blank());
  }
  record.cells[field_index] = number;
  if (!record.cell_is_index.empty()) {
    record.cell_is_index.resize(record.cells.size(), false);
    record.cell_is_index[field_index] = false;
  }
  if (field_index < record.cell_is_date.size()) {
    record.cell_is_date[field_index] = false;
  }
}

}  // namespace

bool is_group_field(const PivotCacheField& field) noexcept {
  return !field.is_database_field && field.field_group.has_value() && field.field_group->base.has_value() &&
         field.field_group->kind != GroupKind::None;
}

std::optional<double> group_date_serial(std::string_view iso, bool date1904) noexcept {
  double serial = 0.0;
  if (!date_time::parse_iso_date_serial(iso, &serial)) {
    return std::nullopt;
  }
  return date1904 ? serial - date_time::kDate1904EpochGap : serial;
}

std::optional<GroupBounds> stated_group_bounds(const RangeGroup& range, bool date1904) noexcept {
  if (!range.is_date) {
    return GroupBounds{range.start_num, range.end_num};
  }
  const std::optional<double> start = group_date_serial(range.start_date, date1904);
  const std::optional<double> end = group_date_serial(range.end_date, date1904);
  if (!start.has_value() || !end.has_value()) {
    return std::nullopt;
  }
  return GroupBounds{*start, *end};
}

std::uint32_t range_group_item_number(const RangeGroup& range, const GroupBounds& bounds, std::size_t item_count,
                                      const Value& value, bool date1904) noexcept {
  if (item_count == 0U || !value.is_number()) {
    return 0U;
  }
  const auto last = static_cast<std::uint32_t>(item_count - 1U);
  const double v = value.as_number();
  if (range.by == GroupBy::Range) {
    if (!std::isfinite(v) || v < bounds.start) {
      return 0U;
    }
    if (v > bounds.end) {
      return last;
    }
    const double width = range.interval > 0.0 ? range.interval : 1.0;
    return inner_item(std::floor((v - bounds.start) / width) + 1.0, last);
  }
  if (!is_valid_date_serial(v, date1904)) {
    return 0U;
  }
  if (range.by == GroupBy::Seconds || range.by == GroupBy::Minutes || range.by == GroupBy::Hours) {
    if (v < bounds.start) {
      return 0U;
    }
    if (v > bounds.end) {
      return last;
    }
    const date_time::HMS hms = date_time::hms_from_fraction(v);
    const unsigned unit = range.by == GroupBy::Hours ? hms.h : (range.by == GroupBy::Minutes ? hms.m : hms.s);
    return inner_item(static_cast<double>(unit) + 1.0, last);
  }
  // Day-granular groups compare whole days, so a time on the start or end day
  // stays inside the range.
  const double day = std::floor(v);
  const double start_day = std::floor(bounds.start);
  if (day < start_day) {
    return 0U;
  }
  if (day > std::floor(bounds.end)) {
    return last;
  }
  const date_time::YMD ymd = ymd_of_day(day, date1904);
  switch (range.by) {
    case GroupBy::Days: {
      if (range.interval > 1.0) {
        return inner_item(std::floor((day - start_day) / std::floor(range.interval)) + 1.0, last);
      }
      const unsigned month = ymd.m >= 1U && ymd.m <= 12U ? ymd.m : 1U;
      return inner_item(static_cast<double>(kLeapDaysBeforeMonth[month - 1U] + ymd.d), last);
    }
    case GroupBy::Months:
      return inner_item(static_cast<double>(ymd.m), last);
    case GroupBy::Quarters:
      return inner_item(static_cast<double>((ymd.m + 2U) / 3U), last);
    case GroupBy::Years: {
      if (!is_valid_date_serial(start_day, date1904)) {
        return 0U;
      }
      const int start_year = ymd_of_day(start_day, date1904).y;
      return inner_item(static_cast<double>(ymd.y - start_year) + 1.0, last);
    }
    default:
      return 0U;
  }
}

std::optional<std::uint32_t> discrete_group_item_number(const FieldGroup& group, std::size_t base_item) noexcept {
  if (group.kind != GroupKind::Discrete || base_item >= group.discrete_map.size()) {
    return std::nullopt;
  }
  return group.discrete_map[base_item];
}

void materialize_group_fields(PivotCache& cache, bool date1904) {
  const std::size_t field_count = cache.fields().size();
  for (std::size_t fi = 0; fi < field_count; ++fi) {
    const PivotCacheField& field = cache.fields()[fi];
    if (!is_group_field(field)) {
      continue;
    }
    const FieldGroup& group = *field.field_group;
    const std::size_t base = *group.base;
    if (base >= field_count || base == fi) {
      continue;
    }
    const GroupBounds bounds =
        group.kind == GroupKind::Range ? resolve_bounds(cache, base, group.range, date1904) : GroupBounds{};
    for (PivotCacheRecord& record : cache.mutable_records()) {
      Value number = Value::blank();
      if (group.kind == GroupKind::Range) {
        number = Value::number(range_group_item_number(group.range, bounds, group.group_items.size(),
                                                       cell_value(cache, record, base), date1904));
      } else if (const std::optional<std::size_t> item = base_item_index(cache, record, base)) {
        if (const std::optional<std::uint32_t> n = discrete_group_item_number(group, *item)) {
          number = Value::number(*n);
        }
      }
      store_group_cell(record, fi, field_count, number);
    }
  }
}

}  // namespace formulon::pivot
