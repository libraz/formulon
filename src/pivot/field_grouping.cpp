// @size-budget: 32 KB

#include "pivot/field_grouping.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/text_format/display_text.h"
#include "excel_locale.h"
#include "num_fmt_locale.h"
#include "pivot/date_serial.h"
#include "pivot/field_grouping_internal.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_refresh.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "utils/cancellation_snap.h"
#include "utils/checked_index.h"
#include "utils/date_time.h"
#include "utils/double_format.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/iso_date.h"
#include "utils/japanese_era.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// 1900-03-01, the first serial past Excel's phantom 1900-02-29 (serial 60).
constexpr double kFirstSerialAfterPhantomDay = 61.0;

// The smallest and largest serial added; both 0 until the first.
struct SerialRange {
  bool any = false;
  double lo = 0.0;
  double hi = 0.0;

  void add(double serial) {
    lo = any ? std::min(lo, serial) : serial;
    hi = any ? std::max(hi, serial) : serial;
    any = true;
  }
};

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

// A date group's bound that is absent or unparsable falls back to the bounds
// Excel proposes for the base data.
GroupBounds resolve_bounds(const PivotCache& cache, std::size_t base, const RangeGroup& range, bool date1904) {
  if (const std::optional<GroupBounds> stated = stated_group_bounds(range, date1904)) {
    return *stated;
  }
  const std::optional<double> start = group_date_serial(range.start_date, date1904);
  const std::optional<double> end = group_date_serial(range.end_date, date1904);
  SerialRange data_range;
  for (const PivotCacheRecord& record : cache.records()) {
    const Value v = cell_value(cache, record, base);
    if (!v.is_number() || !is_valid_date_serial(v.as_number(), date1904)) {
      continue;
    }
    data_range.add(v.as_number());
  }
  const bool have_data = data_range.any;
  const GroupBounds data = have_data ? auto_date_group_bounds(data_range.lo, data_range.hi, date1904) : GroupBounds{};
  GroupBounds bounds;
  bounds.start = start.has_value() ? *start : data.start;
  bounds.end = end.has_value() ? *end : (have_data ? data.end : bounds.start);
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
    if (cell_is_shared_index(record, base_field, base)) {
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
  if (date1904) {
    return serial - date_time::kDate1904EpochGap;
  }
  // Excel counts these bodies from 1899-12-30 without its phantom 1900-02-29,
  // so 1900-02-28 is serial 60 and the earlier days sit one serial higher.
  return serial < kFirstSerialAfterPhantomDay ? serial + 1.0 : serial;
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
    if (!std::isfinite(v)) {
      return 0U;
    }
    // The end value falls in the last range even below the start; the end test precedes the start test.
    if (v == bounds.end) {
      return inner_item(static_cast<double>(last), last);
    }
    if (v > bounds.end) {
      return last;
    }
    if (v < bounds.start) {
      return 0U;
    }
    const double width = range.interval > 0.0 ? range.interval : 1.0;
    // A quotient within 15 digits of a whole number counts as that number (0.6 by 0.1 from 0.3 is 3).
    return inner_item(std::floor(excel_decimal_view((v - bounds.start) / width, 15)) + 1.0, last);
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
      return inner_item(static_cast<double>(kLeapDaysBefore[month - 1U] + ymd.d), last);
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

namespace {

using detail::group_error;
using detail::group_run;
using detail::GroupRun;
using detail::resolve_date_group_bounds;

bool flag_at(const std::vector<bool>& flags, std::size_t i) {
  return i < flags.size() && flags[i];
}

// The calendar periods of a date group, finest first.
constexpr GroupBy kDatePeriods[] = {GroupBy::Seconds, GroupBy::Minutes,  GroupBy::Hours, GroupBy::Days,
                                    GroupBy::Months,  GroupBy::Quarters, GroupBy::Years};

// One value of a base field and whether its source spelled it as a date.
struct BaseValue {
  Value value;
  bool is_date = false;
};

// The values a base field holds: its shared items, unused ones included, or
// its record cells when it has none.
std::vector<BaseValue> base_values(const PivotCache& cache, std::size_t base) {
  const PivotCacheField& field = cache.fields()[base];
  std::vector<BaseValue> values;
  if (!field.shared_items.empty()) {
    values.reserve(field.shared_items.size());
    for (std::size_t i = 0; i < field.shared_items.size(); ++i) {
      values.push_back({field.shared_items[i], flag_at(field.shared_item_is_date, i)});
    }
    return values;
  }
  values.reserve(cache.records().size());
  for (const PivotCacheRecord& record : cache.records()) {
    values.push_back({cell_value(cache, record, base), flag_at(record.cell_is_date, base)});
  }
  return values;
}

// Era-year label of the day holding `serial`: "令和6年".
std::string era_year_label(double serial, bool date1904, double* key) {
  const date_time::YMD ymd = ymd_of_day(std::floor(serial), date1904);
  const japanese_era::EraInfo& era = japanese_era::classify_era(ymd.y, ymd.m, ymd.d);
  const int era_year = ymd.y - era.year_anchor + 1;
  if (key != nullptr) {
    *key = static_cast<double>(era.year_anchor) * 10000.0 + static_cast<double>(era_year);
  }
  std::string label = era.kanji2;
  label.append(std::to_string(era_year));
  label.append("\xE5\xB9\xB4");  // 年
  return label;
}

// True when `field` is a discrete group whose items are the era years of its
// base's shared items.
bool is_era_year_group(const PivotCache& cache, const PivotCacheField& field, bool date1904) {
  if (!is_group_field(field) || field.field_group->kind != GroupKind::Discrete) {
    return false;
  }
  const FieldGroup& group = *field.field_group;
  const std::size_t base = *group.base;
  if (base >= cache.fields().size()) {
    return false;
  }
  const std::vector<Value>& items = cache.fields()[base].shared_items;
  if (items.empty() || items.size() != group.discrete_map.size()) {
    return false;
  }
  for (std::size_t i = 0; i < items.size(); ++i) {
    const std::uint32_t target = group.discrete_map[i];
    if (!items[i].is_number() || target >= group.group_items.size() || !group.group_items[target].is_text() ||
        group.group_items[target].as_text() != era_year_label(items[i].as_number(), date1904, nullptr)) {
      return false;
    }
  }
  return true;
}

// The values a base field holds, as `base_values`, without the shared items
// the source no longer has.
std::vector<BaseValue> present_base_values(const PivotCache& cache, std::size_t base) {
  const PivotCacheField& field = cache.fields()[base];
  if (field.shared_items.empty()) {
    return base_values(cache, base);
  }
  std::vector<BaseValue> values;
  for (std::size_t i = 0; i < field.shared_items.size(); ++i) {
    if (!flag_at(field.shared_item_is_missing, i)) {
      values.push_back({field.shared_items[i], flag_at(field.shared_item_is_date, i)});
    }
  }
  return values;
}

// Excel's unique-item limit of a pivot field.
constexpr std::size_t kMaxFieldItems = 1048576U;

// A range of whole settings reads `a-(a+interval-1)` while its upper edge a+interval is at most this,
// so that label keeps 15 digits; past it the range reads `a-(a+interval)` (Mac Excel, per range).
constexpr double kMaxIntegerLabelEdge = 1e15;

// A label number is spelled in full up to this many characters, else in exponent form
// (`90000000000000000000` vs `1E+20`, `0.000000000000000043` vs `2.43E-17`).
constexpr std::size_t kMaxPlainLabelChars = 20;

// True when start, end and interval of `state` are all whole numbers.
bool whole_settings(const NumericGroupState& state) noexcept {
  return std::floor(state.start) == state.start && std::floor(state.end) == state.end &&
         std::floor(state.interval) == state.interval;
}

// Ranges of `state` from the start up to the first edge at or past the end, at least one (also when
// the start follows the end), as a double so an oversized count is caught before any cast.
double numeric_bucket_count(const NumericGroupState& state) noexcept {
  const double span = (state.end - state.start) / state.interval;
  return std::max(1.0, std::ceil(excel_decimal_view(span, 15)));
}

// Edge `k` of `state`'s ranges.
double numeric_edge(const NumericGroupState& state, std::size_t k) noexcept {
  return snapped_add(state.start, static_cast<double>(k) * state.interval);
}

// Label text of one number: 15 significant digits, spelled in full while the magnitude's text fits
// `kMaxPlainLabelChars` (a minus sign is not counted), else in exponent form; `separator` is the
// decimal point.
std::string label_number(double v, char separator) {
  std::string text;
  format_positional(text, v);
  if (text.size() - (v < 0.0 ? 1U : 0U) > kMaxPlainLabelChars) {
    text.clear();
    format_scientific(text, v);
  }
  std::replace(text.begin(), text.end(), '.', separator);
  return text;
}

}  // namespace

namespace detail {

Expected<GroupBounds, Error> resolve_date_group_bounds(const PivotCache& cache, std::size_t base,
                                                       const DateGroupSpec& spec, bool date1904) {
  if (base >= cache.fields().size()) {
    return group_error(FormulonErrorCode::kIndexOutOfRange, "no such cache field",
                       "field=" + std::to_string(base) + " count=" + std::to_string(cache.fields().size()));
  }
  if (spec.periods == 0U || (spec.periods & ~kAllDatePeriods) != 0U) {
    return group_error(FormulonErrorCode::kInvalidArgument, "periods must be a non-empty mask of date periods",
                       "periods=" + std::to_string(spec.periods));
  }
  if (spec.interval_days == 0U) {
    return group_error(FormulonErrorCode::kInvalidArgument, "interval_days must be non-zero");
  }
  if (spec.interval_days > 1U && spec.periods != date_period_bit(GroupBy::Days)) {
    return group_error(FormulonErrorCode::kInvalidArgument, "an interval of several days needs the days period alone",
                       "periods=" + std::to_string(spec.periods));
  }
  if (spec.calendar != CalendarSystem::Gregorian && spec.calendar != CalendarSystem::Japanese) {
    return group_error(FormulonErrorCode::kInvalidArgument, "unknown calendar");
  }
  for (const std::optional<double>& bound : {spec.start, spec.end}) {
    if (bound.has_value() && !is_valid_date_serial(*bound, date1904)) {
      return group_error(FormulonErrorCode::kInvalidArgument, "a date-group bound must be a valid date serial");
    }
  }
  const PivotCacheField& field = cache.fields()[base];
  if (!field.is_database_field) {
    return group_error(FormulonErrorCode::kPreconditionFailed, "date grouping needs a database field",
                       "field=" + std::to_string(base));
  }
  // A field read from a file states its types, and a date mark declares one;
  // a value not spelled as a date is then a number.
  bool typed = field.shared_items_hints.present;
  for (std::size_t i = 0; i < field.shared_item_is_date.size() && !typed; ++i) {
    typed = field.shared_item_is_date[i];
  }
  for (const PivotCacheRecord& record : cache.records()) {
    typed = typed || flag_at(record.cell_is_date, base);
  }
  SerialRange range;
  for (const BaseValue& v : base_values(cache, base)) {
    if (!v.value.is_number() || !is_valid_date_serial(v.value.as_number(), date1904) || (typed && !v.is_date)) {
      return group_error(FormulonErrorCode::kPreconditionFailed,
                         "only a field holding nothing but dates can be grouped", "field=" + field.name);
    }
    range.add(v.value.as_number());
  }
  if (!range.any) {
    return group_error(FormulonErrorCode::kPreconditionFailed, "the field holds no date to group",
                       "field=" + field.name);
  }
  GroupBounds bounds = auto_date_group_bounds(range.lo, range.hi, date1904);
  const SharedItemsHints& hints = field.shared_items_hints;
  if (hints.has_min_date && hints.has_max_date) {
    const std::optional<double> min_date = group_date_serial(hints.min_date, date1904);
    const std::optional<double> max_date = group_date_serial(hints.max_date, date1904);
    if (min_date.has_value() && max_date.has_value()) {
      bounds = GroupBounds{*min_date, *max_date};
    }
  }
  bounds.start = spec.start.value_or(bounds.start);
  bounds.end = spec.end.value_or(bounds.end);
  if (bounds.start > bounds.end) {
    return group_error(FormulonErrorCode::kInvalidArgument, "a date group's start must not follow its end");
  }
  return bounds;
}

Expected<void, Error> check_numeric_item_count(const NumericGroupState& state) {
  if (!(numeric_bucket_count(state) + 2.0 <= static_cast<double>(kMaxFieldItems))) {
    return group_error(FormulonErrorCode::kSecurityResourceLimit, "the numeric group has more items than a field holds",
                       "limit=" + std::to_string(kMaxFieldItems));
  }
  return Expected<void, Error>::Ok();
}

Expected<NumericGroupState, Error> resolve_numeric_group(const PivotCache& cache, std::size_t base,
                                                         std::optional<double> start, std::optional<double> end,
                                                         double interval) {
  if (base >= cache.fields().size()) {
    return group_error(FormulonErrorCode::kIndexOutOfRange, "no such cache field",
                       "field=" + std::to_string(base) + " count=" + std::to_string(cache.fields().size()));
  }
  const PivotCacheField& field = cache.fields()[base];
  if (!std::isfinite(interval) || !(interval > 0.0)) {
    return group_error(FormulonErrorCode::kInvalidArgument, "a numeric group's interval must be finite and positive");
  }
  for (const std::optional<double>& bound : {start, end}) {
    if (bound.has_value() && !std::isfinite(*bound)) {
      return group_error(FormulonErrorCode::kInvalidArgument, "a numeric group's bound must be finite");
    }
  }
  if (!field.is_database_field) {
    return group_error(FormulonErrorCode::kPreconditionFailed, "numeric grouping needs a database field",
                       "field=" + std::to_string(base));
  }
  SerialRange range;
  for (const BaseValue& v : present_base_values(cache, base)) {
    if (!v.value.is_number() || v.is_date) {
      return group_error(FormulonErrorCode::kPreconditionFailed,
                         "only a field holding nothing but numbers can be grouped", "field=" + field.name);
    }
    range.add(v.value.as_number());
  }
  if (!range.any && (!start.has_value() || !end.has_value())) {
    return group_error(FormulonErrorCode::kPreconditionFailed, "the field holds no number to take a bound from",
                       "field=" + field.name);
  }
  NumericGroupState state;
  state.start = start.value_or(range.lo);
  state.end = end.value_or(range.hi);
  state.interval = interval;
  state.auto_start = !start.has_value();
  state.auto_end = !end.has_value();
  if (state.start > state.end) {
    return group_error(FormulonErrorCode::kInvalidArgument, "a numeric group's start must not follow its end");
  }
  RETURN_IF_ERROR(check_numeric_item_count(state));
  return state;
}

}  // namespace detail

namespace {

Value cache_text(PivotCache& cache, std::string text) {
  cache.mutable_text_storage().push_back(std::move(text));
  return Value::text(cache.mutable_text_storage().back());
}

PivotCacheField group_cache_field(const PivotCache& cache, std::size_t base, GroupBy period, ExcelProfile profile) {
  PivotCacheField field{date_group_field_name(period, cache.fields()[base].name, profile), {}};
  field.is_database_field = false;
  field.num_fmt_id = 0U;
  return field;
}

// The group field of one calendar period, labelled as Excel labels it when
// the group is created.
PivotCacheField range_group_field(PivotCache& cache, std::size_t base, GroupBy period, const DateGroupSpec& spec,
                                  const GroupBounds& bounds, bool date1904, ExcelProfile profile) {
  PivotCacheField field = group_cache_field(cache, base, period, profile);
  FieldGroup group;
  group.base = static_cast<std::uint32_t>(base);
  group.kind = GroupKind::Range;
  group.range.by = period;
  group.range.is_date = true;
  group.range.start_date = group_date_iso(bounds.start, date1904);
  group.range.end_date = group_date_iso(bounds.end, date1904);
  group.range.auto_start = !spec.start.has_value();
  group.range.auto_end = !spec.end.has_value();
  const std::uint32_t interval = period == GroupBy::Days ? spec.interval_days : 1U;
  group.range.interval = static_cast<double>(interval);

  const std::string_view short_date = localized_builtin_num_fmt(14U, profile);
  const auto date_text = [&](double serial) {
    return text_format::format_value_for_display(Value::number(std::floor(serial)), short_date, date1904, profile).text;
  };
  std::vector<std::string> labels;
  labels.push_back(date_group_before_label(date_text(bounds.start)));
  const double start_day = std::floor(bounds.start);
  const double end_day = std::floor(bounds.end);
  if (period == GroupBy::Days && interval > 1U) {
    for (double first = start_day; first <= end_day; first += static_cast<double>(interval)) {
      const double last = std::min(first + static_cast<double>(interval) - 1.0, end_day);
      labels.push_back(date_group_interval_label(date_text(first), date_text(last)));
    }
  } else if (period == GroupBy::Years) {
    const int last_year = ymd_of_day(end_day, date1904).y;
    for (int year = ymd_of_day(start_day, date1904).y; year <= last_year; ++year) {
      labels.push_back(date_group_year_label(year, profile));
    }
  } else {
    for (unsigned index = 0;; ++index) {
      std::string label = date_group_unit_label(period, index, profile);
      if (label.empty()) {
        break;
      }
      labels.push_back(std::move(label));
    }
  }
  labels.push_back(date_group_after_label(date_text(bounds.end)));
  group.group_items.reserve(labels.size());
  for (std::string& label : labels) {
    group.group_items.push_back(cache_text(cache, std::move(label)));
  }
  field.field_group = std::move(group);
  return field;
}

// Gives a base field held inline in its records the shared items a discrete
// group maps from: its distinct values in record order, each record then
// naming one by index.
void share_base_items(PivotCache& cache, std::size_t base, bool date1904) {
  std::vector<PivotCacheField>& fields = cache.mutable_fields();
  if (!fields[base].shared_items.empty()) {
    return;
  }
  std::vector<Value> items;
  for (PivotCacheRecord& record : cache.mutable_records()) {
    if (record.cell_is_index.size() < record.cells.size()) {
      // Pin the encodings inferred so far before the base field turns shared.
      const std::size_t known = record.cell_is_index.size();
      record.cell_is_index.resize(record.cells.size(), false);
      for (std::size_t fi = known; fi < record.cells.size() && fi < fields.size(); ++fi) {
        record.cell_is_index[fi] = !fields[fi].shared_items.empty() && record.cells[fi].is_number();
      }
    }
    if (base >= record.cells.size()) {
      continue;
    }
    const Value value = record.cells[base];
    std::size_t index = 0;
    while (index < items.size() && !(items[index] == value)) {
      ++index;
    }
    if (index == items.size()) {
      items.push_back(value);
    }
    record.cells[base] = Value::number(static_cast<double>(index));
    record.cell_is_index[base] = true;
    if (base < record.cell_is_date.size()) {
      record.cell_is_date[base] = false;
    }
  }
  PivotCacheField& field = fields[base];
  SerialRange range;
  for (const Value& item : items) {
    range.add(item.as_number());
  }
  field.shared_item_is_date.assign(items.size(), true);
  field.shared_items = std::move(items);
  SharedItemsHints& h = field.shared_items_hints;
  h = SharedItemsHints{};
  h.has_contains_semi_mixed = true;
  h.has_contains_non_date = true;
  h.has_contains_date = true;
  h.contains_date = true;
  h.has_contains_string = true;
  const GroupBounds bounds = auto_date_group_bounds(range.lo, range.hi, date1904);
  h.has_min_date = true;
  h.min_date = group_date_iso(bounds.start, date1904);
  h.has_max_date = true;
  h.max_date = group_date_iso(bounds.end, date1904);
  h.present = true;
}

// The years group field of the Japanese calendar: the era years the base's
// dates fall in, in date order, as a discrete group.
PivotCacheField era_year_group_field(PivotCache& cache, std::size_t base, bool date1904, ExcelProfile profile) {
  share_base_items(cache, base, date1904);
  PivotCacheField field = group_cache_field(cache, base, GroupBy::Years, profile);
  const std::vector<Value>& items = cache.fields()[base].shared_items;
  struct EraYear {
    double key;
    std::string label;
  };
  std::vector<EraYear> era_years;
  std::vector<double> item_keys(items.size(), 0.0);
  for (std::size_t i = 0; i < items.size(); ++i) {
    std::string label = era_year_label(items[i].as_number(), date1904, &item_keys[i]);
    era_years.push_back({item_keys[i], std::move(label)});
  }
  sort_by_index(era_years, [](const EraYear& a, const EraYear& b) { return a.key < b.key; });
  FieldGroup group;
  group.base = static_cast<std::uint32_t>(base);
  group.kind = GroupKind::Discrete;
  std::vector<double> keys;
  for (EraYear& era_year : era_years) {
    if (!keys.empty() && keys.back() == era_year.key) {
      continue;
    }
    keys.push_back(era_year.key);
    group.group_items.push_back(cache_text(cache, std::move(era_year.label)));
  }
  group.discrete_map.reserve(items.size());
  for (const double key : item_keys) {
    const auto it = std::lower_bound(keys.begin(), keys.end(), key);
    group.discrete_map.push_back(static_cast<std::uint32_t>(it - keys.begin()));
  }
  field.field_group = std::move(group);
  return field;
}

}  // namespace

namespace detail {

Error group_error(FormulonErrorCode code, std::string message, std::string context) {
  Error err = make_error(code, std::move(message));
  err.context = std::move(context);
  return err;
}

Expected<GroupRun, Error> group_run(const PivotCache& cache, std::size_t base) {
  const std::vector<PivotCacheField>& fields = cache.fields();
  GroupRun run{fields.size(), 0U};
  for (std::size_t fi = 0; fi < fields.size(); ++fi) {
    if (is_group_field(fields[fi]) && *fields[fi].field_group->base == base) {
      run.first = std::min(run.first, fi);
      ++run.count;
    }
  }
  if (run.count != 0U && run.first + run.count != fields.size()) {
    return group_error(FormulonErrorCode::kPreconditionFailed,
                       "the group fields of a field can be removed only while they are the last cache fields",
                       "base=" + std::to_string(base));
  }
  return run;
}

}  // namespace detail

std::size_t group_base_index(const PivotCache& cache, std::size_t field) noexcept {
  if (field < cache.fields().size() && is_group_field(cache.fields()[field])) {
    return *cache.fields()[field].field_group->base;
  }
  return field;
}

std::uint32_t date_group_periods(const PivotCache& cache, std::size_t base, bool date1904) {
  std::uint32_t mask = 0;
  for (const PivotCacheField& field : cache.fields()) {
    if (!is_group_field(field) || *field.field_group->base != base) {
      continue;
    }
    const FieldGroup& group = *field.field_group;
    if (group.kind == GroupKind::Range && group.range.by != GroupBy::Range) {
      mask |= date_period_bit(group.range.by);
    } else if (is_era_year_group(cache, field, date1904)) {
      mask |= date_period_bit(GroupBy::Years);
    }
  }
  return mask;
}

CalendarSystem date_group_calendar(const PivotCache& cache, std::size_t base, bool date1904) {
  for (const PivotCacheField& field : cache.fields()) {
    if (is_group_field(field) && *field.field_group->base == base && is_era_year_group(cache, field, date1904)) {
      return CalendarSystem::Japanese;
    }
  }
  return CalendarSystem::Gregorian;
}

GroupBounds auto_date_group_bounds(double min_serial, double max_serial, bool date1904) noexcept {
  const double day = std::floor(max_serial);
  const double end = max_serial > day ? max_serial : std::min(day + 1.0, last_valid_date_serial(date1904));
  return GroupBounds{min_serial, end};
}

std::string group_date_iso(double serial, bool date1904) {
  double day = std::floor(serial);
  long long seconds = std::llround((serial - day) * static_cast<double>(date_time::kSecondsPerDay));
  if (seconds >= date_time::kSecondsPerDay) {
    day += 1.0;
    seconds -= date_time::kSecondsPerDay;
  }
  // The last representable instant is 9999-12-31T23:59:59.
  if (day > last_valid_date_serial(date1904)) {
    day = last_valid_date_serial(date1904);
    seconds = date_time::kSecondsPerDay - 1;
  }
  // The inverse of `group_date_serial`: serials through 60 name the day before.
  const bool before_phantom = !date1904 && day < kFirstSerialAfterPhantomDay;
  const date_time::YMD ymd = date_time::ymd_from_serial(before_phantom ? day - 1.0 : day, date1904);
  std::string text;
  append_padded(text, ymd.y, 4);
  text.push_back('-');
  append_padded(text, ymd.m, 2);
  text.push_back('-');
  append_padded(text, ymd.d, 2);
  text.push_back('T');
  append_padded(text, seconds / 3600, 2);
  text.push_back(':');
  append_padded(text, seconds / 60 % 60, 2);
  text.push_back(':');
  append_padded(text, seconds % 60, 2);
  return text;
}

Expected<std::size_t, Error> group_cache_by_date(PivotCache& cache, std::size_t base, const DateGroupSpec& spec,
                                                 bool date1904, ExcelProfile profile) {
  ASSIGN_OR_RETURN(GroupBounds bounds, resolve_date_group_bounds(cache, base, spec, date1904));
  const std::optional<FieldGroup>& existing = cache.fields()[base].field_group;
  if (existing.has_value() && existing->par.has_value()) {
    return group_error(FormulonErrorCode::kPreconditionFailed, "the field is grouped already",
                       "field=" + cache.fields()[base].name);
  }
  const std::size_t first = cache.fields().size();
  for (const GroupBy period : kDatePeriods) {
    if ((spec.periods & date_period_bit(period)) == 0U) {
      continue;
    }
    PivotCacheField field = period == GroupBy::Years && spec.calendar == CalendarSystem::Japanese
                                ? era_year_group_field(cache, base, date1904, profile)
                                : range_group_field(cache, base, period, spec, bounds, date1904, profile);
    cache.mutable_fields().push_back(std::move(field));
  }
  std::optional<FieldGroup>& base_group = cache.mutable_fields()[base].field_group;
  if (!base_group.has_value()) {
    base_group.emplace();
  }
  base_group->par = static_cast<std::uint32_t>(cache.fields().size() - 1U);
  materialize_group_fields(cache, date1904);
  return first;
}

Expected<std::size_t, Error> ungroup_cache_field(PivotCache& cache, std::size_t base) {
  ASSIGN_OR_RETURN(GroupRun run, group_run(cache, base));
  const std::size_t first = run.first;
  if (run.count != 0U) {
    cache.mutable_fields().resize(first);
    for (PivotCacheRecord& record : cache.mutable_records()) {
      record.cells.resize(std::min(record.cells.size(), first), Value::blank());
      record.cell_is_index.resize(std::min(record.cell_is_index.size(), first));
      record.cell_is_date.resize(std::min(record.cell_is_date.size(), first));
      record.cell_text_slot.resize(std::min(record.cell_text_slot.size(), first));
    }
  }
  if (base < cache.fields().size()) {
    std::optional<FieldGroup>& group = cache.mutable_fields()[base].field_group;
    if (group.has_value()) {
      group->par.reset();
      if (!group->base.has_value() && group->kind == GroupKind::None && group->group_items.empty() &&
          group->extra_attrs.empty()) {
        group.reset();
      }
    }
  }
  return run.count;
}

namespace detail {

Expected<void, Error> set_numeric_range(PivotCache& cache, FieldGroup& group, const NumericGroupState& state,
                                        ExcelProfile profile) {
  RETURN_IF_ERROR(check_numeric_item_count(state));
  const auto buckets = static_cast<std::size_t>(numeric_bucket_count(state));
  std::vector<Value> items;
  items.reserve(buckets + 2U);
  const char separator = locale_facts(profile).decimal_separator;
  const bool whole = whole_settings(state);
  std::string lower = label_number(state.start, separator);
  items.push_back(cache_text(cache, "<" + lower));
  for (std::size_t n = 0; n < buckets; ++n) {
    const double a = numeric_edge(state, n);
    const double b = numeric_edge(state, n + 1U);
    std::string upper = label_number(b, separator);
    if (whole && a > -kMaxIntegerLabelEdge && b <= kMaxIntegerLabelEdge) {
      // The last range ends at the end when the end is its upper edge (`110-120`).
      const double shown = n + 1U == buckets && b == state.end ? b : b - 1.0;
      items.push_back(cache_text(cache, lower + "-" + label_number(shown, separator)));
    } else {
      items.push_back(cache_text(cache, lower + "-" + upper));
    }
    lower = std::move(upper);
  }
  items.push_back(cache_text(cache, ">" + lower));
  group.kind = GroupKind::Range;
  group.range = RangeGroup{};
  group.range.start_num = state.start;
  group.range.end_num = state.end;
  group.range.interval = state.interval;
  group.range.auto_start = state.auto_start;
  group.range.auto_end = state.auto_end;
  group.discrete_map.clear();
  group.group_items = std::move(items);
  group.group_item_is_date.clear();
  return Expected<void, Error>::Ok();
}

Expected<std::size_t, Error> group_cache_by_number(PivotCache& cache, std::size_t base, const NumericGroupState& state,
                                                   bool date1904, ExcelProfile profile) {
  const std::optional<FieldGroup>& existing = cache.fields()[base].field_group;
  if (existing.has_value() && existing->par.has_value()) {
    return group_error(FormulonErrorCode::kPreconditionFailed, "the field is grouped already",
                       "field=" + cache.fields()[base].name);
  }
  const std::string& base_name = cache.fields()[base].name;
  const auto taken = [&cache](std::string_view name) {
    return std::any_of(cache.fields().begin(), cache.fields().end(),
                       [name](const PivotCacheField& other) { return same_header(other.name, name); });
  };
  std::string name = base_name + "2";
  for (std::size_t n = 3; taken(name); ++n) {
    name = base_name + std::to_string(n);
  }
  PivotCacheField field{std::move(name), {}};
  field.is_database_field = false;
  field.num_fmt_id = 0U;
  FieldGroup group;
  group.base = static_cast<std::uint32_t>(base);
  RETURN_IF_ERROR(set_numeric_range(cache, group, state, profile));
  field.field_group = std::move(group);
  const std::size_t index = cache.fields().size();
  cache.mutable_fields().push_back(std::move(field));
  std::optional<FieldGroup>& base_group = cache.mutable_fields()[base].field_group;
  if (!base_group.has_value()) {
    base_group.emplace();
  }
  base_group->par = static_cast<std::uint32_t>(index);
  materialize_group_fields(cache, date1904);
  return index;
}

}  // namespace detail

std::optional<DateGroupState> date_group_state(const PivotCache& cache, std::size_t field, bool date1904) {
  const std::size_t base = group_base_index(cache, field);
  DateGroupState state;
  state.periods = date_group_periods(cache, base, date1904);
  if (state.periods == 0U) {
    return std::nullopt;
  }
  state.calendar = date_group_calendar(cache, base, date1904);
  RangeGroup range;
  range.is_date = true;
  for (const PivotCacheField& candidate : cache.fields()) {
    if (!is_group_field(candidate) || *candidate.field_group->base != base ||
        candidate.field_group->kind != GroupKind::Range || candidate.field_group->range.by == GroupBy::Range) {
      continue;
    }
    range = candidate.field_group->range;
    if (range.by == GroupBy::Days) {
      break;
    }
  }
  if (range.by == GroupBy::Days && range.interval >= 1.0) {
    state.interval_days = static_cast<std::uint32_t>(range.interval);
  }
  // An era-year grouping alone states no bounds and follows the data.
  const GroupBounds bounds = resolve_bounds(cache, base, range, date1904);
  state.start = bounds.start;
  state.end = bounds.end;
  state.auto_start = range.auto_start;
  state.auto_end = range.auto_end;
  return state;
}

std::optional<NumericGroupState> numeric_group_state(const PivotCache& cache, std::size_t field) {
  const std::size_t base = group_base_index(cache, field);
  for (const PivotCacheField& candidate : cache.fields()) {
    if (!is_group_field(candidate) || *candidate.field_group->base != base) {
      continue;
    }
    const FieldGroup& group = *candidate.field_group;
    if (group.kind != GroupKind::Range || group.range.is_date || group.range.by != GroupBy::Range) {
      continue;
    }
    NumericGroupState state;
    state.start = group.range.start_num;
    state.end = group.range.end_num;
    state.interval = group.range.interval;
    state.auto_start = group.range.auto_start;
    state.auto_end = group.range.auto_end;
    return state;
  }
  return std::nullopt;
}

}  // namespace formulon::pivot
