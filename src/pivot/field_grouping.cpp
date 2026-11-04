#include "pivot/field_grouping.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/text_format/display_text.h"
#include "num_fmt_locale.h"
#include "pivot/date_serial.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "sheet.h"
#include "utils/checked_index.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/iso_date.h"
#include "utils/japanese_era.h"
#include "value.h"
#include "workbook.h"

namespace formulon::pivot {
namespace {

// Days before each month in a leap year: a days group numbers 2/29 as day 60
// in every year, so 3/1 is day 61 whether or not the year has a 2/29.
constexpr unsigned kLeapDaysBeforeMonth[12] = {0U, 31U, 60U, 91U, 121U, 152U, 182U, 213U, 244U, 274U, 305U, 335U};

// 1900-03-01, the first serial past Excel's phantom 1900-02-29 (serial 60).
constexpr double kFirstSerialAfterPhantomDay = 61.0;

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
  bool have_data = false;
  double data_min = 0.0;
  double data_max = 0.0;
  for (const PivotCacheRecord& record : cache.records()) {
    const Value v = cell_value(cache, record, base);
    if (!v.is_number() || !is_valid_date_serial(v.as_number(), date1904)) {
      continue;
    }
    data_min = have_data ? std::min(data_min, v.as_number()) : v.as_number();
    data_max = have_data ? std::max(data_max, v.as_number()) : v.as_number();
    have_data = true;
  }
  const GroupBounds data = have_data ? auto_date_group_bounds(data_min, data_max, date1904) : GroupBounds{};
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

namespace {

Error group_error(FormulonErrorCode code, std::string message, std::string context = {}) {
  Error err = make_error(code, std::move(message));
  err.context = std::move(context);
  return err;
}

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

// Group items over `base` in index order, and whether they are the trailing
// fields of the cache.
struct GroupRun {
  std::size_t first = 0;
  std::size_t count = 0;
};

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
                       "the group fields of a date field can be removed only while they are the last cache fields",
                       "base=" + std::to_string(base));
  }
  return run;
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

// Checks `spec` against the base data and resolves the bounds the group
// fields take.
Expected<GroupBounds, Error> resolve_date_group_bounds(const PivotCache& cache, std::size_t base,
                                                       const DateGroupSpec& spec, bool date1904) {
  if (base >= cache.fields().size() || !cache.fields()[base].is_database_field) {
    return group_error(FormulonErrorCode::kInvalidArgument, "date grouping needs a database field",
                       "field=" + std::to_string(base));
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
  // A field read from a file states its types; a value it did not spell as a
  // date is a number.
  const bool typed = field.shared_items_hints.present;
  bool have_value = false;
  double lo = 0.0;
  double hi = 0.0;
  for (const BaseValue& v : base_values(cache, base)) {
    if (!v.value.is_number() || !is_valid_date_serial(v.value.as_number(), date1904) || (typed && !v.is_date)) {
      return group_error(FormulonErrorCode::kInvalidArgument, "only a field holding nothing but dates can be grouped",
                         "field=" + field.name);
    }
    lo = have_value ? std::min(lo, v.value.as_number()) : v.value.as_number();
    hi = have_value ? std::max(hi, v.value.as_number()) : v.value.as_number();
    have_value = true;
  }
  if (!have_value) {
    return group_error(FormulonErrorCode::kInvalidArgument, "the field holds no date to group", "field=" + field.name);
  }
  GroupBounds bounds = auto_date_group_bounds(lo, hi, date1904);
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
  double lo = 0.0;
  double hi = 0.0;
  for (std::size_t i = 0; i < items.size(); ++i) {
    lo = i == 0U ? items[i].as_number() : std::min(lo, items[i].as_number());
    hi = i == 0U ? items[i].as_number() : std::max(hi, items[i].as_number());
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
  const GroupBounds bounds = auto_date_group_bounds(lo, hi, date1904);
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

// A row or column axis in document order, the Values marker included.
std::vector<std::int64_t> axis_entries(const PivotTable& table, bool row) {
  const std::vector<std::uint32_t>& order = row ? table.row_field_order() : table.col_field_order();
  const std::optional<std::size_t> values = row ? table.row_values_position() : table.col_values_position();
  std::vector<std::int64_t> entries(order.begin(), order.end());
  if (values.has_value()) {
    entries.insert(entries.begin() + static_cast<std::ptrdiff_t>(std::min(*values, entries.size())),
                   kValuesFieldPosition);
  }
  return entries;
}

void set_axis_entries(PivotTable& table, bool row, const std::vector<std::int64_t>& entries) {
  std::vector<std::uint32_t> order;
  std::optional<std::size_t> values;
  for (std::size_t i = 0; i < entries.size(); ++i) {
    if (entries[i] == kValuesFieldPosition) {
      values = i;
    } else {
      order.push_back(static_cast<std::uint32_t>(entries[i]));
    }
  }
  if (row) {
    table.mutable_row_field_order() = std::move(order);
    table.set_row_values_position(values);
  } else {
    table.mutable_col_field_order() = std::move(order);
    table.set_col_values_position(values);
  }
}

PivotAxis axis_of(bool row) {
  return row ? PivotAxis::Row : PivotAxis::Col;
}

// Gives `table` a pivot field for each cache field from `first` on and puts
// them on the axis that holds the base: in its place on the grouped pivot,
// before it on the others.
void add_group_pivot_fields(PivotTable& table, const PivotCache& cache, std::size_t base, std::size_t first,
                            bool replace_base) {
  std::vector<PivotField>& fields = table.mutable_fields();
  while (fields.size() < cache.fields().size()) {
    const std::size_t fi = fields.size();
    PivotField field;
    field.source_name = cache.fields()[fi].name;
    field.axis = PivotAxis::None;
    if (fi >= first) {
      // Excel lists group items in group order, with no sort stated.
      field.sort.manual = true;
      const std::size_t item_count = cache.fields()[fi].field_group->group_items.size();
      for (std::size_t i = 0; i < item_count; ++i) {
        PivotItem item;
        item.has_cache_index = true;
        item.cache_index = static_cast<std::uint32_t>(i);
        field.items.push_back(std::move(item));
      }
    }
    fields.push_back(std::move(field));
  }
  for (const bool row : {true, false}) {
    std::vector<std::int64_t> entries = axis_entries(table, row);
    const auto at = std::find(entries.begin(), entries.end(), static_cast<std::int64_t>(base));
    if (at == entries.end()) {
      continue;
    }
    // Coarsest outermost.
    std::vector<std::int64_t> groups;
    for (std::size_t fi = cache.fields().size(); fi > first; --fi) {
      groups.push_back(static_cast<std::int64_t>(fi - 1U));
      fields[fi - 1U].axis = axis_of(row);
    }
    const auto pos = entries.erase(at, replace_base ? at + 1 : at);
    entries.insert(pos, groups.begin(), groups.end());
    if (replace_base) {
      fields[base].axis = PivotAxis::None;
    }
    set_axis_entries(table, row, entries);
  }
}

// Drops `table`'s pivot fields from `first` on; the base takes the axis
// position the first of them held unless it is on an axis already.
void remove_group_pivot_fields(PivotTable& table, std::size_t base, std::size_t first) {
  std::vector<PivotField>& fields = table.mutable_fields();
  if (fields.size() <= first) {
    return;
  }
  for (const bool row : {true, false}) {
    std::vector<std::int64_t> entries = axis_entries(table, row);
    std::vector<std::int64_t> kept;
    for (const std::int64_t entry : entries) {
      if (entry < 0 || static_cast<std::size_t>(entry) < first) {
        kept.push_back(entry);
        continue;
      }
      if (base < fields.size() && fields[base].axis == PivotAxis::None) {
        kept.push_back(static_cast<std::int64_t>(base));
        fields[base].axis = axis_of(row);
      }
    }
    if (kept != entries) {
      set_axis_entries(table, row, kept);
    }
  }
  std::vector<PivotPageField>& pages = table.mutable_page_fields();
  pages.erase(std::remove_if(pages.begin(), pages.end(),
                             [first](const PivotPageField& page) { return page.field_index >= first; }),
              pages.end());
  std::vector<PivotDataField>& data = table.mutable_data_fields();
  data.erase(
      std::remove_if(data.begin(), data.end(), [first](const PivotDataField& df) { return df.field_index >= first; }),
      data.end());
  table.erase_active_filters_if(
      [first](const PivotFilter& filter) { return filter.source.has_value() && filter.source->field_index >= first; });
  fields.resize(first);
}

void invalidate_tables(Workbook& wb, std::uint32_t cache_id) {
  for (std::size_t si = 0; si < wb.sheet_count(); ++si) {
    for (std::unique_ptr<PivotTable>& table : wb.sheet(si).mutable_pivot_tables()) {
      if (table != nullptr && table->pivot_cache_id() == cache_id) {
        table->clear_last_result();
        table->clear_span_authored();
        table->mark_layout_dirty();
      }
    }
  }
  wb.mark_all_formulas_dirty();
}

// The pivot `(sheet, pivot)`, its cache and the base of its field `field`.
struct GroupTarget {
  PivotTable* table = nullptr;
  PivotCache* cache = nullptr;
  std::size_t base = 0;
};

Expected<GroupTarget, Error> resolve_group_target(Workbook& wb, std::size_t sheet, std::size_t pivot,
                                                  std::size_t field) {
  const std::string where =
      "sheet=" + std::to_string(sheet) + " pivot=" + std::to_string(pivot) + " field=" + std::to_string(field);
  if (sheet >= wb.sheet_count() || pivot >= wb.sheet(sheet).pivot_tables().size() ||
      wb.sheet(sheet).pivot_tables()[pivot] == nullptr) {
    return group_error(FormulonErrorCode::kInvalidArgument, "no such pivot table", where);
  }
  GroupTarget target;
  target.table = wb.sheet(sheet).mutable_pivot_tables()[pivot].get();
  for (std::unique_ptr<PivotCache>& cache : wb.mutable_pivot_caches()) {
    if (cache != nullptr && cache->cache_id() == target.table->pivot_cache_id()) {
      target.cache = cache.get();
    }
  }
  if (target.cache == nullptr || field >= target.table->fields().size() || field >= target.cache->fields().size()) {
    return group_error(FormulonErrorCode::kInvalidArgument, "no such pivot field", where);
  }
  target.base = group_base_index(*target.cache, field);
  return target;
}

// Takes the group fields over `base` off every pivot on `cache` and out of
// the cache; `run` is their checked position.
void remove_date_group(Workbook& wb, PivotCache& cache, std::size_t base, const GroupRun& run) {
  for (std::size_t si = 0; si < wb.sheet_count(); ++si) {
    for (std::unique_ptr<PivotTable>& table : wb.sheet(si).mutable_pivot_tables()) {
      if (table != nullptr && table->pivot_cache_id() == cache.cache_id()) {
        remove_group_pivot_fields(*table, base, run.first);
      }
    }
  }
  (void)ungroup_cache_field(cache, base);
}

}  // namespace

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
  // The inverse of `group_date_serial`: serials through 60 name the day before.
  const bool before_phantom = !date1904 && day < kFirstSerialAfterPhantomDay;
  const date_time::YMD ymd = date_time::ymd_from_serial(before_phantom ? day - 1.0 : day, date1904);
  const auto two = [](std::string& out, long long v) {
    out.push_back(static_cast<char>('0' + (v / 10) % 10));
    out.push_back(static_cast<char>('0' + v % 10));
  };
  std::string text = std::to_string(ymd.y);
  while (text.size() < 4U) {
    text.insert(text.begin(), '0');
  }
  text.push_back('-');
  two(text, ymd.m);
  text.push_back('-');
  two(text, ymd.d);
  text.push_back('T');
  two(text, seconds / 3600);
  text.push_back(':');
  two(text, seconds / 60 % 60);
  text.push_back(':');
  two(text, seconds % 60);
  return text;
}

Expected<std::size_t, Error> group_cache_by_date(PivotCache& cache, std::size_t base, const DateGroupSpec& spec,
                                                 bool date1904, ExcelProfile profile) {
  Expected<GroupBounds, Error> bounds = resolve_date_group_bounds(cache, base, spec, date1904);
  if (!bounds) {
    return bounds.error();
  }
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
                                : range_group_field(cache, base, period, spec, bounds.value(), date1904, profile);
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
  Expected<GroupRun, Error> run = group_run(cache, base);
  if (!run) {
    return run.error();
  }
  const std::size_t first = run.value().first;
  if (run.value().count != 0U) {
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
  return run.value().count;
}

Expected<void, Error> set_date_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field,
                                     std::uint32_t periods, CalendarSystem calendar, std::uint32_t interval_days,
                                     std::optional<double> start, std::optional<double> end) {
  Expected<GroupTarget, Error> target = resolve_group_target(wb, sheet, pivot, field);
  if (!target) {
    return target.error();
  }
  PivotCache& cache = *target.value().cache;
  const std::size_t base = target.value().base;
  const DateGroupSpec spec{periods, calendar, interval_days, start, end};
  if (Expected<GroupBounds, Error> checked = resolve_date_group_bounds(cache, base, spec, wb.date1904()); !checked) {
    return checked.error();
  }
  Expected<GroupRun, Error> run = group_run(cache, base);
  if (!run) {
    return run.error();
  }
  remove_date_group(wb, cache, base, run.value());
  Expected<std::size_t, Error> first = group_cache_by_date(cache, base, spec, wb.date1904(), wb.excel_profile());
  if (!first) {
    return first.error();
  }
  for (std::size_t si = 0; si < wb.sheet_count(); ++si) {
    for (std::unique_ptr<PivotTable>& table : wb.sheet(si).mutable_pivot_tables()) {
      if (table != nullptr && table->pivot_cache_id() == cache.cache_id()) {
        add_group_pivot_fields(*table, cache, base, first.value(), table.get() == target.value().table);
        resolve_pivot_names(*table, cache);
      }
    }
  }
  invalidate_tables(wb, cache.cache_id());
  return Expected<void, Error>::Ok();
}

Expected<void, Error> clear_date_group(Workbook& wb, std::size_t sheet, std::size_t pivot, std::size_t field) {
  Expected<GroupTarget, Error> target = resolve_group_target(wb, sheet, pivot, field);
  if (!target) {
    return target.error();
  }
  PivotCache& cache = *target.value().cache;
  Expected<GroupRun, Error> run = group_run(cache, target.value().base);
  if (!run) {
    return run.error();
  }
  if (run.value().count == 0U) {
    return Expected<void, Error>::Ok();
  }
  remove_date_group(wb, cache, target.value().base, run.value());
  invalidate_tables(wb, cache.cache_id());
  return Expected<void, Error>::Ok();
}

}  // namespace formulon::pivot
