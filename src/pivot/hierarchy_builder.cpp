
#include "pivot/hierarchy_builder.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "eval/date_time.h"
#include "eval/japanese_era.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// ---------------------------------------------------------------------------
// Date grouping.
// ---------------------------------------------------------------------------
//
// When a `PivotField` declares a `date_group`, the field's source `Value`
// (which the cache stores as a date serial) is mapped to a bucket
// `(sort_key, label)` pair. The sort key feeds the hierarchy's
// `std::map<Value, HierNode, ValueLess>` so leaves order chronologically
// (e.g. 2023 < 2024 < 2025); the label is what the renderer / GETPIVOTDATA
// see as the bucket name.
//
// MVP scope: Year, Quarter, Month, Day for both Gregorian and Japanese
// calendars; Week / Hour / Minute / Second pass through unchanged
// (deferred — Excel's calendar-week rules and time-of-day grouping
// require additional oracle authoring). For Japanese-calendar Year /
// Month, the era prefix is the full kanji name (令和 / 平成 / ...);
// pre-Meiji dates classify as Meiji, mirroring Excel's lenient
// behaviour.

struct DateBucket {
  Value sort_key;     ///< Numeric sort handle (chronological order).
  std::string label;  ///< Display string for the bucket.
};

// Two-digit zero-padded decimal append (no <iomanip>).
void append_pad2(std::string& out, unsigned n) {
  if (n < 10u) {
    out.push_back('0');
  }
  out.append(std::to_string(n));
}

// Four-digit (or wider) zero-padded year append. Negative years emit a
// leading `-`; the resulting label is sortable lexicographically only
// within the same sign, which is fine because the sort key is
// independent of the label.
void append_year(std::string& out, int y) {
  if (y < 0) {
    out.push_back('-');
    y = -y;
  }
  if (y < 10) {
    out.append("000");
  } else if (y < 100) {
    out.append("00");
  } else if (y < 1000) {
    out.push_back('0');
  }
  out.append(std::to_string(y));
}

DateBucket bucket_date(double serial, const PivotDateGroup& dg) {
  using formulon::eval::date_time::civil_from_days;
  using formulon::eval::date_time::days_from_civil;
  using formulon::eval::date_time::HMS;
  using formulon::eval::date_time::hms_from_fraction;
  using formulon::eval::date_time::YMD;
  using formulon::eval::date_time::ymd_from_serial;
  using formulon::eval::japanese_era::classify_era;
  using formulon::eval::japanese_era::EraInfo;

  // Negative / non-finite serials are not valid Excel dates; pass them
  // through so the existing display path renders the raw number.
  if (!(serial >= 0.0)) {
    Value raw = Value::number(serial);
    return {raw, display_string(raw)};
  }
  const double serial_floor = std::floor(serial);
  const YMD ymd = ymd_from_serial(serial_floor);

  switch (dg.granularity) {
    case DateGrouping::Year: {
      if (dg.calendar == CalendarSystem::Japanese) {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        const int era_year = ymd.y - era.year_anchor + 1;
        const double key = static_cast<double>(era.year_anchor) * 10000.0 + static_cast<double>(era_year);
        std::string label = era.kanji2;
        label.append(std::to_string(era_year));
        // 年
        label.append("\xE5\xB9\xB4");
        return {Value::number(key), std::move(label)};
      }
      const double key = static_cast<double>(ymd.y);
      std::string label;
      append_year(label, ymd.y);
      return {Value::number(key), std::move(label)};
    }
    case DateGrouping::Quarter: {
      const int quarter = (static_cast<int>(ymd.m) - 1) / 3 + 1;
      const double key = static_cast<double>(ymd.y) * 4.0 + static_cast<double>(quarter - 1);
      std::string label;
      append_year(label, ymd.y);
      label.append("-Q");
      label.append(std::to_string(quarter));
      return {Value::number(key), std::move(label)};
    }
    case DateGrouping::Month: {
      const double key = static_cast<double>(ymd.y) * 100.0 + static_cast<double>(ymd.m);
      std::string label;
      if (dg.calendar == CalendarSystem::Japanese) {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        const int era_year = ymd.y - era.year_anchor + 1;
        label = era.kanji2;
        label.append(std::to_string(era_year));
        label.append("\xE5\xB9\xB4");  // 年
        label.append(std::to_string(ymd.m));
        label.append("\xE6\x9C\x88");  // 月
      } else {
        append_year(label, ymd.y);
        label.push_back('-');
        append_pad2(label, ymd.m);
      }
      return {Value::number(key), std::move(label)};
    }
    case DateGrouping::Day: {
      // Use the floor serial as the sort key; it is already
      // chronological. Label is `YYYY-MM-DD` (Gregorian) regardless of
      // the calendar — Excel renders day-level buckets in the workbook
      // locale's short-date pattern, which for ja-JP and en-US alike
      // collapses to a numeric YMD.
      std::string label;
      append_year(label, ymd.y);
      label.push_back('-');
      append_pad2(label, ymd.m);
      label.push_back('-');
      append_pad2(label, ymd.d);
      return {Value::number(serial_floor), std::move(label)};
    }
    case DateGrouping::Week: {
      // Sunday-start week. 1970-01-01 was a Thursday, so with Sun=0..Sat=6
      // the day-of-week is `((days % 7) + 7 + 4) % 7`. Subtract that to
      // land on the week's Sunday. The civil-day count is monotone across
      // calendar years, so it doubles as the chronological sort key.
      // Excel ja-JP renders weekly buckets as Gregorian YYYY-MM-DD even
      // when the field's date group nominally targets the Japanese
      // calendar; mirror that by ignoring `dg.calendar` here.
      const std::int64_t days = days_from_civil(ymd.y, ymd.m, ymd.d);
      const std::int64_t dow = ((days % 7) + 7 + 4) % 7;
      const std::int64_t week_start_days = days - dow;
      const YMD ws = civil_from_days(week_start_days);
      std::string label;
      append_year(label, ws.y);
      label.push_back('-');
      append_pad2(label, ws.m);
      label.push_back('-');
      append_pad2(label, ws.d);
      return {Value::number(static_cast<double>(week_start_days)), std::move(label)};
    }
    case DateGrouping::Hour: {
      const double hour_index = std::floor(serial * 24.0);
      const HMS hms = hms_from_fraction(serial);
      std::string label;
      append_year(label, ymd.y);
      label.push_back('-');
      append_pad2(label, ymd.m);
      label.push_back('-');
      append_pad2(label, ymd.d);
      label.push_back(' ');
      append_pad2(label, hms.h);
      return {Value::number(hour_index), std::move(label)};
    }
    case DateGrouping::Minute: {
      const double minute_index = std::floor(serial * 1440.0);
      const HMS hms = hms_from_fraction(serial);
      std::string label;
      append_year(label, ymd.y);
      label.push_back('-');
      append_pad2(label, ymd.m);
      label.push_back('-');
      append_pad2(label, ymd.d);
      label.push_back(' ');
      append_pad2(label, hms.h);
      label.push_back(':');
      append_pad2(label, hms.m);
      return {Value::number(minute_index), std::move(label)};
    }
    case DateGrouping::Second: {
      const double second_index = std::floor(serial * 86400.0);
      const HMS hms = hms_from_fraction(serial);
      std::string label;
      append_year(label, ymd.y);
      label.push_back('-');
      append_pad2(label, ymd.m);
      label.push_back('-');
      append_pad2(label, ymd.d);
      label.push_back(' ');
      append_pad2(label, hms.h);
      label.push_back(':');
      append_pad2(label, hms.m);
      label.push_back(':');
      append_pad2(label, hms.s);
      return {Value::number(second_index), std::move(label)};
    }
  }
  Value raw = Value::number(serial);
  return {raw, display_string(raw)};
}

}  // namespace

HierNode* insert_path(const PivotCache& cache, const std::vector<HierLevel>& levels, const PivotCacheRecord& record,
                      std::size_t record_index, HierNode& root) {
  HierNode* cursor = &root;
  for (const HierLevel& level : levels) {
    const Value raw = cell_value(cache, record, level.field_index);
    Value key = raw;
    std::string label_override;
    if (level.date_group != nullptr && raw.is_number()) {
      DateBucket bucket = bucket_date(raw.as_number(), *level.date_group);
      key = bucket.sort_key;
      label_override = std::move(bucket.label);
    }
    auto [it, inserted] = cursor->children.emplace(key, HierNode{});
    if (inserted && !label_override.empty()) {
      it->second.label_override = std::move(label_override);
    }
    cursor = &it->second;
    cursor->record_indices.push_back(record_index);
  }
  return cursor;
}

std::string node_label(const Value& key, const HierNode& child) {
  if (!child.label_override.empty()) {
    return child.label_override;
  }
  return display_string(key);
}

}  // namespace formulon::pivot
