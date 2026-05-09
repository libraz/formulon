// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Pivot evaluator implementation. See header / §15.1.3 of the design
// corpus for the algorithm overview. The MVP path implemented here
// produces enough of a `PivotResult` that GETPIVOTDATA can resolve
// label/data tuples against the freshest evaluation snapshot.

#include "pivot/pivot_evaluator.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

#include "eval/date_time.h"
#include "eval/japanese_era.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// ---------------------------------------------------------------------------
// Cross-kind ordering for hierarchy keys.
// ---------------------------------------------------------------------------
//
// Excel's default sort on a row/column field is ascending by the field's
// display value. The cache stores discrete values as the typed `Value`
// they came from, so we need a total order that:
//
//   * groups same-kind values together (so we don't interleave numbers
//     into strings);
//   * sorts numerically within `Number`;
//   * sorts case-sensitively within `Text` (good enough for MVP — the
//     primary oracle is ja-JP which Excel collates byte-wise on the
//     wire; the locale-aware compare will arrive with date grouping);
//   * is total across kinds so `std::map` stays well-formed.
//
// Kind ordering: Number < Bool < Text < Error < everything else. The
// cache only ever produces the first four for discrete fields, but the
// fall-through keeps the comparator total in case a future cache field
// surfaces an Array or similar.
int kind_rank(ValueKind k) noexcept {
  switch (k) {
    case ValueKind::Number:
      return 0;
    case ValueKind::Bool:
      return 1;
    case ValueKind::Text:
      return 2;
    case ValueKind::Error:
      return 3;
    case ValueKind::Blank:
      return 4;
    case ValueKind::Array:
      return 5;
    case ValueKind::Ref:
      return 6;
    case ValueKind::Lambda:
      return 7;
  }
  return 8;
}

bool value_less(const Value& a, const Value& b) noexcept {
  const int ra = kind_rank(a.kind());
  const int rb = kind_rank(b.kind());
  if (ra != rb) {
    return ra < rb;
  }
  switch (a.kind()) {
    case ValueKind::Number:
      return a.as_number() < b.as_number();
    case ValueKind::Bool:
      // false < true.
      return !a.as_boolean() && b.as_boolean();
    case ValueKind::Text:
      return a.as_text() < b.as_text();
    case ValueKind::Error:
      return static_cast<std::uint16_t>(a.as_error()) < static_cast<std::uint16_t>(b.as_error());
    default:
      // Same kind, no in-kind ordering defined — treat as equal.
      return false;
  }
}

struct ValueLess {
  bool operator()(const Value& a, const Value& b) const noexcept { return value_less(a, b); }
};

// ---------------------------------------------------------------------------
// Display / equality helpers for filter visibility.
// ---------------------------------------------------------------------------

// Renders a cache `Value` to the same string that would appear in a
// `PivotItem::name`. Mirrors what the OOXML reader would have written
// out: numbers as their canonical decimal, bools as TRUE/FALSE, errors
// as their `#…` token, blanks as the empty string. This is sufficient
// for matching against `PivotItem::name`, which is the only place we
// use it (manual-filter visibility check).
std::string display_string(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Blank:
      return std::string{};
    case ValueKind::Number: {
      // Excel renders integers without a trailing `.0`. We don't go
      // through the full number-format pipeline here: pivot item names
      // are produced by the OOXML reader from cache `<n v="…"/>` /
      // `<s v="…"/>` literals, and matching on the textual form is
      // robust enough for MVP. A more faithful renderer can replace
      // this when item-level filter parity is required.
      const double d = v.as_number();
      const auto i = static_cast<long long>(d);
      if (static_cast<double>(i) == d) {
        return std::to_string(i);
      }
      return std::to_string(d);
    }
    case ValueKind::Bool:
      return v.as_boolean() ? "TRUE" : "FALSE";
    case ValueKind::Text:
      return std::string{v.as_text()};
    case ValueKind::Error:
      return std::string{display_name(v.as_error())};
    default:
      return std::string{};
  }
}

// ---------------------------------------------------------------------------
// Cache-record value extraction.
// ---------------------------------------------------------------------------

// Pulls the effective `Value` for `(record, field)`. When the field is
// shared (`shared_items` non-empty), the record stores a `Number` index
// into `shared_items`; otherwise the record stores the value inline.
// Out-of-range record/field/index references collapse to `Blank` so the
// rest of the evaluator can stay branch-free; cache-record corruption
// is the OOXML reader's responsibility to surface.
Value cell_value(const PivotCache& cache, const PivotCacheRecord& record, std::size_t field_index) {
  if (field_index >= cache.fields().size() || field_index >= record.cells.size()) {
    return Value::blank();
  }
  const auto& field = cache.fields()[field_index];
  const Value& cell = record.cells[field_index];
  if (field.shared_items.empty()) {
    return cell;
  }
  if (!cell.is_number()) {
    return cell;  // Inline override (rare; Excel allows it).
  }
  const double idx = cell.as_number();
  if (idx < 0.0) {
    return Value::blank();
  }
  const auto i = static_cast<std::size_t>(idx);
  if (i >= field.shared_items.size()) {
    return Value::blank();
  }
  return field.shared_items[i];
}

// ---------------------------------------------------------------------------
// Aggregation primitives.
// ---------------------------------------------------------------------------
//
// Each aggregator ignores `Blank`. Errors propagate: the first error in
// the input dominates the output, matching `SUM(#DIV/0!, 1) -> #DIV/0!`.
// Booleans coerce numerically (TRUE=1, FALSE=0) for arithmetic
// aggregations; for `Count`, booleans are non-blank so they count, which
// also matches Excel's COUNTA on a boolean column.

// Returns the first `Value::error` found in `values`, or `std::nullopt`.
const Value* first_error(const std::vector<Value>& values) {
  for (const auto& v : values) {
    if (v.is_error()) {
      return &v;
    }
  }
  return nullptr;
}

// Numeric coercion for arithmetic aggregations. Booleans coerce; text
// is skipped (Excel's SUM/MAX/MIN over a Value column ignore text).
// `out` receives the coerced number on success.
bool coerce_arithmetic(const Value& v, double& out) noexcept {
  switch (v.kind()) {
    case ValueKind::Number:
      out = v.as_number();
      return true;
    case ValueKind::Bool:
      out = v.as_boolean() ? 1.0 : 0.0;
      return true;
    default:
      return false;
  }
}

Value AggregateSum(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double sum = 0.0;
  for (const auto& v : values) {
    double n = 0.0;
    if (coerce_arithmetic(v, n)) {
      sum += n;
    }
  }
  return Value::number(sum);
}

// Excel's pivot `Count` mirrors COUNTA: any non-blank cell counts,
// including text and booleans.
Value AggregateCount(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double count = 0.0;
  for (const auto& v : values) {
    if (!v.is_blank()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

// Excel's pivot `CountNumbers` mirrors COUNT: only numeric cells
// (booleans included, per Excel).
Value AggregateCountNumbers(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double count = 0.0;
  for (const auto& v : values) {
    if (v.is_number() || v.is_boolean()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

Value AggregateAverage(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double sum = 0.0;
  std::size_t n = 0;
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      sum += x;
      ++n;
    }
  }
  if (n == 0) {
    return Value::error(ErrorCode::Div0);
  }
  return Value::number(sum / static_cast<double>(n));
}

Value AggregateMax(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  bool seen = false;
  double best = -std::numeric_limits<double>::infinity();
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      if (!seen || x > best) {
        best = x;
        seen = true;
      }
    }
  }
  // Excel's pivot MAX over an empty/all-text group returns 0.
  return Value::number(seen ? best : 0.0);
}

Value AggregateMin(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  bool seen = false;
  double best = std::numeric_limits<double>::infinity();
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      if (!seen || x < best) {
        best = x;
        seen = true;
      }
    }
  }
  return Value::number(seen ? best : 0.0);
}

// Two-pass variance computation. Mirrors `VAR.S` / `VAR.P` from
// `src/eval/builtins/stats.cpp`: collect numerics (booleans coerce as
// 0/1, text and blanks are ignored), then compute mean and sum of
// squared deviations. `population` controls the divisor: when true, n;
// otherwise n - 1. The `min_n` guard rejects samples too small for the
// chosen variant (n < 2 for sample, n < 1 for population) with
// `#DIV/0!`, matching Excel's pivot behaviour. Errors in the input
// dominate (handled by the caller's `first_error` short-circuit).
Value variance_helper(const std::vector<Value>& values, bool population) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  std::vector<double> xs;
  xs.reserve(values.size());
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      xs.push_back(x);
    }
  }
  const std::size_t min_n = population ? 1u : 2u;
  if (xs.size() < min_n) {
    return Value::error(ErrorCode::Div0);
  }
  double sum = 0.0;
  for (double x : xs) {
    sum += x;
  }
  const double mean = sum / static_cast<double>(xs.size());
  double ss = 0.0;
  for (double x : xs) {
    const double d = x - mean;
    ss += d * d;
  }
  const double divisor = population ? static_cast<double>(xs.size()) : static_cast<double>(xs.size() - 1u);
  const double r = ss / divisor;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value AggregateVar(const std::vector<Value>& values) {
  return variance_helper(values, /*population=*/false);
}

Value AggregateVarP(const std::vector<Value>& values) {
  return variance_helper(values, /*population=*/true);
}

Value AggregateStdDev(const std::vector<Value>& values) {
  Value v = variance_helper(values, /*population=*/false);
  if (!v.is_number()) {
    return v;
  }
  const double r = std::sqrt(v.as_number());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value AggregateStdDevP(const std::vector<Value>& values) {
  Value v = variance_helper(values, /*population=*/true);
  if (!v.is_number()) {
    return v;
  }
  const double r = std::sqrt(v.as_number());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value AggregateProduct(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  bool seen = false;
  double product = 1.0;
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      product *= x;
      seen = true;
    }
  }
  // Excel's pivot PRODUCT on an empty/all-text group returns 0.
  return Value::number(seen ? product : 0.0);
}

Value apply_aggregation(Aggregation agg, const std::vector<Value>& values) {
  switch (agg) {
    case Aggregation::Sum:
      return AggregateSum(values);
    case Aggregation::Count:
      return AggregateCount(values);
    case Aggregation::Average:
      return AggregateAverage(values);
    case Aggregation::Max:
      return AggregateMax(values);
    case Aggregation::Min:
      return AggregateMin(values);
    case Aggregation::Product:
      return AggregateProduct(values);
    case Aggregation::CountNumbers:
      return AggregateCountNumbers(values);
    case Aggregation::StdDev:
      return AggregateStdDev(values);
    case Aggregation::StdDevP:
      return AggregateStdDevP(values);
    case Aggregation::Var:
      return AggregateVar(values);
    case Aggregation::VarP:
      return AggregateVarP(values);
  }
  return Value::error(ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Record-level filters: manual `items[]` visibility + axis-level label
// filters from `active_filters`.
// ---------------------------------------------------------------------------

// Resolves a filter `field_name` against `table.fields()`. Returns the
// field's index on match, `std::nullopt` if no field has that name.
// Cache fields and pivot fields share the same index in MVP, so this
// also indexes into the cache record.
std::optional<std::size_t> resolve_filter_field(const PivotTable& table, const std::string& name) {
  for (std::size_t i = 0; i < table.fields().size(); ++i) {
    if (table.fields()[i].source_name == name) {
      return i;
    }
  }
  return std::nullopt;
}

// Pulls the string payload of a filter when one is expected. Returns an
// empty `string_view` when the variant carries a non-string value, which
// effectively makes the filter a no-op rather than a runtime error.
std::string_view filter_string_value(const PivotFilter& f) {
  if (const auto* s = std::get_if<std::string>(&f.value)) {
    return *s;
  }
  return {};
}

// Pulls the numeric payload of a filter when one is expected. The
// variant carries either `int` or `double`; both coerce to double here.
// Returns 0 for non-numeric variants (caller treats as no-op).
double filter_number_value(const PivotFilter& f) {
  if (const auto* d = std::get_if<double>(&f.value)) {
    return *d;
  }
  if (const auto* i = std::get_if<int>(&f.value)) {
    return static_cast<double>(*i);
  }
  return 0.0;
}

// Pulls the numeric upper-bound payload of a range filter. Returns
// `std::nullopt` when `value_high` is `monostate` (i.e. caller didn't
// set it), signalling the range is unbounded above; the calling filter
// then no-ops.
std::optional<double> filter_number_value_high(const PivotFilter& f) {
  if (const auto* d = std::get_if<double>(&f.value_high)) {
    return *d;
  }
  if (const auto* i = std::get_if<int>(&f.value_high)) {
    return static_cast<double>(*i);
  }
  return std::nullopt;
}

// Evaluates a `LabelDate` filter against the underlying numeric value
// of a record's field. The `field_name` ostensibly identifies a date
// column, so the cache value should be a date serial; non-numeric cells
// (text labels, blanks, errors) are not in the date domain and skip the
// filter rather than being dropped.
bool label_date_filter_passes(const PivotFilter& f, const Value& v) {
  if (!v.is_number()) {
    return true;  // Non-numeric cells skip the date filter.
  }
  const double serial = v.as_number();
  const double lo = filter_number_value(f);
  const auto hi_or = filter_number_value_high(f);
  if (!hi_or) {
    return true;  // Half-open / no upper bound: treat as no-op.
  }
  return serial >= lo && serial <= *hi_or;
}

// Evaluates a single label-flavoured filter against `label`. Value-filter
// types short-circuit to true: those are applied post-aggregation.
bool label_filter_passes(const PivotFilter& f, const std::string& label) {
  switch (f.type) {
    case FilterType::LabelContains: {
      const std::string_view needle = filter_string_value(f);
      if (needle.empty()) {
        return true;
      }
      return label.find(needle) != std::string::npos;
    }
    case FilterType::LabelBeginsWith: {
      const std::string_view needle = filter_string_value(f);
      if (needle.empty()) {
        return true;
      }
      return label.size() >= needle.size() && label.compare(0, needle.size(), needle) == 0;
    }
    case FilterType::ValueTop10:
    case FilterType::ValueGreaterThan:
    case FilterType::ValueBetween:
    case FilterType::LabelDate:
      // Value-axis filters: handled later by `apply_value_filters`.
      // `LabelDate` would need a date-range payload; deferred.
      return true;
  }
  return true;
}

// True iff `record` survives the manual `items[]` filter on every field
// that declares one, AND the axis-level label filters in
// `active_filters`. Empty `items` lists match all values (Excel default
// — items[] is only authored when the user has hidden at least one
// value); axis filters with unresolved field names are skipped.
bool record_passes_manual_filter(const PivotTable& table, const PivotCache& cache, const PivotCacheRecord& record) {
  for (std::size_t fi = 0; fi < table.fields().size(); ++fi) {
    const PivotField& field = table.fields()[fi];
    if (field.items.empty()) {
      continue;
    }
    const Value v = cell_value(cache, record, fi);
    const std::string name = display_string(v);
    for (const PivotItem& item : field.items) {
      if (!item.visible && item.name == name) {
        return false;
      }
    }
  }
  for (const PivotFilter& f : table.active_filters()) {
    auto fi_or = resolve_filter_field(table, f.field_name);
    if (!fi_or) {
      continue;
    }
    const Value v = cell_value(cache, record, *fi_or);
    if (f.type == FilterType::LabelDate) {
      // Date-range filters need the underlying numeric serial; the
      // rendered label string would lose precision and locale.
      if (!label_date_filter_passes(f, v)) {
        return false;
      }
      continue;
    }
    const std::string label = display_string(v);
    if (!label_filter_passes(f, label)) {
      return false;
    }
  }
  return true;
}

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

// ---------------------------------------------------------------------------
// Hierarchy construction.
// ---------------------------------------------------------------------------
//
// We build the tree as a nested `std::map<Value, ...>` so the ordering
// emerges naturally from `ValueLess`. After all records are inserted we
// flatten into `RowHierarchyNode` / `ColHierarchyNode` and remember the
// leaf path that each surviving record lands on so the per-leaf
// aggregation pass can reuse the work without rewalking the tree.

struct HierNode {
  std::map<Value, HierNode, ValueLess> children;
  // Index into the flat leaf array assigned during finalisation. Leaves
  // only.
  std::size_t leaf_index = static_cast<std::size_t>(-1);
  // When non-empty, used in place of `display_string(key)` for this
  // node's label. Set by `insert_path` for date-grouped fields where the
  // bucket label diverges from the raw value's textual form.
  std::string label_override;
};

struct HierLevel {
  std::uint32_t field_index;         ///< Index into `PivotTable::fields()`.
  const PivotDateGroup* date_group;  ///< Non-null when this level buckets dates.
};

// Inserts `record` into `tree`, walking `levels`. Returns the leaf
// `HierNode*`. The caller assigns leaf indices in a second pass. When a
// level carries a `date_group`, the cache value is bucketed first; the
// label is stashed on the inserted child for the renderer to surface.
HierNode* insert_path(const PivotCache& cache, const std::vector<HierLevel>& levels, const PivotCacheRecord& record,
                      HierNode& root) {
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
  }
  return cursor;
}

// Returns the display label for `(key, child)`: the override if set,
// otherwise the standard `display_string(key)`. Used by all hierarchy
// flatten / subtotal-walk sites so date-grouped buckets surface their
// formatted label rather than the synthetic numeric sort key.
std::string node_label(const Value& key, const HierNode& child) {
  if (!child.label_override.empty()) {
    return child.label_override;
  }
  return display_string(key);
}

// Recursively flattens a hierarchy into `Node` form (template so we can
// produce both `RowHierarchyNode` and `ColHierarchyNode` from one
// implementation). On the way, assigns each leaf a dense index and
// pushes the corresponding `HierNode*` into `leaves` so a second pass
// can attach record indices.
template <class Node>
void finalize_hierarchy(HierNode& tree, std::vector<Node>& out, std::vector<HierNode*>& leaves) {
  if (tree.children.empty()) {
    return;
  }
  out.reserve(tree.children.size());
  for (auto& [key, child] : tree.children) {
    Node node;
    node.label = node_label(key, child);
    if (child.children.empty()) {
      child.leaf_index = leaves.size();
      leaves.push_back(&child);
    } else {
      finalize_hierarchy<Node>(child, node.children, leaves);
    }
    out.push_back(std::move(node));
  }
}

// Removes leaves at positions where `keep[i]` is false, then prunes any
// interior node whose subtree becomes empty. `leaf_cursor` is advanced
// once per visited leaf so the caller's flat `keep` vector lines up with
// the DFS pre-order leaf enumeration produced by `finalize_hierarchy`.
// Returns true iff `node` (or any of its descendants) survives the prune.
template <class Node>
bool prune_node(Node& node, const std::vector<bool>& keep, std::size_t& leaf_cursor) {
  if (node.children.empty()) {
    const bool survives = (leaf_cursor < keep.size()) ? keep[leaf_cursor] : true;
    ++leaf_cursor;
    return survives;
  }
  std::vector<Node> kept;
  kept.reserve(node.children.size());
  for (auto& child : node.children) {
    if (prune_node(child, keep, leaf_cursor)) {
      kept.push_back(std::move(child));
    }
  }
  node.children = std::move(kept);
  return !node.children.empty();
}

// Top-level driver for `prune_node`: walks each root in document order
// while threading a single leaf cursor through the whole tree so the
// caller's `keep` vector, indexed by DFS pre-order leaf position, lines
// up correctly across roots. Roots whose subtrees become empty are
// discarded.
template <class Node>
void prune_top_level(std::vector<Node>& roots, const std::vector<bool>& keep) {
  std::size_t cursor = 0;
  std::vector<Node> kept;
  kept.reserve(roots.size());
  for (auto& root : roots) {
    if (prune_node(root, keep, cursor)) {
      kept.push_back(std::move(root));
    }
  }
  roots = std::move(kept);
}

// ---------------------------------------------------------------------------
// Result-side text reification.
// ---------------------------------------------------------------------------
//
// `PivotResult::values` / `subtotals` / grand totals must outlive the
// cache they were computed against (GETPIVOTDATA reads them outside of
// any specific evaluation arena). Numbers, bools, errors, and blanks
// are trivially copyable. Text is the only kind that needs storage —
// we copy the bytes into `result.text_storage` and rebuild a `Value`
// pointing into the deque entry. Pointer/iterator stability of
// `std::deque` keeps the views valid across subsequent appends.
Value reify(const Value& v, PivotResult& result) {
  if (!v.is_text()) {
    return v;
  }
  result.text_storage.emplace_back(v.as_text());
  return Value::text(result.text_storage.back());
}

}  // namespace

Expected<PivotResult, Error> evaluate(const PivotTable& table, const PivotCache& cache) {
  // 1. Validate.
  if (table.pivot_cache_id() != cache.cache_id()) {
    return make_error(FormulonErrorCode::kEvalPivotMissing, "pivot table cache_id does not match supplied PivotCache",
                      "table=" + table.name() + " table.cache_id=" + std::to_string(table.pivot_cache_id()) +
                          " cache.cache_id=" + std::to_string(cache.cache_id()));
  }
  for (std::size_t i = 0; i < table.data_fields().size(); ++i) {
    const PivotDataField& df = table.data_fields()[i];
    if (df.field_index >= cache.fields().size()) {
      return make_error(FormulonErrorCode::kEvalPivotInvalid, "data field references out-of-range cache field",
                        "data_field=" + df.name + " field_index=" + std::to_string(df.field_index) +
                            " cache_fields=" + std::to_string(cache.fields().size()));
    }
  }

  // 2. Filter records.
  std::vector<std::size_t> surviving;
  surviving.reserve(cache.records().size());
  for (std::size_t i = 0; i < cache.records().size(); ++i) {
    if (record_passes_manual_filter(table, cache, cache.records()[i])) {
      surviving.push_back(i);
    }
  }

  // 3. Build hierarchies.
  //
  // A field whose `date_group` is set bucketises the cache value at
  // hierarchy-insertion time; we plumb the optional through `HierLevel`
  // so `insert_path` can call the bucketer without re-walking the table
  // metadata.
  auto level_for = [&](std::uint32_t fi) -> HierLevel {
    const PivotDateGroup* dg = nullptr;
    if (fi < table.fields().size() && table.fields()[fi].date_group.has_value()) {
      dg = &*table.fields()[fi].date_group;
    }
    return HierLevel{fi, dg};
  };
  std::vector<HierLevel> row_levels;
  row_levels.reserve(table.row_field_order().size());
  for (std::uint32_t fi : table.row_field_order()) {
    row_levels.push_back(level_for(fi));
  }
  std::vector<HierLevel> col_levels;
  col_levels.reserve(table.col_field_order().size());
  for (std::uint32_t fi : table.col_field_order()) {
    col_levels.push_back(level_for(fi));
  }

  HierNode row_tree;
  HierNode col_tree;

  // For each surviving record, remember which leaf it lands on (row +
  // col). Indices are looked up after finalisation so we don't need to
  // walk the tree a second time during aggregation.
  std::vector<HierNode*> row_leaves_for_record(surviving.size(), nullptr);
  std::vector<HierNode*> col_leaves_for_record(surviving.size(), nullptr);

  for (std::size_t i = 0; i < surviving.size(); ++i) {
    const PivotCacheRecord& rec = cache.records()[surviving[i]];
    if (!row_levels.empty()) {
      row_leaves_for_record[i] = insert_path(cache, row_levels, rec, row_tree);
    }
    if (!col_levels.empty()) {
      col_leaves_for_record[i] = insert_path(cache, col_levels, rec, col_tree);
    }
  }

  PivotResult result;
  std::vector<HierNode*> row_leaves;
  std::vector<HierNode*> col_leaves;
  finalize_hierarchy<RowHierarchyNode>(row_tree, result.rows, row_leaves);
  finalize_hierarchy<ColHierarchyNode>(col_tree, result.cols, col_leaves);

  // Degenerate axis: if a side has no field configured, treat it as a
  // single implicit leaf so the values matrix still has a slot per
  // surviving record group on the populated axis.
  const std::size_t row_leaf_count = row_levels.empty() ? 1 : row_leaves.size();
  const std::size_t col_leaf_count = col_levels.empty() ? 1 : col_leaves.size();
  const std::size_t data_field_count = table.data_fields().size();

  // Bucket surviving record indices by (row_leaf, col_leaf).
  // `[row_leaf][col_leaf]` -> indices into `cache.records()`.
  std::vector<std::vector<std::vector<std::size_t>>> buckets(row_leaf_count,
                                                             std::vector<std::vector<std::size_t>>(col_leaf_count));

  for (std::size_t i = 0; i < surviving.size(); ++i) {
    const std::size_t r = row_levels.empty() ? 0 : row_leaves_for_record[i]->leaf_index;
    const std::size_t c = col_levels.empty() ? 0 : col_leaves_for_record[i]->leaf_index;
    buckets[r][c].push_back(surviving[i]);
  }

  // 4. Aggregate per (row_leaf, col_leaf, data_field).
  //
  // Defensive overflow guard: on 32-bit `size_t` (WASM) a pathological
  // pivot configuration with very large axis cardinalities could wrap
  // `row_leaf_count * col_leaf_count`, leaving the nested vector
  // inconsistent. Checked multiplication keeps the failure recoverable
  // (caller surfaces `kFnOverflow`) rather than silently corrupting the
  // result matrix.
  auto value_count_or = checked_mul_size_t(row_leaf_count, col_leaf_count);
  if (!value_count_or) {
    return value_count_or.error();
  }
  result.values.assign(row_leaf_count, std::vector<std::vector<Value>>(col_leaf_count));
  for (std::size_t r = 0; r < row_leaf_count; ++r) {
    for (std::size_t c = 0; c < col_leaf_count; ++c) {
      result.values[r][c].reserve(data_field_count);
      const std::vector<std::size_t>& records = buckets[r][c];
      for (const PivotDataField& df : table.data_fields()) {
        std::vector<Value> column;
        column.reserve(records.size());
        for (std::size_t rec_idx : records) {
          column.push_back(cell_value(cache, cache.records()[rec_idx], df.field_index));
        }
        result.values[r][c].push_back(reify(apply_aggregation(df.aggregation, column), result));
      }
    }
  }

  // 5. Row-direction subtotals.
  //
  // Walk the row hierarchy; at each non-leaf level whose field declares
  // `subtotal_top` or any `subtotal_fns`, aggregate the union of all
  // descendant leaves' records using the data field's own aggregation.
  // For MVP we surface one subtotal slot per data field. Column-axis
  // subtotals are deferred.
  //
  // The flat-list shape (`subtotals[i]` is one row of the result, no
  // tree mirror) is convenient for GETPIVOTDATA, which addresses
  // subtotals by the sequence in which they appear when walking the row
  // hierarchy in document order.
  std::vector<std::vector<std::size_t>> row_subtotal_leaf_sets;
  std::vector<std::vector<std::size_t>> col_subtotal_leaf_sets;

  if (!row_levels.empty() && data_field_count > 0) {
    // A reusable closure that walks the tree depth-first and emits
    // subtotal rows where appropriate.
    std::vector<std::size_t> stack_row_leaves;  // current path's leaf indices
    std::vector<std::vector<Value>>& subtotals = result.subtotals;

    // Per-level cursor into row_levels keyed by depth.
    auto field_at_depth = [&](std::size_t depth) -> const PivotField* {
      if (depth >= row_levels.size()) {
        return nullptr;
      }
      const std::uint32_t fi = row_levels[depth].field_index;
      if (fi >= table.fields().size()) {
        return nullptr;
      }
      return &table.fields()[fi];
    };

    // Recursive walk implemented with an explicit stack to avoid lambda
    // recursion gymnastics. Each `Frame` owns iterators into its level.
    struct Frame {
      HierNode* node;
      std::map<Value, HierNode, ValueLess>::iterator it;
      std::size_t depth;
      std::size_t collected_start;  // index into `stack_row_leaves` at frame entry
      std::vector<std::string> labels;
    };

    std::vector<Frame> stack;
    stack.push_back({&row_tree, row_tree.children.begin(), 0, 0, {}});

    while (!stack.empty()) {
      Frame& top = stack.back();
      if (top.it == top.node->children.end()) {
        // All children processed: emit a subtotal for this non-root,
        // non-leaf node when the field requests it.
        if (top.depth > 0 && !top.node->children.empty()) {
          const PivotField* field = field_at_depth(top.depth - 1);
          const bool wants_subtotal = field != nullptr && (field->subtotal_top || !field->subtotal_fns.empty());
          if (wants_subtotal) {
            // Aggregate over all surviving records whose row-leaf
            // index is in [collected_start .. stack_row_leaves.size()).
            std::vector<Value> row_values(data_field_count, Value::blank());
            std::vector<std::vector<Value>> col_values(col_leaf_count,
                                                       std::vector<Value>(data_field_count, Value::blank()));
            for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
              const PivotDataField& df = table.data_fields()[df_idx];
              std::vector<Value> column;
              std::vector<std::vector<Value>> columns_by_col(col_leaf_count);
              for (std::size_t leaf_idx_iter = top.collected_start; leaf_idx_iter < stack_row_leaves.size();
                   ++leaf_idx_iter) {
                const std::size_t leaf_idx = stack_row_leaves[leaf_idx_iter];
                for (std::size_t c = 0; c < col_leaf_count; ++c) {
                  for (std::size_t rec_idx : buckets[leaf_idx][c]) {
                    Value v = cell_value(cache, cache.records()[rec_idx], df.field_index);
                    column.push_back(v);
                    columns_by_col[c].push_back(v);
                  }
                }
              }
              row_values[df_idx] = reify(apply_aggregation(df.aggregation, column), result);
              for (std::size_t c = 0; c < col_leaf_count; ++c) {
                col_values[c][df_idx] = reify(apply_aggregation(df.aggregation, columns_by_col[c]), result);
              }
            }
            RowSubtotal subtotal;
            subtotal.labels = top.labels;
            subtotal.depth = static_cast<std::uint32_t>(top.depth - 1);
            subtotal.values = row_values;
            subtotal.col_values = std::move(col_values);
            row_subtotal_leaf_sets.emplace_back(
                stack_row_leaves.begin() + static_cast<std::ptrdiff_t>(top.collected_start), stack_row_leaves.end());
            result.row_subtotals.push_back(std::move(subtotal));
            subtotals.push_back(std::move(row_values));
          }
        }
        stack.pop_back();
        continue;
      }
      HierNode* child = &top.it->second;
      const std::string label = node_label(top.it->first, *child);
      ++top.it;
      if (child->children.empty()) {
        // Leaf: contributes to its enclosing frames' subtotal.
        stack_row_leaves.push_back(child->leaf_index);
      } else {
        std::vector<std::string> labels = top.labels;
        labels.push_back(label);
        stack.push_back({child, child->children.begin(), top.depth + 1, stack_row_leaves.size(), std::move(labels)});
      }
    }
  }

  // 5b. Column-direction subtotals. The shape mirrors row_subtotals but each
  // subtotal stores one row-leaf x data-field matrix because a rendered
  // subtotal column has one value per row leaf.
  if (!col_levels.empty() && data_field_count > 0) {
    std::vector<std::size_t> stack_col_leaves;

    auto field_at_depth = [&](std::size_t depth) -> const PivotField* {
      if (depth >= col_levels.size()) {
        return nullptr;
      }
      const std::uint32_t fi = col_levels[depth].field_index;
      if (fi >= table.fields().size()) {
        return nullptr;
      }
      return &table.fields()[fi];
    };

    struct Frame {
      HierNode* node;
      std::map<Value, HierNode, ValueLess>::iterator it;
      std::size_t depth;
      std::size_t collected_start;
      std::vector<std::string> labels;
    };

    std::vector<Frame> stack;
    stack.push_back({&col_tree, col_tree.children.begin(), 0, 0, {}});

    while (!stack.empty()) {
      Frame& top = stack.back();
      if (top.it == top.node->children.end()) {
        if (top.depth > 0 && !top.node->children.empty()) {
          const PivotField* field = field_at_depth(top.depth - 1);
          const bool wants_subtotal = field != nullptr && (field->subtotal_top || !field->subtotal_fns.empty());
          if (wants_subtotal) {
            ColSubtotal subtotal;
            subtotal.labels = top.labels;
            subtotal.depth = static_cast<std::uint32_t>(top.depth - 1);
            subtotal.values.assign(row_leaf_count, std::vector<Value>(data_field_count, Value::blank()));

            for (std::size_t r = 0; r < row_leaf_count; ++r) {
              for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
                const PivotDataField& df = table.data_fields()[df_idx];
                std::vector<Value> column;
                for (std::size_t leaf_idx_iter = top.collected_start; leaf_idx_iter < stack_col_leaves.size();
                     ++leaf_idx_iter) {
                  const std::size_t leaf_idx = stack_col_leaves[leaf_idx_iter];
                  for (std::size_t rec_idx : buckets[r][leaf_idx]) {
                    column.push_back(cell_value(cache, cache.records()[rec_idx], df.field_index));
                  }
                }
                subtotal.values[r][df_idx] = reify(apply_aggregation(df.aggregation, column), result);
              }
            }
            col_subtotal_leaf_sets.emplace_back(
                stack_col_leaves.begin() + static_cast<std::ptrdiff_t>(top.collected_start), stack_col_leaves.end());
            result.col_subtotals.push_back(std::move(subtotal));
          }
        }
        stack.pop_back();
        continue;
      }
      HierNode* child = &top.it->second;
      const std::string label = node_label(top.it->first, *child);
      ++top.it;
      if (child->children.empty()) {
        stack_col_leaves.push_back(child->leaf_index);
      } else {
        std::vector<std::string> labels = top.labels;
        labels.push_back(label);
        stack.push_back({child, child->children.begin(), top.depth + 1, stack_col_leaves.size(), std::move(labels)});
      }
    }
  }

  if (!result.row_subtotals.empty() && !result.col_subtotals.empty() && data_field_count > 0) {
    for (std::size_t rs = 0; rs < result.row_subtotals.size(); ++rs) {
      RowSubtotal& row_subtotal = result.row_subtotals[rs];
      row_subtotal.col_subtotal_values.assign(result.col_subtotals.size(),
                                              std::vector<Value>(data_field_count, Value::blank()));
      if (rs >= row_subtotal_leaf_sets.size()) {
        continue;
      }
      for (std::size_t cs = 0; cs < result.col_subtotals.size(); ++cs) {
        if (cs >= col_subtotal_leaf_sets.size()) {
          continue;
        }
        for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
          const PivotDataField& df = table.data_fields()[df_idx];
          std::vector<Value> column;
          for (std::size_t row_leaf : row_subtotal_leaf_sets[rs]) {
            if (row_leaf >= row_leaf_count) {
              continue;
            }
            for (std::size_t col_leaf : col_subtotal_leaf_sets[cs]) {
              if (col_leaf >= col_leaf_count) {
                continue;
              }
              for (std::size_t rec_idx : buckets[row_leaf][col_leaf]) {
                column.push_back(cell_value(cache, cache.records()[rec_idx], df.field_index));
              }
            }
          }
          row_subtotal.col_subtotal_values[cs][df_idx] = reify(apply_aggregation(df.aggregation, column), result);
        }
      }
    }
  }

  // 6. Grand totals, one slot per data field. The legacy single-value
  // `grand_total` mirrors slot 0 for existing GETPIVOTDATA callers.
  if ((table.grand_totals_rows() || table.grand_totals_cols()) && data_field_count > 0) {
    result.grand_totals.reserve(data_field_count);
    for (const PivotDataField& df : table.data_fields()) {
      std::vector<Value> column;
      column.reserve(surviving.size());
      for (std::size_t rec_idx : surviving) {
        column.push_back(cell_value(cache, cache.records()[rec_idx], df.field_index));
      }
      result.grand_totals.push_back(reify(apply_aggregation(df.aggregation, column), result));
    }
    if (!result.grand_totals.empty()) {
      result.grand_total = result.grand_totals[0];
    }
  }

  // 7. Value-axis filters (Top-N, GreaterThan, Between).
  //
  // Applied last so the pre-aggregation filter set has already shaped
  // `result.values`; the pruning here only drops surviving leaves.
  // Multi-level hierarchies are honoured: we score each leaf in DFS
  // pre-order (the order `finalize_hierarchy` assigned), compute the
  // keep-mask for the whole leaf array, then collapse the row/col tree
  // by dropping leaves whose mask is false and any interior node whose
  // subtree becomes empty. Subtotals + grand totals retain their
  // pre-filter values so a Top-N report can still surface "X out of
  // total" framing.
  for (const PivotFilter& f : table.active_filters()) {
    if (f.type != FilterType::ValueTop10 && f.type != FilterType::ValueGreaterThan &&
        f.type != FilterType::ValueBetween) {
      continue;  // Label/Date filters handled pre-aggregation.
    }
    if (data_field_count == 0) {
      continue;
    }
    // Score per leaf is the data-field-0 aggregate. For column-axis
    // filtering we score across all rows for that column; row-axis
    // similarly across all columns.
    auto leaf_score = [&](std::size_t r, std::size_t c) -> Value {
      if (r >= result.values.size() || c >= result.values[r].size() || result.values[r][c].empty()) {
        return Value::blank();
      }
      return result.values[r][c][0];
    };
    if (f.axis == PivotAxis::Row && !table.row_field_order().empty()) {
      // `n` is the number of row leaves (DFS pre-order), which is what
      // `result.values` is indexed by; `result.rows.size()` would be the
      // number of top-level row nodes and would understate `n` whenever
      // the row hierarchy is multi-level.
      const std::size_t n = result.values.size();
      if (n == 0) {
        continue;
      }
      // Compute a per-leaf scalar by summing data-field-0 across all cols.
      std::vector<double> scores(n, 0.0);
      std::vector<bool> all_blank(n, true);
      for (std::size_t r = 0; r < n; ++r) {
        for (std::size_t c = 0; c < (col_levels.empty() ? 1u : col_leaf_count); ++c) {
          const Value v = leaf_score(r, c);
          if (v.is_number()) {
            scores[r] += v.as_number();
            all_blank[r] = false;
          } else if (v.is_boolean()) {
            scores[r] += v.as_boolean() ? 1.0 : 0.0;
            all_blank[r] = false;
          }
        }
      }
      std::vector<bool> keep(n, false);
      bool keep_initialised = true;
      if (f.type == FilterType::ValueTop10) {
        const auto top_n = static_cast<std::size_t>(filter_number_value(f));
        // Sort indices by score descending; rows with all-blank scores
        // sink to the bottom regardless of N.
        std::vector<std::size_t> order(n);
        for (std::size_t i = 0; i < n; ++i) {
          order[i] = i;
        }
        std::sort(order.begin(), order.end(), [&](std::size_t a, std::size_t b) {
          if (all_blank[a] != all_blank[b]) {
            return !all_blank[a];
          }
          return scores[a] > scores[b];
        });
        const std::size_t k = std::min(top_n, n);
        for (std::size_t i = 0; i < k; ++i) {
          if (!all_blank[order[i]]) {
            keep[order[i]] = true;
          }
        }
      } else if (f.type == FilterType::ValueGreaterThan) {
        const double threshold = filter_number_value(f);
        for (std::size_t i = 0; i < n; ++i) {
          if (!all_blank[i] && scores[i] > threshold) {
            keep[i] = true;
          }
        }
      } else if (f.type == FilterType::ValueBetween) {
        const double lo = filter_number_value(f);
        const auto hi_or = filter_number_value_high(f);
        if (!hi_or) {
          // Unbounded above -> no-op for this filter.
          keep_initialised = false;
        } else {
          for (std::size_t i = 0; i < n; ++i) {
            if (!all_blank[i] && scores[i] >= lo && scores[i] <= *hi_or) {
              keep[i] = true;
            }
          }
        }
      }
      if (!keep_initialised) {
        continue;
      }
      // Prune the row hierarchy: leaves survive when `keep[leaf] == true`
      // and interior nodes survive when at least one descendant leaf
      // does. Then compact `result.values` to the surviving leaves,
      // preserving DFS order.
      prune_top_level(result.rows, keep);
      std::vector<std::vector<std::vector<Value>>> new_values;
      new_values.reserve(n);
      for (std::size_t i = 0; i < n; ++i) {
        if (keep[i]) {
          new_values.push_back(std::move(result.values[i]));
        }
      }
      result.values = std::move(new_values);
    } else if (f.axis == PivotAxis::Col && !table.col_field_order().empty()) {
      // `n` is the number of column leaves (DFS pre-order). When the row
      // axis has at least one materialised slot we read the leaf count
      // from the first row's column slice; otherwise the matrix is empty
      // and the filter is a no-op below.
      const std::size_t n = result.values.empty() ? 0 : result.values[0].size();
      if (n == 0) {
        continue;
      }
      std::vector<double> scores(n, 0.0);
      std::vector<bool> all_blank(n, true);
      const std::size_t row_n = row_levels.empty() ? 1u : result.values.size();
      for (std::size_t c = 0; c < n; ++c) {
        for (std::size_t r = 0; r < row_n; ++r) {
          const Value v = leaf_score(r, c);
          if (v.is_number()) {
            scores[c] += v.as_number();
            all_blank[c] = false;
          } else if (v.is_boolean()) {
            scores[c] += v.as_boolean() ? 1.0 : 0.0;
            all_blank[c] = false;
          }
        }
      }
      std::vector<bool> keep(n, false);
      bool keep_initialised = true;
      if (f.type == FilterType::ValueTop10) {
        const auto top_n = static_cast<std::size_t>(filter_number_value(f));
        std::vector<std::size_t> order(n);
        for (std::size_t i = 0; i < n; ++i) {
          order[i] = i;
        }
        std::sort(order.begin(), order.end(), [&](std::size_t a, std::size_t b) {
          if (all_blank[a] != all_blank[b]) {
            return !all_blank[a];
          }
          return scores[a] > scores[b];
        });
        const std::size_t k = std::min(top_n, n);
        for (std::size_t i = 0; i < k; ++i) {
          if (!all_blank[order[i]]) {
            keep[order[i]] = true;
          }
        }
      } else if (f.type == FilterType::ValueGreaterThan) {
        const double threshold = filter_number_value(f);
        for (std::size_t i = 0; i < n; ++i) {
          if (!all_blank[i] && scores[i] > threshold) {
            keep[i] = true;
          }
        }
      } else if (f.type == FilterType::ValueBetween) {
        const double lo = filter_number_value(f);
        const auto hi_or = filter_number_value_high(f);
        if (!hi_or) {
          // Unbounded above -> no-op for this filter.
          keep_initialised = false;
        } else {
          for (std::size_t i = 0; i < n; ++i) {
            if (!all_blank[i] && scores[i] >= lo && scores[i] <= *hi_or) {
              keep[i] = true;
            }
          }
        }
      }
      if (!keep_initialised) {
        continue;
      }
      // Prune the col hierarchy then compact every row's per-col slice.
      prune_top_level(result.cols, keep);
      for (auto& row_slot : result.values) {
        std::vector<std::vector<Value>> new_row;
        new_row.reserve(n);
        for (std::size_t c = 0; c < row_slot.size() && c < n; ++c) {
          if (keep[c]) {
            new_row.push_back(std::move(row_slot[c]));
          }
        }
        row_slot = std::move(new_row);
      }
    }
    // Mixed-direction (e.g. row-axis filter referencing a column field)
    // remains out of scope; such filters fall through here as a no-op.
  }

  // 8. Show-values-as transforms.
  //
  // Applied last so the input matrix already reflects pre-filter
  // aggregation + post-filter pruning. For each data field with
  // `show_as != Normal`, walk `result.values` and replace each cell with
  // the derived value. Modes:
  //
  //   * PercentOfRow / PercentOfCol: cell / row-or-col sum (Div0 when 0).
  //   * PercentOfTotal: cell / table grand total.
  //   * RunningTotalInRow / RunningTotalInCol: cumulative sum along the
  //     axis (errors short-circuit later cells in that row / column).
  //   * Index: (cell * total) / (row_sum * col_sum); Div0 if either
  //     partial is 0.
  //
  // Subtotal / grand-total propagation policy (partial today):
  //   * The three Percent* ratio modes (PercentOfRow / PercentOfCol /
  //     PercentOfTotal) propagate the transform to `row_subtotals`,
  //     `col_subtotals`, and `grand_totals` so the rendered subtotal /
  //     grand-total cells display the same ratio Excel would emit at
  //     those positions.
  //   * RunningTotalInRow / RunningTotalInCol leave subtotals and grand
  //     totals at their raw aggregate. A running total at a subtotal
  //     break is semantically the cumulative position at that point, but
  //     our subtotal rows are aggregated independently from the leaf
  //     cells; the running cumulative position is no longer recoverable
  //     post-aggregation, so we surface the raw subtotal aggregate
  //     instead of synthesising a misleading running value.
  //   * Index uses partials (row_sum * col_sum / total) that have no
  //     meaningful analogue at a subtotal break, so subtotals and grand
  //     totals stay raw for the same reason.
  // After the transform mutates `row_subtotals[i].values`, the legacy
  // mirror `result.subtotals[i]` is re-synced; likewise `result.grand_total`
  // is re-synced from `result.grand_totals[0]`.
  if (!result.values.empty() && data_field_count > 0) {
    auto cell_num = [](const Value& v) -> std::pair<bool, double> {
      if (v.is_number()) {
        return {true, v.as_number()};
      }
      if (v.is_boolean()) {
        return {true, v.as_boolean() ? 1.0 : 0.0};
      }
      return {false, 0.0};
    };
    // Scales `cell` in place by `denom`. Only acts when `cell` is a
    // numeric aggregate; non-numeric (blank / text / error) slots are
    // left untouched. Emits Div0 when `denom == 0` and the slot was
    // numeric.
    auto scale_cell = [&cell_num](Value& cell, double denom) {
      auto [ok, n] = cell_num(cell);
      if (!ok) {
        return;
      }
      if (denom == 0.0) {
        cell = Value::error(ErrorCode::Div0);
      } else {
        cell = Value::number(n / denom);
      }
    };
    const std::size_t actual_row_count = result.values.size();
    const std::size_t actual_col_count = result.values[0].size();

    for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
      const ShowValuesAs mode = table.data_fields()[df_idx].show_as;
      if (mode == ShowValuesAs::Normal) {
        continue;
      }
      switch (mode) {
        case ShowValuesAs::PercentOfRow: {
          // Per-leaf-row sums (used for both leaf cells and any
          // col_subtotal cell that lives in that leaf row).
          std::vector<double> row_sums(actual_row_count, 0.0);
          std::vector<bool> row_any_numeric(actual_row_count, false);
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count; ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              auto [ok, n] = cell_num(result.values[r][c][df_idx]);
              if (ok) {
                row_sums[r] += n;
                row_any_numeric[r] = true;
              }
            }
          }
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count; ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              if (!row_any_numeric[r]) {
                continue;
              }
              scale_cell(result.values[r][c][df_idx], row_sums[r]);
            }
          }
          // Row subtotals: each is its own "row" with its own row sum
          // taken over `col_values`. The total slot becomes 1.0 (all of
          // the row's contribution lives within itself); col_values are
          // their share of that sum; col_subtotal_values are partial
          // shares of the same row sum, matching Excel's "% of row"
          // treatment of intersection cells.
          for (RowSubtotal& sub : result.row_subtotals) {
            double sub_row_sum = 0.0;
            bool sub_row_any_numeric = false;
            for (const auto& col_slot : sub.col_values) {
              if (df_idx >= col_slot.size()) {
                continue;
              }
              auto [ok, n] = cell_num(col_slot[df_idx]);
              if (ok) {
                sub_row_sum += n;
                sub_row_any_numeric = true;
              }
            }
            for (auto& col_slot : sub.col_values) {
              if (df_idx >= col_slot.size()) {
                continue;
              }
              if (!sub_row_any_numeric) {
                continue;
              }
              scale_cell(col_slot[df_idx], sub_row_sum);
            }
            for (auto& cs_slot : sub.col_subtotal_values) {
              if (df_idx >= cs_slot.size()) {
                continue;
              }
              if (!sub_row_any_numeric) {
                continue;
              }
              scale_cell(cs_slot[df_idx], sub_row_sum);
            }
            if (df_idx < sub.values.size() && sub_row_any_numeric) {
              if (sub_row_sum == 0.0) {
                sub.values[df_idx] = Value::error(ErrorCode::Div0);
              } else {
                sub.values[df_idx] = Value::number(1.0);
              }
            }
          }
          // Col subtotals: each cell sits in some leaf row `r`, so
          // divide by the same `row_sums[r]` used for the leaf-row
          // transform.
          for (ColSubtotal& csub : result.col_subtotals) {
            for (std::size_t r = 0; r < csub.values.size() && r < actual_row_count; ++r) {
              if (df_idx >= csub.values[r].size()) {
                continue;
              }
              if (!row_any_numeric[r]) {
                continue;
              }
              scale_cell(csub.values[r][df_idx], row_sums[r]);
            }
          }
          // Grand total: under PercentOfRow the grand-total row sums to
          // itself, so the displayed value is 1.0 (Div0 if no row had
          // any numeric content).
          if (df_idx < result.grand_totals.size()) {
            bool any_row_numeric = false;
            for (std::size_t r = 0; r < actual_row_count; ++r) {
              if (row_any_numeric[r]) {
                any_row_numeric = true;
                break;
              }
            }
            auto [ok, _n] = cell_num(result.grand_totals[df_idx]);
            (void)_n;
            if (ok) {
              result.grand_totals[df_idx] = any_row_numeric ? Value::number(1.0) : Value::error(ErrorCode::Div0);
            }
          }
          break;
        }
        case ShowValuesAs::PercentOfCol: {
          // Per-leaf-col sums (mirror of PercentOfRow).
          std::vector<double> col_sums(actual_col_count, 0.0);
          std::vector<bool> col_any_numeric(actual_col_count, false);
          for (std::size_t c = 0; c < actual_col_count; ++c) {
            for (std::size_t r = 0; r < actual_row_count; ++r) {
              if (c >= result.values[r].size() || df_idx >= result.values[r][c].size()) {
                continue;
              }
              auto [ok, n] = cell_num(result.values[r][c][df_idx]);
              if (ok) {
                col_sums[c] += n;
                col_any_numeric[c] = true;
              }
            }
          }
          // Capture the grand total before any mutation; under
          // PercentOfCol the row-subtotal "row total" slot collapses to
          // its share of the grand total (the row's contribution to the
          // single-column world that PercentOfCol presents).
          double total = 0.0;
          bool total_known = false;
          if (df_idx < result.grand_totals.size()) {
            auto [ok, n] = cell_num(result.grand_totals[df_idx]);
            if (ok) {
              total = n;
              total_known = true;
            }
          }
          if (!total_known) {
            for (std::size_t c = 0; c < actual_col_count; ++c) {
              total += col_sums[c];
              if (col_any_numeric[c]) {
                total_known = true;
              }
            }
          }
          // Col-subtotal column totals: per col_subtotal, the
          // subtotal-column total = sum across its `values[r][df_idx]`
          // slots. Used for both the col_subtotal cells themselves and
          // for any row_subtotal cell that lives in that col_subtotal.
          std::vector<double> col_subtotal_totals(result.col_subtotals.size(), 0.0);
          std::vector<bool> col_subtotal_any_numeric(result.col_subtotals.size(), false);
          for (std::size_t cs = 0; cs < result.col_subtotals.size(); ++cs) {
            for (const auto& row_slot : result.col_subtotals[cs].values) {
              if (df_idx >= row_slot.size()) {
                continue;
              }
              auto [ok, n] = cell_num(row_slot[df_idx]);
              if (ok) {
                col_subtotal_totals[cs] += n;
                col_subtotal_any_numeric[cs] = true;
              }
            }
          }
          for (std::size_t c = 0; c < actual_col_count; ++c) {
            for (std::size_t r = 0; r < actual_row_count; ++r) {
              if (c >= result.values[r].size() || df_idx >= result.values[r][c].size()) {
                continue;
              }
              if (!col_any_numeric[c]) {
                continue;
              }
              scale_cell(result.values[r][c][df_idx], col_sums[c]);
            }
          }
          // Col subtotals: each col_subtotal column's cells divide by
          // that subtotal column's own sum (same shape as a leaf col).
          for (std::size_t cs = 0; cs < result.col_subtotals.size(); ++cs) {
            ColSubtotal& csub = result.col_subtotals[cs];
            for (auto& row_slot : csub.values) {
              if (df_idx >= row_slot.size()) {
                continue;
              }
              if (!col_subtotal_any_numeric[cs]) {
                continue;
              }
              scale_cell(row_slot[df_idx], col_subtotal_totals[cs]);
            }
          }
          // Row subtotals: each col_values[col] cell divides by that
          // leaf col's sum. Each col_subtotal_values[cs] cell divides by
          // the corresponding col-subtotal column total. The row's
          // overall `values[df]` slot collapses to its share of the
          // grand total (the row's contribution under col percentages).
          for (RowSubtotal& sub : result.row_subtotals) {
            for (std::size_t c = 0; c < sub.col_values.size() && c < actual_col_count; ++c) {
              if (df_idx >= sub.col_values[c].size()) {
                continue;
              }
              if (!col_any_numeric[c]) {
                continue;
              }
              scale_cell(sub.col_values[c][df_idx], col_sums[c]);
            }
            for (std::size_t cs = 0; cs < sub.col_subtotal_values.size() && cs < result.col_subtotals.size(); ++cs) {
              if (df_idx >= sub.col_subtotal_values[cs].size()) {
                continue;
              }
              if (!col_subtotal_any_numeric[cs]) {
                continue;
              }
              scale_cell(sub.col_subtotal_values[cs][df_idx], col_subtotal_totals[cs]);
            }
            if (df_idx < sub.values.size() && total_known) {
              scale_cell(sub.values[df_idx], total);
            }
          }
          // Grand total: under PercentOfCol every column sums to itself,
          // so the column-direction grand-total cell is 1.0.
          if (df_idx < result.grand_totals.size()) {
            bool any_col_numeric = false;
            for (std::size_t c = 0; c < actual_col_count; ++c) {
              if (col_any_numeric[c]) {
                any_col_numeric = true;
                break;
              }
            }
            auto [ok, _n] = cell_num(result.grand_totals[df_idx]);
            (void)_n;
            if (ok) {
              result.grand_totals[df_idx] = any_col_numeric ? Value::number(1.0) : Value::error(ErrorCode::Div0);
            }
          }
          break;
        }
        case ShowValuesAs::PercentOfTotal: {
          // Capture grand total once before mutating anything; reuse
          // the same denominator for every slot.
          double total = 0.0;
          bool total_known = false;
          if (df_idx < result.grand_totals.size()) {
            auto [ok, n] = cell_num(result.grand_totals[df_idx]);
            if (ok) {
              total = n;
              total_known = true;
            }
          }
          if (!total_known) {
            for (std::size_t r = 0; r < actual_row_count; ++r) {
              for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
                if (df_idx >= result.values[r][c].size()) {
                  continue;
                }
                auto [ok, n] = cell_num(result.values[r][c][df_idx]);
                if (ok) {
                  total += n;
                  total_known = true;
                }
              }
            }
          }
          if (!total_known) {
            break;
          }
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              scale_cell(result.values[r][c][df_idx], total);
            }
          }
          for (RowSubtotal& sub : result.row_subtotals) {
            if (df_idx < sub.values.size()) {
              scale_cell(sub.values[df_idx], total);
            }
            for (auto& col_slot : sub.col_values) {
              if (df_idx < col_slot.size()) {
                scale_cell(col_slot[df_idx], total);
              }
            }
            for (auto& cs_slot : sub.col_subtotal_values) {
              if (df_idx < cs_slot.size()) {
                scale_cell(cs_slot[df_idx], total);
              }
            }
          }
          for (ColSubtotal& csub : result.col_subtotals) {
            for (auto& row_slot : csub.values) {
              if (df_idx < row_slot.size()) {
                scale_cell(row_slot[df_idx], total);
              }
            }
          }
          if (df_idx < result.grand_totals.size()) {
            auto [ok, _n] = cell_num(result.grand_totals[df_idx]);
            (void)_n;
            if (ok) {
              result.grand_totals[df_idx] = (total == 0.0) ? Value::error(ErrorCode::Div0) : Value::number(1.0);
            }
          }
          break;
        }
        case ShowValuesAs::RunningTotalInRow: {
          // Subtotals and grand totals are intentionally left at their
          // raw aggregate; see the header comment for this section.
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            double running = 0.0;
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              Value& cell = result.values[r][c][df_idx];
              auto [ok, n] = cell_num(cell);
              if (!ok) {
                continue;
              }
              running += n;
              cell = Value::number(running);
            }
          }
          break;
        }
        case ShowValuesAs::RunningTotalInCol: {
          // Subtotals and grand totals are intentionally left at their
          // raw aggregate; see the header comment for this section.
          for (std::size_t c = 0; c < actual_col_count; ++c) {
            double running = 0.0;
            for (std::size_t r = 0; r < actual_row_count; ++r) {
              if (c >= result.values[r].size() || df_idx >= result.values[r][c].size()) {
                continue;
              }
              Value& cell = result.values[r][c][df_idx];
              auto [ok, n] = cell_num(cell);
              if (!ok) {
                continue;
              }
              running += n;
              cell = Value::number(running);
            }
          }
          break;
        }
        case ShowValuesAs::Index: {
          // Index = (cell * grand_total) / (row_sum * col_sum). Compute
          // partials on demand; if any partial is zero or non-numeric,
          // surface Div0 / leave as-is. Subtotals + grand totals remain
          // at their raw aggregate; see the header comment for this
          // section.
          double total = 0.0;
          if (df_idx < result.grand_totals.size()) {
            auto [ok, n] = cell_num(result.grand_totals[df_idx]);
            if (ok) {
              total = n;
            }
          }
          // Precompute row sums + col sums for this df.
          std::vector<double> row_sums(actual_row_count, 0.0);
          std::vector<double> col_sums(actual_col_count, 0.0);
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              auto [ok, n] = cell_num(result.values[r][c][df_idx]);
              if (ok) {
                row_sums[r] += n;
                col_sums[c] += n;
              }
            }
          }
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              Value& cell = result.values[r][c][df_idx];
              auto [ok, n] = cell_num(cell);
              if (!ok) {
                continue;
              }
              const double denom = row_sums[r] * col_sums[c];
              if (denom == 0.0) {
                cell = Value::error(ErrorCode::Div0);
              } else {
                cell = Value::number((n * total) / denom);
              }
            }
          }
          break;
        }
        case ShowValuesAs::DifferenceFrom:
        case ShowValuesAs::PercentDifferenceFrom: {
          // Resolve the base axis from `show_as_base_field`. We support
          // single-level base axis only: a base field that lives in a
          // multi-level hierarchy (e.g. `{Region, Product}` row order)
          // is treated as a no-op so the rendered values stay at their
          // raw aggregate. The same fallback applies when `base_field`
          // is unset and both axes are multi-level — the unambiguous
          // "previous" only makes sense when one axis has a single
          // ordering.
          const PivotDataField& df = table.data_fields()[df_idx];
          enum class BaseAxis { None, Row, Col } base_axis = BaseAxis::None;
          if (df.show_as_base_field.has_value()) {
            const std::uint32_t bf = *df.show_as_base_field;
            for (const std::uint32_t fi : table.row_field_order()) {
              if (fi == bf) {
                base_axis = BaseAxis::Row;
                break;
              }
            }
            if (base_axis == BaseAxis::None) {
              for (const std::uint32_t fi : table.col_field_order()) {
                if (fi == bf) {
                  base_axis = BaseAxis::Col;
                  break;
                }
              }
            }
          } else {
            // No base field set: fall back to the row axis if it is
            // single-level and has more than one leaf, else the col
            // axis if single-level, else give up.
            if (table.row_field_order().size() == 1 && actual_row_count > 1) {
              base_axis = BaseAxis::Row;
            } else if (table.col_field_order().size() == 1 && actual_col_count > 1) {
              base_axis = BaseAxis::Col;
            }
          }
          // MVP scope: only single-level base axis supported.
          if (base_axis == BaseAxis::Row && table.row_field_order().size() != 1) {
            base_axis = BaseAxis::None;
          }
          if (base_axis == BaseAxis::Col && table.col_field_order().size() != 1) {
            base_axis = BaseAxis::None;
          }
          if (base_axis == BaseAxis::None) {
            break;
          }
          const std::size_t axis_n = base_axis == BaseAxis::Row ? actual_row_count : actual_col_count;
          // Build base_pos[p] -> optional reference position along the
          // base axis. Sentinels resolve to (p-1) / (p+1); a specific
          // item index resolves to the leaf whose label matches the
          // base field's `items[index].name`.
          std::vector<std::optional<std::size_t>> base_pos(axis_n);
          const std::uint32_t base_item = df.show_as_base_item.value_or(kShowAsBasePrev);
          if (base_item == kShowAsBasePrev) {
            for (std::size_t p = 1; p < axis_n; ++p) {
              base_pos[p] = p - 1;
            }
          } else if (base_item == kShowAsBaseNext) {
            for (std::size_t p = 0; p + 1 < axis_n; ++p) {
              base_pos[p] = p + 1;
            }
          } else {
            // Specific item: look up the field's `items[base_item].name`
            // and find the matching leaf label on the chosen axis.
            std::optional<std::size_t> fixed;
            if (df.show_as_base_field.has_value()) {
              const std::uint32_t bf = *df.show_as_base_field;
              if (bf < table.fields().size()) {
                const auto& items = table.fields()[bf].items;
                if (base_item < items.size()) {
                  const std::string& target = items[base_item].name;
                  if (base_axis == BaseAxis::Row) {
                    for (std::size_t p = 0; p < result.rows.size() && p < axis_n; ++p) {
                      if (result.rows[p].label == target) {
                        fixed = p;
                        break;
                      }
                    }
                  } else {
                    for (std::size_t p = 0; p < result.cols.size() && p < axis_n; ++p) {
                      if (result.cols[p].label == target) {
                        fixed = p;
                        break;
                      }
                    }
                  }
                }
              }
            }
            if (fixed.has_value()) {
              for (std::size_t p = 0; p < axis_n; ++p) {
                base_pos[p] = *fixed;
              }
            }
            // No match -> all base_pos remain nullopt; every cell
            // becomes blank, matching Excel's behaviour for an
            // unresolved base item.
          }
          const bool percent = (mode == ShowValuesAs::PercentDifferenceFrom);
          // Snapshot original numeric cells before mutating, so a
          // mutated cell never serves as somebody else's base reference.
          // Layout: [r][c] -> {has_value, number}.
          std::vector<std::vector<std::pair<bool, double>>> snapshot(
              actual_row_count, std::vector<std::pair<bool, double>>(actual_col_count, {false, 0.0}));
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              snapshot[r][c] = cell_num(result.values[r][c][df_idx]);
            }
          }
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              const std::size_t p = base_axis == BaseAxis::Row ? r : c;
              Value& cell = result.values[r][c][df_idx];
              const auto& cur = snapshot[r][c];
              if (!cur.first) {
                continue;
              }
              if (!base_pos[p].has_value()) {
                cell = Value::blank();
                continue;
              }
              const std::size_t bp = *base_pos[p];
              const std::size_t br = base_axis == BaseAxis::Row ? bp : r;
              const std::size_t bc = base_axis == BaseAxis::Col ? bp : c;
              if (br >= snapshot.size() || bc >= snapshot[br].size()) {
                cell = Value::blank();
                continue;
              }
              const auto& base = snapshot[br][bc];
              if (!base.first) {
                cell = Value::blank();
                continue;
              }
              if (percent) {
                if (base.second == 0.0) {
                  cell = Value::error(ErrorCode::Div0);
                } else {
                  cell = Value::number(cur.second / base.second - 1.0);
                }
              } else {
                cell = Value::number(cur.second - base.second);
              }
            }
          }
          break;
        }
        case ShowValuesAs::PercentOfParentRow:
        case ShowValuesAs::PercentOfParentCol:
        case ShowValuesAs::PercentOfParent: {
          // Resolve which axis hosts the parent and which depth (in
          // that axis's field-order) the parent field sits at. For
          // PercentOfParent the axis is determined by which order the
          // base field belongs to; for the *Row / *Col variants the
          // axis is fixed and `base_field` is optional (defaults to
          // immediate parent for multi-level, grand total otherwise).
          const PivotDataField& df = table.data_fields()[df_idx];
          enum class Axis { None, Row, Col } parent_axis = Axis::None;
          std::optional<std::size_t> base_depth;
          auto find_in = [&](const std::vector<std::uint32_t>& order, std::uint32_t fi) -> std::optional<std::size_t> {
            for (std::size_t i = 0; i < order.size(); ++i) {
              if (order[i] == fi) {
                return i;
              }
            }
            return std::nullopt;
          };
          if (mode == ShowValuesAs::PercentOfParent) {
            if (df.show_as_base_field.has_value()) {
              auto rd = find_in(table.row_field_order(), *df.show_as_base_field);
              if (rd.has_value()) {
                parent_axis = Axis::Row;
                base_depth = rd;
              } else {
                auto cd = find_in(table.col_field_order(), *df.show_as_base_field);
                if (cd.has_value()) {
                  parent_axis = Axis::Col;
                  base_depth = cd;
                }
              }
            }
          } else if (mode == ShowValuesAs::PercentOfParentRow) {
            parent_axis = Axis::Row;
            if (df.show_as_base_field.has_value()) {
              base_depth = find_in(table.row_field_order(), *df.show_as_base_field);
            }
          } else {  // PercentOfParentCol
            parent_axis = Axis::Col;
            if (df.show_as_base_field.has_value()) {
              base_depth = find_in(table.col_field_order(), *df.show_as_base_field);
            }
          }
          if (parent_axis == Axis::None) {
            break;
          }
          // Build per-leaf parent total along the chosen axis.
          // Strategy: for each leaf p, find the row_subtotal/col_subtotal
          // whose `depth == base_depth` AND whose leaf-set contains p.
          // If `base_depth` is unset, fall back to the deepest enclosing
          // subtotal (for a single-level axis there is none → use the
          // grand total).
          const auto& subs_leaf_sets = parent_axis == Axis::Row ? row_subtotal_leaf_sets : col_subtotal_leaf_sets;
          const std::size_t axis_n = parent_axis == Axis::Row ? actual_row_count : actual_col_count;
          std::vector<std::optional<double>> parent_total(axis_n);
          // Compute the grand total as a fallback denominator.
          std::optional<double> grand;
          if (df_idx < result.grand_totals.size()) {
            auto [ok, n] = cell_num(result.grand_totals[df_idx]);
            if (ok) {
              grand = n;
            }
          }
          if (!grand.has_value()) {
            double t = 0.0;
            bool any = false;
            for (std::size_t r = 0; r < actual_row_count; ++r) {
              for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
                if (df_idx >= result.values[r][c].size()) {
                  continue;
                }
                auto [ok, n] = cell_num(result.values[r][c][df_idx]);
                if (ok) {
                  t += n;
                  any = true;
                }
              }
            }
            if (any) {
              grand = t;
            }
          }
          // Pick the right subtotal-values vector for the parent axis.
          for (std::size_t p = 0; p < axis_n; ++p) {
            std::optional<std::size_t> chosen_sub;
            std::size_t chosen_depth = 0;
            for (std::size_t s = 0; s < subs_leaf_sets.size(); ++s) {
              const auto& set = subs_leaf_sets[s];
              if (std::find(set.begin(), set.end(), p) == set.end()) {
                continue;
              }
              std::uint32_t sub_depth = 0;
              if (parent_axis == Axis::Row && s < result.row_subtotals.size()) {
                sub_depth = result.row_subtotals[s].depth;
              } else if (parent_axis == Axis::Col && s < result.col_subtotals.size()) {
                sub_depth = result.col_subtotals[s].depth;
              }
              if (base_depth.has_value()) {
                if (sub_depth == *base_depth) {
                  chosen_sub = s;
                  break;
                }
              } else {
                if (!chosen_sub.has_value() || sub_depth >= chosen_depth) {
                  chosen_sub = s;
                  chosen_depth = sub_depth;
                }
              }
            }
            if (chosen_sub.has_value()) {
              if (parent_axis == Axis::Row) {
                const RowSubtotal& sub = result.row_subtotals[*chosen_sub];
                if (df_idx < sub.values.size()) {
                  auto [ok, n] = cell_num(sub.values[df_idx]);
                  if (ok) {
                    parent_total[p] = n;
                  }
                }
              } else {
                // Col subtotal "row total" for a leaf row is the sum
                // across that leaf's row in the col_subtotal columns.
                // But here we need the col-axis parent total at leaf
                // col `p`: it is the sum across the col_subtotal whose
                // leaf set contains p, taken over all row leaves. We
                // surface that as the sum of the col_subtotal's per-row
                // values at this df.
                const ColSubtotal& sub = result.col_subtotals[*chosen_sub];
                double t = 0.0;
                bool any = false;
                for (const auto& row_slot : sub.values) {
                  if (df_idx >= row_slot.size()) {
                    continue;
                  }
                  auto [ok, n] = cell_num(row_slot[df_idx]);
                  if (ok) {
                    t += n;
                    any = true;
                  }
                }
                if (any) {
                  parent_total[p] = t;
                }
              }
            } else if (grand.has_value()) {
              parent_total[p] = grand;
            }
          }
          // Apply the transform. Only the leaf cells are scaled; the
          // subtotal / grand-total cells stay at their raw aggregate
          // (consistent with how RunningTotal / Index leave them alone).
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              const std::size_t p = parent_axis == Axis::Row ? r : c;
              if (!parent_total[p].has_value()) {
                continue;
              }
              scale_cell(result.values[r][c][df_idx], *parent_total[p]);
            }
          }
          break;
        }
        case ShowValuesAs::Normal:
          break;
      }
    }

    // Re-sync the legacy mirrors after the transform pass: the flat
    // `result.subtotals[i][df]` view tracks `result.row_subtotals[i].values[df]`,
    // and `result.grand_total` tracks `result.grand_totals[0]`.
    for (std::size_t i = 0; i < result.row_subtotals.size() && i < result.subtotals.size(); ++i) {
      const std::size_t n = std::min(result.subtotals[i].size(), result.row_subtotals[i].values.size());
      for (std::size_t df_idx = 0; df_idx < n; ++df_idx) {
        result.subtotals[i][df_idx] = result.row_subtotals[i].values[df_idx];
      }
    }
    if (!result.grand_totals.empty()) {
      result.grand_total = result.grand_totals[0];
    }
  }

  return result;
}

}  // namespace formulon::pivot
