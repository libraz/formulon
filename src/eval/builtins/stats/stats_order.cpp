// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Order-statistic and quantile family for the statistical builtins:
// MEDIAN, MODE / MODE.SNGL / MODE.MULT, LARGE / SMALL, PERCENTILE[.INC] /
// PERCENTILE.EXC, QUARTILE[.INC] / QUARTILE.EXC, TRIMMEAN.
//
// All entries share the skip-non-numeric rule: only `Number` kinds
// participate (the dispatcher already short-circuits error arguments via
// `propagate_errors = true`). LARGE / SMALL go through `collect_small_large`
// which coerces direct scalar Bool / Text to numeric and lets the
// dispatcher's `range_filter_numeric_only` filter the range-sourced cells.
//
// Registration lives in the sibling `stats.cpp` translation unit; the
// shared helpers (`collect_numerics`, `collect_small_large`,
// `build_mode_frequencies`, `percentile_inc_sorted`, `percentile_exc_sorted`)
// are declared in `stats/stats_helpers.h` so this TU can call them by
// reference without owning their bodies.

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/builtins/stats/stats_helpers.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

// MEDIAN(value, ...) - median of numeric values. Non-numerics are skipped;
// an empty collection yields `#NUM!`. For an even count the result is the
// arithmetic mean of the two middle elements.
Value Median(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  if ((n % 2u) == 1u) {
    return Value::number(xs[n / 2u]);
  }
  return Value::number(0.5 * (xs[n / 2u - 1u] + xs[n / 2u]));
}

// MODE / MODE.SNGL(value, ...) - most-frequent numeric value. Ties resolve
// to the first occurrence in the input order; if every value is unique the
// result is `#N/A`. Empty numeric slice also yields `#N/A`.
Value Mode(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::NA);
  }
  const ModeFrequencies freq = build_mode_frequencies(xs);
  if (freq.best_count < 2u) {
    return Value::error(ErrorCode::NA);
  }
  for (std::size_t i = 0; i < freq.values.size(); ++i) {
    if (freq.counts[i] == freq.best_count) {
      return Value::number(freq.values[i]);
    }
  }
  return Value::error(ErrorCode::NA);
}

// MODE.MULT(value, ...) - vertical (column) array of every value tied for
// the maximum frequency. Order is first-occurrence in the input. If no
// value repeats (or the numeric slice is empty), the result is `#N/A`,
// matching MODE / MODE.SNGL. The output is always a 1-column array even
// when only one mode exists (Excel: MODE.MULT({1,2,1}) -> {1} as a 1x1
// vertical array, which spills as a single cell).
Value ModeMult(const Value* args, std::uint32_t arity, Arena& arena) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::NA);
  }
  const ModeFrequencies freq = build_mode_frequencies(xs);
  if (freq.best_count < 2u) {
    return Value::error(ErrorCode::NA);
  }
  std::vector<double> modes;
  modes.reserve(freq.values.size());
  for (std::size_t i = 0; i < freq.values.size(); ++i) {
    if (freq.counts[i] == freq.best_count) {
      modes.push_back(freq.values[i]);
    }
  }
  const auto rows = static_cast<std::uint32_t>(modes.size());
  Value* buffer = arena.create_array<Value>(rows);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < modes.size(); ++i) {
    buffer[i] = Value::number(modes[i]);
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = rows;
  arr->cols = 1u;
  arr->cells = buffer;
  return Value::array(arr);
}

// Shared body of LARGE / SMALL. `LARGE(arr, k) == SMALL(arr, N + 1 - k)`
// with TRUNC indexing on the derived position. The bounds check is
// performed on the *raw* k, not on the truncated index: Mac Excel 365
// rejects `LARGE({10;20;30}, 0.5)` with `#NUM!` even though
// `TRUNC(N + 1 - 0.5) = 3` would land in range
// (probe `large_k_below_one_fractional`). Conversely
// `LARGE({10;20;30}, 1.9)` succeeds because raw k=1.9 satisfies
// `1 <= k <= N` and `TRUNC(N + 1 - 1.9) = TRUNC(2.1) = 2`, picking the
// second-largest element (probe `large_k_fractional_truncates`). SMALL
// applies the same rule on the lower end: `SMALL({10;20;30}, 3.0001)`
// is `#NUM!` (probe `small_k_above_n_fractional`).
//
// Direct scalar Text / Bool arguments coerce through `collect_small_large`
// (Bool -> 1 / 0, Text -> strict numeric coercion with `#VALUE!` on
// failure). Range-sourced and array-literal-sourced non-Number cells are
// dropped by the dispatcher via `range_filter_numeric_only` before
// reaching this impl.
static Value large_small(const Value* args, std::uint32_t arity, bool want_large) {
  const std::uint32_t data_count = arity - 1u;
  auto k_raw = read_kth_arg(args[arity - 1u]);
  if (!k_raw) {
    return Value::error(k_raw.error());
  }
  const double k = k_raw.value();
  auto xs_e = collect_small_large(args, data_count);
  if (!xs_e) {
    return Value::error(xs_e.error());
  }
  std::vector<double>& xs = xs_e.value();
  const auto n = static_cast<double>(xs.size());
  if (xs.empty() || k < 1.0 || k > n) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const double idx_d = want_large ? std::trunc(n + 1.0 - k) : std::trunc(k);
  const auto idx = static_cast<std::size_t>(idx_d);
  return Value::number(xs[idx - 1u]);
}

Value Large(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return large_small(args, arity, true);
}

Value Small(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return large_small(args, arity, false);
}

// PERCENTILE.INC(array, k) / PERCENTILE(array, k) - linear-interpolation
// percentile. k is the fractional rank in [0, 1]; out-of-range yields
// `#NUM!`. Empty numeric slice yields `#NUM!`. The interpolation point is
// `pos = k * (n - 1)` (0-based); fractional `pos` blends the two
// neighbours.
Value PercentileInc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto k_raw = read_kth_arg(args[arity - 1u]);
  if (!k_raw) {
    return Value::error(k_raw.error());
  }
  const double k = k_raw.value();
  if (k < 0.0 || k > 1.0) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  return percentile_inc_sorted(xs, k);
}

// PERCENTILE.EXC(array, k) - exclusive-interpolation percentile. `k` must
// lie strictly inside the open interval (1/(n+1), n/(n+1)); values at or
// beyond the boundary yield `#NUM!`. Empty numeric slice yields `#NUM!`.
// The interpolation point is `pos = k * (n + 1)` (1-based); if
// `1 <= floor(pos) < n` the result is
// `xs[idx-1] + (pos - idx) * (xs[idx] - xs[idx-1])`.
Value PercentileExc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto k_raw = read_kth_arg(args[arity - 1u]);
  if (!k_raw) {
    return Value::error(k_raw.error());
  }
  const double k = k_raw.value();
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  return percentile_exc_sorted(xs, k);
}

// QUARTILE.INC(array, quart) / QUARTILE(array, quart) - quartile by
// `PERCENTILE.INC(array, quart/4)`. `quart` must be in [0, 5);
// Excel truncates a fractional `quart` toward zero, so `1.5` is `Q1`.
Value QuartileInc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto q_raw = read_kth_arg(args[arity - 1u]);
  if (!q_raw) {
    return Value::error(q_raw.error());
  }
  const double q_in = q_raw.value();
  if (q_in < 0.0 || q_in >= 5.0) {
    return Value::error(ErrorCode::Num);
  }
  const double q = std::trunc(q_in);
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const double k = q / 4.0;
  return percentile_inc_sorted(xs, k);
}

// QUARTILE.EXC(array, quart) - exclusive quartile, equivalent to
// `PERCENTILE.EXC(array, quart/4)` with `quart` restricted to {1, 2, 3}.
// Unlike QUARTILE.INC there is no Q0 or Q4: `quart < 1` or `quart >= 4`
// yields `#NUM!`. Excel truncates a fractional `quart` toward zero, so
// `quart = 1.5` is treated as `1`. Empty numeric slice yields `#NUM!`.
Value QuartileExc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto q_raw = read_kth_arg(args[arity - 1u]);
  if (!q_raw) {
    return Value::error(q_raw.error());
  }
  const double q_in = q_raw.value();
  if (q_in < 1.0 || q_in >= 4.0) {
    return Value::error(ErrorCode::Num);
  }
  const double q = std::trunc(q_in);
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const double k = q / 4.0;
  return percentile_exc_sorted(xs, k);
}

// TRIMMEAN(array, percent) - mean after trimming `percent / 2` from each
// tail. `percent` must be in [0, 1); out-of-range yields `#NUM!`. The
// trimmed count is `floor(n * percent / 2) * 2`, i.e. rounded down to the
// nearest even integer so the two tails are symmetric. Empty numeric
// slice yields `#NUM!`.
Value TrimMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto p_raw = read_kth_arg(args[arity - 1u]);
  if (!p_raw) {
    return Value::error(p_raw.error());
  }
  const double percent = p_raw.value();
  if (percent < 0.0 || percent >= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  const std::size_t trim_each = static_cast<std::size_t>(std::floor(static_cast<double>(n) * percent / 2.0));
  const std::size_t kept = n - 2u * trim_each;
  if (kept == 0) {
    return Value::error(ErrorCode::Num);
  }
  double sum = 0.0;
  for (std::size_t i = trim_each; i < n - trim_each; ++i) {
    sum += xs[i];
  }
  const double r = sum / static_cast<double>(kept);
  return finite_number_result(r);
}

}  // namespace stats_detail
}  // namespace eval
}  // namespace formulon
