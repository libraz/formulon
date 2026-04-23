// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's statistical built-in functions:
// MEDIAN, MODE / MODE.SNGL, LARGE / SMALL, PERCENTILE[.INC],
// QUARTILE[.INC], STDEV[.S] / STDEV.P, VAR[.S] / VAR.P.
//
// Argument-type rule (DIFFERENT from SUM / AVERAGE / MIN / MAX / PRODUCT):
// these functions silently SKIP text, boolean, and blank inputs instead of
// coercing them. Only values whose kind is `Number` participate. The
// dispatcher runs with `propagate_errors = true` so error-typed arguments
// still short-circuit before the impl executes; once inside the body we
// only need to filter the non-numeric non-error kinds. Contrast this with
// `Sum` in `builtins/aggregate.cpp`, which coerces every argument through
// `coerce_to_number` and surfaces `#VALUE!` on text like `"abc"`.
//
// For `LARGE`, `SMALL`, `PERCENTILE.INC`, and `QUARTILE.INC` the dispatcher
// lays out arguments in source order: any leading range expands into
// scalar cells, and the trailing scalar (k or quart) lives at
// `args[arity - 1]`. The data slice is therefore `args[0 .. arity - 2]`.
// That trimming is done explicitly at each callsite (not in the helper)
// so the slice boundary stays visible.

#include "eval/builtins/stats.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Extracts the numeric values from `args[0..count-1]`. Non-Number values
// (text / bool / blank after range expansion) are silently skipped. Errors
// never reach this helper because the dispatcher short-circuits with
// `propagate_errors = true`.
std::vector<double> collect_numerics(const Value* args, std::uint32_t count) {
  std::vector<double> out;
  out.reserve(count);
  for (std::uint32_t i = 0; i < count; ++i) {
    if (args[i].is_number()) {
      out.push_back(args[i].as_number());
    }
  }
  return out;
}

// MEDIAN(value, ...) - median of numeric values. Non-numerics are skipped;
// an empty collection yields `#NUM!`. For an even count the result is the
// average of the two central values.
Value Median(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  const double r = (n % 2u == 1u) ? xs[n / 2u] : 0.5 * (xs[(n / 2u) - 1u] + xs[n / 2u]);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// MODE(value, ...) / MODE.SNGL(value, ...) - most frequent numeric value.
// Ties are broken by first-occurrence order. An empty collection or a
// collection whose values are all unique (no value appears >= 2 times)
// yields `#N/A`. Implementation: stable-sort a (value, original-index)
// array, then sweep once to find the longest run; ties are broken by the
// smallest first-occurrence index of tied runs.
Value Mode(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < 2u) {
    return Value::error(ErrorCode::NA);
  }
  // Pair (value, original-index). Sorting by value lets us find runs, and
  // we remember the smallest original index within each run for tie-break.
  std::vector<std::pair<double, std::size_t>> pairs;
  pairs.reserve(xs.size());
  for (std::size_t i = 0; i < xs.size(); ++i) {
    pairs.emplace_back(xs[i], i);
  }
  std::sort(pairs.begin(), pairs.end(),
            [](const std::pair<double, std::size_t>& a, const std::pair<double, std::size_t>& b) {
              return a.first < b.first;
            });
  std::size_t best_run_len = 1;
  std::size_t best_first_idx = pairs[0].second;
  double best_value = pairs[0].first;
  bool best_valid = false;  // Needs run length >= 2 to qualify.
  std::size_t i = 0;
  while (i < pairs.size()) {
    std::size_t j = i;
    std::size_t first_idx = pairs[i].second;
    while (j < pairs.size() && pairs[j].first == pairs[i].first) {
      if (pairs[j].second < first_idx) {
        first_idx = pairs[j].second;
      }
      ++j;
    }
    const std::size_t run_len = j - i;
    if (run_len >= 2u) {
      // Strictly longer wins; on ties the earlier first-occurrence wins.
      if (!best_valid || run_len > best_run_len || (run_len == best_run_len && first_idx < best_first_idx)) {
        best_run_len = run_len;
        best_first_idx = first_idx;
        best_value = pairs[i].first;
        best_valid = true;
      }
    }
    i = j;
  }
  if (!best_valid) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(best_value);
}

// Helper: read the trailing scalar `k` argument for LARGE / SMALL /
// PERCENTILE / QUARTILE. Returns the raw coerced double so each caller
// applies its own range / truncation rules.
Expected<double, ErrorCode> read_kth_arg(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return d;
}

// LARGE(array, k) - k-th largest numeric. k is truncated toward zero; any
// fractional part is discarded. k must be in [1, numeric_count], else
// `#NUM!`. An empty numeric slice trivially fails the upper bound.
Value Large(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto k_raw = read_kth_arg(args[arity - 1u]);
  if (!k_raw) {
    return Value::error(k_raw.error());
  }
  const double k_floor = std::floor(k_raw.value());
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty() || k_floor < 1.0 || k_floor > static_cast<double>(xs.size())) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const auto k = static_cast<std::size_t>(k_floor);
  return Value::number(xs[xs.size() - k]);
}

// SMALL(array, k) - k-th smallest numeric. Same rules as LARGE.
Value Small(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto k_raw = read_kth_arg(args[arity - 1u]);
  if (!k_raw) {
    return Value::error(k_raw.error());
  }
  const double k_floor = std::floor(k_raw.value());
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty() || k_floor < 1.0 || k_floor > static_cast<double>(xs.size())) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const auto k = static_cast<std::size_t>(k_floor);
  return Value::number(xs[k - 1u]);
}

// PERCENTILE.INC(array, k) / PERCENTILE(array, k) - linear-interpolation
// percentile. k is the fractional rank in [0, 1]; out-of-range yields
// `#NUM!`. Empty numeric slice yields `#NUM!`. The interpolation point is
// `pos = k * (n - 1)` (0-based); fractional `pos` blends the two neighbours.
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
  const double pos = k * static_cast<double>(xs.size() - 1u);
  const double floor_pos = std::floor(pos);
  const auto idx = static_cast<std::size_t>(floor_pos);
  const double frac = pos - floor_pos;
  double r;
  if (frac == 0.0 || idx + 1u >= xs.size()) {
    r = xs[idx];
  } else {
    r = xs[idx] + frac * (xs[idx + 1u] - xs[idx]);
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// QUARTILE.INC(array, quart) / QUARTILE(array, quart) - quartile by
// `PERCENTILE.INC(array, quart/4)`. `quart` must be an integer in
// {0, 1, 2, 3, 4}; non-integer or out-of-range yields `#NUM!`.
Value QuartileInc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto q_raw = read_kth_arg(args[arity - 1u]);
  if (!q_raw) {
    return Value::error(q_raw.error());
  }
  const double q = q_raw.value();
  // Excel truncates QUARTILE's quart toward zero and rejects values
  // outside [0, 4]. A fractional quart (e.g. 1.5) is rejected.
  if (q < 0.0 || q > 4.0 || std::floor(q) != q) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const double k = q / 4.0;
  const double pos = k * static_cast<double>(xs.size() - 1u);
  const double floor_pos = std::floor(pos);
  const auto idx = static_cast<std::size_t>(floor_pos);
  const double frac = pos - floor_pos;
  double r;
  if (frac == 0.0 || idx + 1u >= xs.size()) {
    r = xs[idx];
  } else {
    r = xs[idx] + frac * (xs[idx + 1u] - xs[idx]);
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Helper: compute `(mean, sum_of_squared_deviations)` over a numeric slice.
// Empty input returns `{0, 0}` which the callers treat as a DIV/0! case.
struct MeanSS {
  double mean;
  double ss;  // Sum of squared deviations from the mean.
};
MeanSS compute_mean_ss(const std::vector<double>& xs) {
  if (xs.empty()) {
    return {0.0, 0.0};
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
  return {mean, ss};
}

// VAR.S(value, ...) / VAR(value, ...) - sample variance with divisor n - 1.
// Fewer than 2 numeric inputs yields `#DIV/0!`.
Value VarS(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < 2u) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double r = ms.ss / static_cast<double>(xs.size() - 1u);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// VAR.P(value, ...) - population variance with divisor n. A single numeric
// input yields 0; no numeric inputs yields `#DIV/0!`.
Value VarP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double r = ms.ss / static_cast<double>(xs.size());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// STDEV.S(value, ...) / STDEV(value, ...) - sample standard deviation,
// `sqrt(VAR.S)`. Fewer than 2 numeric inputs yields `#DIV/0!`.
Value StdevS(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < 2u) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double var = ms.ss / static_cast<double>(xs.size() - 1u);
  const double r = std::sqrt(var);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// STDEV.P(value, ...) - population standard deviation, `sqrt(VAR.P)`.
Value StdevP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double var = ms.ss / static_cast<double>(xs.size());
  const double r = std::sqrt(var);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace

void register_stats_builtins(FunctionRegistry& registry) {
  // Statistical aggregators. Every entry below is range-aware and keeps the
  // default `propagate_errors = true`: errors short-circuit before the impl
  // runs, and the impls filter non-numeric kinds (text / bool / blank)
  // themselves -- see the block comment at the top of this file.
  {
    FunctionDef def{"MEDIAN", 1u, kVariadic, &Median};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"MODE", 1u, kVariadic, &Mode};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // MODE.SNGL is Excel 2010+'s canonical spelling; implementation is
    // identical to MODE. Registry already handles the dotted name.
    FunctionDef def{"MODE.SNGL", 1u, kVariadic, &Mode};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // LARGE(array, k) - trailing scalar k lives at `args[arity - 1]` after
    // the dispatcher has expanded any leading range; the impl trims the
    // slice explicitly before collecting numerics.
    FunctionDef def{"LARGE", 2u, kVariadic, &Large};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"SMALL", 2u, kVariadic, &Small};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"PERCENTILE.INC", 2u, kVariadic, &PercentileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // PERCENTILE is the pre-2010 spelling of PERCENTILE.INC; same impl.
    FunctionDef def{"PERCENTILE", 2u, kVariadic, &PercentileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"QUARTILE.INC", 2u, kVariadic, &QuartileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"QUARTILE", 2u, kVariadic, &QuartileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV.S", 1u, kVariadic, &StdevS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV", 1u, kVariadic, &StdevS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV.P", 1u, kVariadic, &StdevP};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR.S", 1u, kVariadic, &VarS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR", 1u, kVariadic, &VarS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR.P", 1u, kVariadic, &VarP};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
}

}  // namespace eval
}  // namespace formulon
