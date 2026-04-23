// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's statistical built-in functions:
// MEDIAN, MODE / MODE.SNGL, LARGE / SMALL, PERCENTILE[.INC],
// QUARTILE[.INC], STDEV[.S] / STDEV.P, VAR[.S] / VAR.P, plus the "A"
// family (AVERAGEA / MAXA / MINA / VARA / VARPA / STDEVA / STDEVPA)
// and the higher-moment descriptive statistics (GEOMEAN / HARMEAN /
// DEVSQ / AVEDEV / TRIMMEAN / SKEW / SKEW.P / KURT / STANDARDIZE).
// The probability-distribution catalog (NORM.*, BINOM.DIST, POISSON.DIST,
// EXPON.DIST, CHISQ.*, T.*, F.*, plus legacy NORMSDIST / TDIST) lives
// in the sibling `stats/stats_distributions.cpp` translation unit; this
// file's `register_stats_builtins` is the single place that wires every
// entry into the registry.
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

#include "eval/builtins/stats/stats_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

// --- Shared helpers (declared in `stats/stats_helpers.h`). ---------------

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

std::vector<double> collect_a(const Value* args, std::uint32_t count) {
  std::vector<double> out;
  out.reserve(count);
  for (std::uint32_t i = 0; i < count; ++i) {
    const Value& v = args[i];
    switch (v.kind()) {
      case ValueKind::Number:
        out.push_back(v.as_number());
        break;
      case ValueKind::Bool:
        out.push_back(v.as_boolean() ? 1.0 : 0.0);
        break;
      case ValueKind::Text:
        out.push_back(0.0);
        break;
      case ValueKind::Blank:
        // Direct Blank arguments (e.g. =AVERAGEA(A1, 1) where A1 is blank)
        // count as 0 per Excel's direct-coercion rule. Range-sourced Blanks
        // were already dropped by the dispatcher.
        out.push_back(0.0);
        break;
      default:
        // Errors / arrays / refs: unreachable in practice because the
        // dispatcher short-circuits on errors and scalar-flattens refs.
        break;
    }
  }
  return out;
}

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

double mean_of(const std::vector<double>& xs) noexcept {
  double s = 0.0;
  for (double x : xs) {
    s += x;
  }
  return s / static_cast<double>(xs.size());
}

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

double InverseStandardNormal(double p) {
  // Acklam's coefficient tables. The 'a' / 'b' coefficients cover the
  // central region; 'c' / 'd' cover both tails.
  static constexpr double a1 = -3.969683028665376e+01;
  static constexpr double a2 = 2.209460984245205e+02;
  static constexpr double a3 = -2.759285104469687e+02;
  static constexpr double a4 = 1.383577518672690e+02;
  static constexpr double a5 = -3.066479806614716e+01;
  static constexpr double a6 = 2.506628277459239e+00;

  static constexpr double b1 = -5.447609879822406e+01;
  static constexpr double b2 = 1.615858368580409e+02;
  static constexpr double b3 = -1.556989798598866e+02;
  static constexpr double b4 = 6.680131188771972e+01;
  static constexpr double b5 = -1.328068155288572e+01;

  static constexpr double c1 = -7.784894002430293e-03;
  static constexpr double c2 = -3.223964580411365e-01;
  static constexpr double c3 = -2.400758277161838e+00;
  static constexpr double c4 = -2.549732539343734e+00;
  static constexpr double c5 = 4.374664141464968e+00;
  static constexpr double c6 = 2.938163982698783e+00;

  static constexpr double d1 = 7.784695709041462e-03;
  static constexpr double d2 = 3.224671290700398e-01;
  static constexpr double d3 = 2.445134137142996e+00;
  static constexpr double d4 = 3.754408661907416e+00;

  static constexpr double p_low = 0.02425;
  static constexpr double p_high = 1.0 - p_low;

  // Constant pi, matches `std::acos(-1.0)` on any IEEE-754 system. Kept
  // local to the Halley refinement below because the normalisation is the
  // only place outside `stats/stats_distributions.cpp` that needs it.
  static constexpr double kStatsPi = 3.14159265358979323846;

  double z;
  if (p < p_low) {
    const double q = std::sqrt(-2.0 * std::log(p));
    z = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1.0);
  } else if (p <= p_high) {
    const double q = p - 0.5;
    const double r = q * q;
    z = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q /
        (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1.0);
  } else {
    const double q = std::sqrt(-2.0 * std::log(1.0 - p));
    z = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1.0);
  }
  // Halley polish: one iteration is enough to converge from Acklam's
  // ~1e-6 residual down to ~1e-14. The update formula uses:
  //   residual e  = Phi(z) - p   (standard-normal CDF minus target)
  //   derivative  f'(z) = phi(z) (standard-normal PDF)
  //   f''(z) / f'(z) = -z  (ratio of PDF derivative to PDF itself)
  // giving z_new = z - (e / f'(z)) * (1 + (e / f'(z)) * z / 2). Two
  // iterations are harmless but unnecessary; one pass brings the
  // absolute error well below `1e-12` for p in (1e-300, 1 - 1e-16).
  for (int i = 0; i < 2; ++i) {
    const double phi_cdf = 0.5 * std::erfc(-z / std::sqrt(2.0));
    const double phi_pdf = std::exp(-0.5 * z * z) / std::sqrt(2.0 * kStatsPi);
    if (phi_pdf == 0.0) {
      break;
    }
    const double e = phi_cdf - p;
    const double u = e / phi_pdf;
    z -= u * (1.0 + 0.5 * z * u);
  }
  return z;
}

// --- Core descriptive-statistics builtins. --------------------------------

// MEDIAN(value, ...) - median of numeric values. Non-numerics are skipped;
// an empty collection yields `#NUM!`. For an even count the result is the
// arithmetic mean of the two middle elements.
static Value Median(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
static Value Mode(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::NA);
  }
  // Walk the input once to record first-occurrence order, then again over a
  // sorted copy to count runs; this keeps ties resolving to the earliest
  // appearance even though the frequency count itself is O(n log n).
  std::vector<double> sorted = xs;
  std::sort(sorted.begin(), sorted.end());
  std::size_t best_count = 1;
  double best_value = 0.0;
  bool have_best = false;
  std::size_t run_len = 1;
  for (std::size_t i = 1; i <= sorted.size(); ++i) {
    if (i < sorted.size() && sorted[i] == sorted[i - 1u]) {
      ++run_len;
      continue;
    }
    if (run_len > best_count) {
      best_count = run_len;
      // Tie-break: first occurrence in the original input.
      for (double x : xs) {
        if (x == sorted[i - 1u]) {
          best_value = x;
          have_best = true;
          break;
        }
      }
    } else if (run_len == best_count && have_best) {
      // Same run length as the current best; keep the earlier original-order
      // occurrence. We only need to compare if the new run's value appears
      // in `xs` before the current `best_value`.
      const double candidate = sorted[i - 1u];
      for (double x : xs) {
        if (x == best_value) {
          break;  // current best appears first; keep it.
        }
        if (x == candidate) {
          best_value = candidate;
          break;
        }
      }
    }
    run_len = 1;
  }
  if (best_count < 2u) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(best_value);
}

// LARGE(array, k) - k-th largest numeric. k is truncated toward zero; any
// fractional part is discarded. k must be in [1, numeric_count], else
// `#NUM!`. An empty numeric slice trivially fails the upper bound.
static Value Large(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
static Value Small(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
static Value PercentileInc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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

// PERCENTILE.EXC(array, k) - exclusive-interpolation percentile. `k` must
// lie strictly inside the open interval (1/(n+1), n/(n+1)); values at or
// beyond the boundary yield `#NUM!`. Empty numeric slice yields `#NUM!`.
// The interpolation point is `pos = k * (n + 1)` (1-based); if
// `1 <= floor(pos) < n` the result is
// `xs[idx-1] + (pos - idx) * (xs[idx] - xs[idx-1])`.
static Value PercentileExc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
  const std::size_t n = xs.size();
  const double pos = k * (static_cast<double>(n) + 1.0);
  const double floor_pos = std::floor(pos);
  const auto idx = static_cast<std::int64_t>(floor_pos);
  const double frac = pos - floor_pos;
  // k <= 1/(n+1) puts idx < 1; k >= n/(n+1) puts idx >= n. Both are the
  // exclusive-method boundaries and yield `#NUM!` per Excel.
  if (idx < 1 || idx >= static_cast<std::int64_t>(n)) {
    return Value::error(ErrorCode::Num);
  }
  const auto lo = static_cast<std::size_t>(idx - 1);
  const double r = xs[lo] + frac * (xs[lo + 1u] - xs[lo]);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// QUARTILE.INC(array, quart) / QUARTILE(array, quart) - quartile by
// `PERCENTILE.INC(array, quart/4)`. `quart` must be an integer in
// {0, 1, 2, 3, 4}; non-integer or out-of-range yields `#NUM!`.
static Value QuartileInc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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

// QUARTILE.EXC(array, quart) - exclusive quartile, equivalent to
// `PERCENTILE.EXC(array, quart/4)` with `quart` restricted to {1, 2, 3}.
// Unlike QUARTILE.INC there is no Q0 or Q4: `quart <= 0`, `quart >= 4`, or
// a non-integer `quart` yields `#NUM!`. Empty numeric slice yields `#NUM!`.
static Value QuartileExc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto q_raw = read_kth_arg(args[arity - 1u]);
  if (!q_raw) {
    return Value::error(q_raw.error());
  }
  const double q = q_raw.value();
  if (q < 1.0 || q > 3.0 || std::floor(q) != q) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  const double k = q / 4.0;
  const double pos = k * (static_cast<double>(n) + 1.0);
  const double floor_pos = std::floor(pos);
  const auto idx = static_cast<std::int64_t>(floor_pos);
  const double frac = pos - floor_pos;
  if (idx < 1 || idx >= static_cast<std::int64_t>(n)) {
    return Value::error(ErrorCode::Num);
  }
  const auto lo = static_cast<std::size_t>(idx - 1);
  const double r = xs[lo] + frac * (xs[lo + 1u] - xs[lo]);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// VAR.S(value, ...) / VAR(value, ...) - sample variance with divisor n - 1.
// Fewer than 2 numeric inputs yields `#DIV/0!`.
static Value VarS(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
static Value VarP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
static Value StdevS(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
static Value StdevP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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

// ---------------------------------------------------------------------------
// "A" family: AVERAGEA / MAXA / MINA / VARA / VARPA / STDEVA / STDEVPA.
// Differ from the non-A counterparts by evaluating text as 0 and Bool as
// 0 / 1 instead of skipping them. Range-sourced Blanks are still dropped;
// direct Blank arguments are counted as 0 (see `collect_a`). Registered
// with `range_filter_a_coerce = true` so the dispatcher performs the
// provenance-aware transform on range cells before the impl runs.
// ---------------------------------------------------------------------------

static Value AverageA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_a(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  double total = 0.0;
  for (double x : xs) {
    total += x;
  }
  const double r = total / static_cast<double>(xs.size());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

static Value MaxA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_a(args, arity);
  if (xs.empty()) {
    return Value::number(0.0);
  }
  double best = xs[0];
  for (std::size_t i = 1; i < xs.size(); ++i) {
    if (xs[i] > best) {
      best = xs[i];
    }
  }
  if (std::isnan(best) || std::isinf(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

static Value MinA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_a(args, arity);
  if (xs.empty()) {
    return Value::number(0.0);
  }
  double best = xs[0];
  for (std::size_t i = 1; i < xs.size(); ++i) {
    if (xs[i] < best) {
      best = xs[i];
    }
  }
  if (std::isnan(best) || std::isinf(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

static Value VarA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_a(args, arity);
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

static Value VarPA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_a(args, arity);
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

static Value StdevA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_a(args, arity);
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

static Value StdevPA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_a(args, arity);
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

// ---------------------------------------------------------------------------
// Descriptive-statistics family: GEOMEAN / HARMEAN / DEVSQ / AVEDEV /
// TRIMMEAN / SKEW / SKEW.P / KURT / STANDARDIZE.
// ---------------------------------------------------------------------------
//
// These share the skip-non-numeric rule of the MEDIAN / VAR family (only
// `Number` kind participates; Text / Bool / Blank are ignored; Errors are
// short-circuited by the dispatcher). STANDARDIZE is the odd one out --
// scalar-only, no range expansion -- but keeping it next to the family
// clusters the moment-based functions together.

// GEOMEAN(value, ...) - geometric mean. Every numeric input must be
// strictly positive; a zero or negative value (including a range cell
// coerced via the numeric provenance rule) yields `#NUM!`. Computed in
// log-space to avoid overflow for long data sets: exp(mean(ln(x_i))).
static Value GeoMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  double log_sum = 0.0;
  for (double x : xs) {
    if (x <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
    log_sum += std::log(x);
  }
  const double r = std::exp(log_sum / static_cast<double>(xs.size()));
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// HARMEAN(value, ...) - harmonic mean. Every input must be strictly
// positive; any value <= 0 yields `#NUM!`. `n / sum(1/x_i)`.
static Value HarMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  double inv_sum = 0.0;
  for (double x : xs) {
    if (x <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
    inv_sum += 1.0 / x;
  }
  if (inv_sum == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = static_cast<double>(xs.size()) / inv_sum;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// DEVSQ(value, ...) - sum of squared deviations from the mean,
// `sum((x_i - mean)^2)`. Empty numeric slice yields 0 per Excel (the
// empty sum is zero; there is no division involved).
static Value DevSq(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::number(0.0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  if (std::isnan(ms.ss) || std::isinf(ms.ss)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(ms.ss);
}

// AVEDEV(value, ...) - mean absolute deviation from the mean,
// `sum(|x_i - mean|) / n`. Empty numeric slice yields `#NUM!`.
static Value AveDev(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  const double mean = mean_of(xs);
  double abs_sum = 0.0;
  for (double x : xs) {
    abs_sum += std::fabs(x - mean);
  }
  const double r = abs_sum / static_cast<double>(xs.size());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// TRIMMEAN(array, percent) - mean after trimming `percent / 2` from each
// tail. `percent` must be in [0, 1); out-of-range yields `#NUM!`. The
// trimmed count is `floor(n * percent / 2) * 2`, i.e. rounded down to the
// nearest even integer so the two tails are symmetric. Empty numeric
// slice yields `#NUM!`.
static Value TrimMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// SKEW(value, ...) - sample skewness,
// `(n / ((n - 1)(n - 2))) * sum(((x_i - mean) / s)^3)` where `s` is the
// sample stdev. Requires at least 3 distinct non-zero deviations; fewer
// than 3 numeric inputs or zero sample variance yields `#DIV/0!`.
static Value Skew(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < 3u) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double n = static_cast<double>(xs.size());
  const double sample_var = ms.ss / (n - 1.0);
  if (sample_var == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double s = std::sqrt(sample_var);
  double cubed_sum = 0.0;
  for (double x : xs) {
    const double z = (x - ms.mean) / s;
    cubed_sum += z * z * z;
  }
  const double coeff = n / ((n - 1.0) * (n - 2.0));
  const double r = coeff * cubed_sum;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// SKEW.P(value, ...) - population skewness,
// `(1 / n) * sum(((x_i - mean) / sigma)^3)` where `sigma` is the
// population stdev. A constant data set (sigma == 0) yields `#DIV/0!`;
// fewer than 1 numeric input also yields `#DIV/0!` so callers never see
// a silent zero.
static Value SkewP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double n = static_cast<double>(xs.size());
  const double pop_var = ms.ss / n;
  if (pop_var == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double sigma = std::sqrt(pop_var);
  double cubed_sum = 0.0;
  for (double x : xs) {
    const double z = (x - ms.mean) / sigma;
    cubed_sum += z * z * z;
  }
  const double r = cubed_sum / n;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// KURT(value, ...) - excess kurtosis (Fisher's definition),
// `(n(n+1) / ((n-1)(n-2)(n-3))) * sum(((x - mean) / s)^4)
//   - 3(n-1)^2 / ((n-2)(n-3))`.
// Requires at least 4 numeric inputs (the cubic denominator collapses
// otherwise) and non-zero sample variance; both failures yield `#DIV/0!`.
static Value Kurt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < 4u) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double n = static_cast<double>(xs.size());
  const double sample_var = ms.ss / (n - 1.0);
  if (sample_var == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double s = std::sqrt(sample_var);
  double quartic_sum = 0.0;
  for (double x : xs) {
    const double z = (x - ms.mean) / s;
    quartic_sum += z * z * z * z;
  }
  const double coeff_a = (n * (n + 1.0)) / ((n - 1.0) * (n - 2.0) * (n - 3.0));
  const double coeff_b = (3.0 * (n - 1.0) * (n - 1.0)) / ((n - 2.0) * (n - 3.0));
  const double r = coeff_a * quartic_sum - coeff_b;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// STANDARDIZE(x, mean, standard_dev) - z-score, `(x - mean) / standard_dev`.
// Scalar-only: `accepts_ranges` stays false so the dispatcher coerces each
// argument directly. `standard_dev <= 0` yields `#NUM!`.
static Value Standardize(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_e = coerce_to_number(args[0]);
  if (!x_e) {
    return Value::error(x_e.error());
  }
  auto mean_e = coerce_to_number(args[1]);
  if (!mean_e) {
    return Value::error(mean_e.error());
  }
  auto sd_e = coerce_to_number(args[2]);
  if (!sd_e) {
    return Value::error(sd_e.error());
  }
  const double sd = sd_e.value();
  if (sd <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = (x_e.value() - mean_e.value()) / sd;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace stats_detail

void register_stats_builtins(FunctionRegistry& registry) {
  // Statistical aggregators. Every entry below is range-aware and keeps the
  // default `propagate_errors = true`: errors short-circuit before the impl
  // runs, and the impls filter non-numeric kinds (text / bool / blank)
  // themselves -- see the block comment at the top of this file.
  {
    FunctionDef def{"MEDIAN", 1u, kVariadic, &stats_detail::Median};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"MODE", 1u, kVariadic, &stats_detail::Mode};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // MODE.SNGL is Excel 2010+'s canonical spelling; implementation is
    // identical to MODE. Registry already handles the dotted name.
    FunctionDef def{"MODE.SNGL", 1u, kVariadic, &stats_detail::Mode};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // LARGE(array, k) - trailing scalar k lives at `args[arity - 1]` after
    // the dispatcher has expanded any leading range; the impl trims the
    // slice explicitly before collecting numerics.
    FunctionDef def{"LARGE", 2u, kVariadic, &stats_detail::Large};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"SMALL", 2u, kVariadic, &stats_detail::Small};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"PERCENTILE.INC", 2u, kVariadic, &stats_detail::PercentileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // PERCENTILE is the pre-2010 spelling of PERCENTILE.INC; same impl.
    FunctionDef def{"PERCENTILE", 2u, kVariadic, &stats_detail::PercentileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // PERCENTILE.EXC: exclusive-interpolation variant. `k` must lie in the
    // open interval (1/(n+1), n/(n+1)); out-of-range yields `#NUM!`.
    FunctionDef def{"PERCENTILE.EXC", 2u, kVariadic, &stats_detail::PercentileExc};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"QUARTILE.INC", 2u, kVariadic, &stats_detail::QuartileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"QUARTILE", 2u, kVariadic, &stats_detail::QuartileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // QUARTILE.EXC: exclusive quartile with `quart` restricted to {1, 2, 3}.
    // Shares the interpolation kernel with PERCENTILE.EXC.
    FunctionDef def{"QUARTILE.EXC", 2u, kVariadic, &stats_detail::QuartileExc};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV.S", 1u, kVariadic, &stats_detail::StdevS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV", 1u, kVariadic, &stats_detail::StdevS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV.P", 1u, kVariadic, &stats_detail::StdevP};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR.S", 1u, kVariadic, &stats_detail::VarS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR", 1u, kVariadic, &stats_detail::VarS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR.P", 1u, kVariadic, &stats_detail::VarP};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  // Legacy (pre-2010) spellings of VAR.P / STDEV.P. Same signature, same
  // impl; only the canonical name changed. Kept for Excel-97..2007 workbooks
  // whose formulas have not been rewritten to the .NEW form.
  {
    FunctionDef def{"VARP", 1u, kVariadic, &stats_detail::VarP};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEVP", 1u, kVariadic, &stats_detail::StdevP};
    def.accepts_ranges = true;
    registry.register_function(def);
  }

  // "A" family. Registered with `range_filter_a_coerce = true` so the
  // dispatcher transforms range-sourced Bool / Text / Blank cells into
  // numbers (TRUE->1, FALSE->0, Text->0, Blank dropped) before the impl
  // runs. Direct scalar arguments are handled inside `collect_a`.
  {
    FunctionDef def{"AVERAGEA", 1u, kVariadic, &stats_detail::AverageA};
    def.accepts_ranges = true;
    def.range_filter_a_coerce = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"MAXA", 1u, kVariadic, &stats_detail::MaxA};
    def.accepts_ranges = true;
    def.range_filter_a_coerce = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"MINA", 1u, kVariadic, &stats_detail::MinA};
    def.accepts_ranges = true;
    def.range_filter_a_coerce = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VARA", 1u, kVariadic, &stats_detail::VarA};
    def.accepts_ranges = true;
    def.range_filter_a_coerce = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VARPA", 1u, kVariadic, &stats_detail::VarPA};
    def.accepts_ranges = true;
    def.range_filter_a_coerce = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEVA", 1u, kVariadic, &stats_detail::StdevA};
    def.accepts_ranges = true;
    def.range_filter_a_coerce = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEVPA", 1u, kVariadic, &stats_detail::StdevPA};
    def.accepts_ranges = true;
    def.range_filter_a_coerce = true;
    registry.register_function(def);
  }

  // Descriptive statistics: GEOMEAN / HARMEAN / DEVSQ / AVEDEV / TRIMMEAN /
  // SKEW / SKEW.P / KURT / STANDARDIZE. All range-aware except STANDARDIZE
  // (scalar-only).
  {
    FunctionDef def{"GEOMEAN", 1u, kVariadic, &stats_detail::GeoMean};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"HARMEAN", 1u, kVariadic, &stats_detail::HarMean};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"DEVSQ", 1u, kVariadic, &stats_detail::DevSq};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"AVEDEV", 1u, kVariadic, &stats_detail::AveDev};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    // TRIMMEAN takes (array, percent). The trailing scalar `percent` lives
    // at args[arity-1]; the dispatcher flattens a leading RangeOp into
    // scalar cells, so the data slice is args[0..arity-2].
    FunctionDef def{"TRIMMEAN", 2u, kVariadic, &stats_detail::TrimMean};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"SKEW", 1u, kVariadic, &stats_detail::Skew};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"SKEW.P", 1u, kVariadic, &stats_detail::SkewP};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"KURT", 1u, kVariadic, &stats_detail::Kurt};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  // STANDARDIZE: strict 3-arg, scalar-only. `accepts_ranges` stays false.
  registry.register_function(FunctionDef{"STANDARDIZE", 3u, 3u, &stats_detail::Standardize});

  // Probability-distribution family. Scalar-only: `accepts_ranges` is left
  // at its default `false`, and `propagate_errors` stays `true` so the
  // dispatcher short-circuits error arguments before the impl runs. Impls
  // live in `stats/stats_distributions.cpp`.
  registry.register_function(FunctionDef{"NORM.DIST", 4u, 4u, &stats_detail::NormDist});
  registry.register_function(FunctionDef{"NORM.S.DIST", 2u, 2u, &stats_detail::NormSDist});
  registry.register_function(FunctionDef{"NORM.INV", 3u, 3u, &stats_detail::NormInv});
  registry.register_function(FunctionDef{"NORM.S.INV", 1u, 1u, &stats_detail::NormSInv});
  registry.register_function(FunctionDef{"BINOM.DIST", 4u, 4u, &stats_detail::BinomDist});
  registry.register_function(FunctionDef{"POISSON.DIST", 3u, 3u, &stats_detail::PoissonDist});
  registry.register_function(FunctionDef{"EXPON.DIST", 3u, 3u, &stats_detail::ExponDist});

  // Chi-squared distribution family. All four are scalar-only (no range
  // expansion) and lean on `stats::p_gamma` / `stats::q_gamma` from
  // `eval/stats/special_functions.h`; the inverses close the loop with a
  // Newton-Raphson iteration seeded by Wilson-Hilferty.
  registry.register_function(FunctionDef{"CHISQ.DIST", 3u, 3u, &stats_detail::ChisqDist});
  registry.register_function(FunctionDef{"CHISQ.DIST.RT", 2u, 2u, &stats_detail::ChisqDistRt});
  registry.register_function(FunctionDef{"CHISQ.INV", 2u, 2u, &stats_detail::ChisqInv});
  registry.register_function(FunctionDef{"CHISQ.INV.RT", 2u, 2u, &stats_detail::ChisqInvRt});

  // Student's t distribution family. Scalar-only; all share
  // `stats::regularized_incomplete_beta` for the CDF surface. The inverses
  // use Hill's approximation for the initial guess with a bisection
  // fallback in the steep tails.
  registry.register_function(FunctionDef{"T.DIST", 3u, 3u, &stats_detail::TDist});
  registry.register_function(FunctionDef{"T.DIST.2T", 2u, 2u, &stats_detail::TDist2T});
  registry.register_function(FunctionDef{"T.DIST.RT", 2u, 2u, &stats_detail::TDistRt});
  registry.register_function(FunctionDef{"T.INV", 2u, 2u, &stats_detail::TInv});
  registry.register_function(FunctionDef{"T.INV.2T", 2u, 2u, &stats_detail::TInv2T});

  // Snedecor's F distribution family. Scalar-only; the inverses use a
  // bisection warm-up on [1e-10, 1e10] before switching to Newton to
  // avoid oscillation near the steep upper tail.
  registry.register_function(FunctionDef{"F.DIST", 4u, 4u, &stats_detail::FDist});
  registry.register_function(FunctionDef{"F.DIST.RT", 3u, 3u, &stats_detail::FDistRt});
  registry.register_function(FunctionDef{"F.INV", 3u, 3u, &stats_detail::FInv});
  registry.register_function(FunctionDef{"F.INV.RT", 3u, 3u, &stats_detail::FInvRt});

  // Legacy (pre-2010) spellings. Signature-compatible aliases share the
  // canonical impl pointer; NORMSDIST and TDIST have bespoke wrappers
  // because their arities / tail semantics differ from the .NEW form.
  registry.register_function(FunctionDef{"NORMDIST", 4u, 4u, &stats_detail::NormDist});
  registry.register_function(FunctionDef{"NORMINV", 3u, 3u, &stats_detail::NormInv});
  registry.register_function(FunctionDef{"NORMSDIST", 1u, 1u, &stats_detail::NormSDistLegacy});
  registry.register_function(FunctionDef{"NORMSINV", 1u, 1u, &stats_detail::NormSInv});
  registry.register_function(FunctionDef{"BINOMDIST", 4u, 4u, &stats_detail::BinomDist});
  registry.register_function(FunctionDef{"POISSON", 3u, 3u, &stats_detail::PoissonDist});
  registry.register_function(FunctionDef{"EXPONDIST", 3u, 3u, &stats_detail::ExponDist});
  registry.register_function(FunctionDef{"CHIDIST", 2u, 2u, &stats_detail::ChisqDistRt});
  registry.register_function(FunctionDef{"CHIINV", 2u, 2u, &stats_detail::ChisqInvRt});
  registry.register_function(FunctionDef{"FDIST", 3u, 3u, &stats_detail::FDistRt});
  registry.register_function(FunctionDef{"FINV", 3u, 3u, &stats_detail::FInvRt});
  registry.register_function(FunctionDef{"TDIST", 3u, 3u, &stats_detail::TDistLegacy});
  registry.register_function(FunctionDef{"TINV", 2u, 2u, &stats_detail::TInv2T});

  // Confidence-interval half-widths (CONFIDENCE / .NORM / .T). All
  // scalar-only; share the normal impl pointer for the two equivalent
  // spellings. The T variant uses TInvCore on df = size - 1.
  registry.register_function(FunctionDef{"CONFIDENCE", 3u, 3u, &stats_detail::ConfidenceNorm});
  registry.register_function(FunctionDef{"CONFIDENCE.NORM", 3u, 3u, &stats_detail::ConfidenceNorm});
  registry.register_function(FunctionDef{"CONFIDENCE.T", 3u, 3u, &stats_detail::ConfidenceT});

  // Binomial quantile (BINOM.INV / CRITBINOM legacy alias) and range
  // probability (BINOM.DIST.RANGE, 3 or 4 args).
  registry.register_function(FunctionDef{"BINOM.INV", 3u, 3u, &stats_detail::BinomInv});
  registry.register_function(FunctionDef{"CRITBINOM", 3u, 3u, &stats_detail::BinomInv});
  registry.register_function(FunctionDef{"BINOM.DIST.RANGE", 3u, 4u, &stats_detail::BinomDistRange});

  // Fisher transformation and inverse, both scalar-only.
  registry.register_function(FunctionDef{"FISHER", 1u, 1u, &stats_detail::Fisher});
  registry.register_function(FunctionDef{"FISHERINV", 1u, 1u, &stats_detail::FisherInv});

  // Standard-normal helpers: GAUSS = NORM.S.DIST(x, TRUE) - 0.5, PHI is
  // the standard-normal PDF.
  registry.register_function(FunctionDef{"GAUSS", 1u, 1u, &stats_detail::Gauss});
  registry.register_function(FunctionDef{"PHI", 1u, 1u, &stats_detail::Phi});

  // Negative binomial distribution. 4-arg canonical form plus the
  // pre-2010 3-arg spelling (always PMF).
  registry.register_function(FunctionDef{"NEGBINOM.DIST", 4u, 4u, &stats_detail::NegBinomDist});
  registry.register_function(FunctionDef{"NEGBINOMDIST", 3u, 3u, &stats_detail::NegBinomDistLegacy});

  // The pairwise linear-regression family (CORREL, COVARIANCE.P,
  // COVARIANCE.S, SLOPE, INTERCEPT, RSQ, FORECAST / FORECAST.LINEAR)
  // is routed through the lazy dispatch table (see
  // `eval_*_lazy` in `regression_lazy.cpp`) because each array
  // argument must preserve its (rows, cols) shape so the two inputs
  // can be shape-matched cell-by-cell. Pre-evaluating every arg via
  // the eager path would erase that shape.
}

}  // namespace eval
}  // namespace formulon
