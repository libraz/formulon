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
#include "eval/stats/special_functions.h"
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

// ---------------------------------------------------------------------------
// Probability-distribution family
// ---------------------------------------------------------------------------
//
// These functions are scalar-only (no range expansion): every argument is
// coerced via `coerce_to_number` or `coerce_to_bool`. The `cumulative`
// flag follows Excel's convention where any non-zero numeric (or TRUE)
// selects the CDF, and zero (or FALSE) selects the PDF/PMF. The
// dispatcher short-circuits on error-typed arguments before the impl
// runs, so no explicit error-passthrough is needed here.

// Mathematical constant pi, used to normalise the standard-normal PDF.
// Matches `std::acos(-1.0)` on any IEEE-754 system.
static constexpr double kStatsPi = 3.14159265358979323846;

// Shared body of NORM.DIST / NORM.S.DIST. Returns the PDF or CDF of a
// normal distribution with the given `mean` and `sd`. Assumes `sd > 0`;
// callers must reject `sd <= 0` with `#NUM!` up front.
Value NormDistCompute(double x, double mean, double sd, bool cumulative) {
  const double z = (x - mean) / sd;
  double r;
  if (cumulative) {
    // P(X <= x) = 0.5 * erfc(-z / sqrt(2)). `std::erfc` is the
    // complementary error function, available since C++11 <cmath>.
    r = 0.5 * std::erfc(-z / std::sqrt(2.0));
  } else {
    r = std::exp(-0.5 * z * z) / (sd * std::sqrt(2.0 * kStatsPi));
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// NORM.DIST(x, mean, sd, cumulative) - normal distribution PDF or CDF.
// `sd <= 0` yields `#NUM!`; all other finite inputs are accepted.
Value NormDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  auto mean = coerce_to_number(args[1]);
  if (!mean) {
    return Value::error(mean.error());
  }
  auto sd = coerce_to_number(args[2]);
  if (!sd) {
    return Value::error(sd.error());
  }
  auto cum = coerce_to_bool(args[3]);
  if (!cum) {
    return Value::error(cum.error());
  }
  if (sd.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return NormDistCompute(x.value(), mean.value(), sd.value(), cum.value());
}

// NORM.S.DIST(z, cumulative) - thin wrapper over NORM.DIST with
// mean = 0 and sd = 1. Does not need the sd-domain check since 1 > 0.
Value NormSDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto z = coerce_to_number(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  auto cum = coerce_to_bool(args[1]);
  if (!cum) {
    return Value::error(cum.error());
  }
  return NormDistCompute(z.value(), 0.0, 1.0, cum.value());
}

// Inverse standard-normal CDF. Uses Peter Acklam's rational
// approximation for the initial guess (good to ~1e-6 in practice across
// the unit interval) and then runs a single Halley-method refinement
// step to bring the accuracy up to ~1e-14, comfortably inside Excel's
// reported precision. Callers must guarantee `0 < p < 1`; `p <= 0` or
// `p >= 1` should surface `#NUM!` before calling in.
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

// NORM.INV(p, mean, sd) - inverse normal CDF. Excel rejects `p <= 0`,
// `p >= 1`, and `sd <= 0` with `#NUM!`.
Value NormInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p = coerce_to_number(args[0]);
  if (!p) {
    return Value::error(p.error());
  }
  auto mean = coerce_to_number(args[1]);
  if (!mean) {
    return Value::error(mean.error());
  }
  auto sd = coerce_to_number(args[2]);
  if (!sd) {
    return Value::error(sd.error());
  }
  if (p.value() <= 0.0 || p.value() >= 1.0 || sd.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double z = InverseStandardNormal(p.value());
  const double r = mean.value() + sd.value() * z;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// NORM.S.INV(p) - inverse standard-normal CDF. Equivalent to
// NORM.INV(p, 0, 1); same domain checks apply.
Value NormSInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p = coerce_to_number(args[0]);
  if (!p) {
    return Value::error(p.error());
  }
  if (p.value() <= 0.0 || p.value() >= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = InverseStandardNormal(p.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Log-PMF of Binomial(n, p) at k: computes
//   lgamma(n+1) - lgamma(k+1) - lgamma(n-k+1) + k*log(p) + (n-k)*log(1-p)
// with explicit handling for the boundary probabilities p == 0 and
// p == 1 where log(p) or log(1-p) would be -inf. Returns the PMF
// itself (exponentiated) so callers do not need to re-exp for each k.
double BinomPmf(double k, double n, double prob) {
  // Boundary cases: the generic log-space formula would produce
  // NaN from 0 * (-inf). Handle them directly so the formula stays
  // correct at the endpoints.
  if (prob == 0.0) {
    return k == 0.0 ? 1.0 : 0.0;
  }
  if (prob == 1.0) {
    return k == n ? 1.0 : 0.0;
  }
  const double log_pmf = std::lgamma(n + 1.0) - std::lgamma(k + 1.0) - std::lgamma(n - k + 1.0) + k * std::log(prob) +
                         (n - k) * std::log(1.0 - prob);
  return std::exp(log_pmf);
}

// BINOM.DIST(number_s, trials, probability_s, cumulative) - binomial
// distribution PMF or CDF. `number_s` and `trials` are floored toward
// negative infinity (Excel truncates non-integer inputs before the
// domain check). Any negative count, `number_s > trials`, or probability
// outside [0, 1] yields `#NUM!`. The CDF sums PMFs from 0 to number_s;
// this is O(trials) but matches Excel's approach for moderate n.
Value BinomDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n_s = coerce_to_number(args[0]);
  if (!n_s) {
    return Value::error(n_s.error());
  }
  auto trials = coerce_to_number(args[1]);
  if (!trials) {
    return Value::error(trials.error());
  }
  auto prob = coerce_to_number(args[2]);
  if (!prob) {
    return Value::error(prob.error());
  }
  auto cum = coerce_to_bool(args[3]);
  if (!cum) {
    return Value::error(cum.error());
  }
  const double k = std::floor(n_s.value());
  const double n = std::floor(trials.value());
  const double p = prob.value();
  if (k < 0.0 || n < 0.0 || k > n || p < 0.0 || p > 1.0) {
    return Value::error(ErrorCode::Num);
  }
  double r;
  if (cum.value()) {
    // CDF: sum pmf(i) for i in [0, k]. Loop count fits in size_t because
    // k >= 0 and k <= n, and n is a finite double bounded by the caller.
    r = 0.0;
    const auto k_int = static_cast<std::uint64_t>(k);
    for (std::uint64_t i = 0; i <= k_int; ++i) {
      r += BinomPmf(static_cast<double>(i), n, p);
    }
  } else {
    r = BinomPmf(k, n, p);
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Log-space PMF of Poisson(mean) at k: exp(-mean + k*log(mean) - lgamma(k+1)).
// Assumes `mean > 0` and `k >= 0`; callers enforce both.
double PoissonPmf(double k, double mean) {
  const double log_pmf = -mean + k * std::log(mean) - std::lgamma(k + 1.0);
  return std::exp(log_pmf);
}

// POISSON.DIST(x, mean, cumulative) - Poisson distribution PMF or CDF.
// `x` is floored toward negative infinity. `x < 0` or `mean <= 0`
// yields `#NUM!`. CDF is the O(x) partial sum of PMFs.
Value PoissonDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  auto mean_arg = coerce_to_number(args[1]);
  if (!mean_arg) {
    return Value::error(mean_arg.error());
  }
  auto cum = coerce_to_bool(args[2]);
  if (!cum) {
    return Value::error(cum.error());
  }
  const double x = std::floor(x_arg.value());
  const double mean = mean_arg.value();
  if (x < 0.0 || mean <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  double r;
  if (cum.value()) {
    r = 0.0;
    const auto x_int = static_cast<std::uint64_t>(x);
    for (std::uint64_t i = 0; i <= x_int; ++i) {
      r += PoissonPmf(static_cast<double>(i), mean);
    }
  } else {
    r = PoissonPmf(x, mean);
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Chi-squared PDF at `x` with `df` degrees of freedom, evaluated in log
// space to stay stable for large `df`:
//   pdf(x) = exp(-x/2 + (df/2-1)*log(x) - (df/2)*log(2) - lgamma(df/2))
// Callers must pre-reject `df < 1`, `x < 0`, and handle the `x == 0`
// boundary (where `log(0) == -inf` would produce NaN). When `df == 2` the
// `(df/2 - 1) * log(x)` term is `0 * log(x)`, which is 0 at x == 0 but
// surfaces as NaN via IEEE-754; callers short-circuit the boundary.
double ChisqPdf(double x, double df) noexcept {
  return std::exp(-0.5 * x + (0.5 * df - 1.0) * std::log(x) - 0.5 * df * std::log(2.0) - std::lgamma(0.5 * df));
}

// Upper bound for `df` accepted by Excel 365's CHISQ.* family. Values above
// this surface `#NUM!`; the cap matches the threshold documented in Excel's
// online help and tracked by the oracle suite.
constexpr double kChisqDfMax = 1.0e10;

// CHISQ.DIST(x, df, cumulative) - chi-squared distribution CDF or PDF.
// Excel floors `df` toward -inf and rejects non-positive `df`, `df` above
// 1e10, and negative `x` with `#NUM!`. The PDF singularity at `x == 0`
// with `df == 1` also surfaces `#NUM!`.
Value ChisqDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  auto cum = coerce_to_bool(args[2]);
  if (!cum) {
    return Value::error(cum.error());
  }
  const double x = x_arg.value();
  const double df = std::floor(df_arg.value());
  if (x < 0.0 || df < 1.0 || df > kChisqDfMax) {
    return Value::error(ErrorCode::Num);
  }
  double r;
  if (cum.value()) {
    if (x == 0.0) {
      r = 0.0;
    } else {
      r = stats::p_gamma(0.5 * df, 0.5 * x);
    }
  } else {
    // PDF with Excel-observed boundary handling at `x == 0`:
    //   df == 1: pdf diverges  -> #NUM!
    //   df == 2: pdf == 0.5
    //   df >  2: pdf == 0
    if (x == 0.0) {
      if (df == 1.0) {
        return Value::error(ErrorCode::Num);
      }
      r = (df == 2.0) ? 0.5 : 0.0;
    } else {
      r = ChisqPdf(x, df);
    }
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// CHISQ.DIST.RT(x, df) - right-tailed chi-squared CDF, `1 - CDF(x)`.
// Domain matches CHISQ.DIST's CDF branch. `df` is floored toward -inf.
// (Excel 365 does accept non-integer `df` and floors it before the domain
// check; the CHISQ.INV family uses the same convention.)
Value ChisqDistRt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  const double x = x_arg.value();
  const double df = std::floor(df_arg.value());
  if (x < 0.0 || df < 1.0 || df > kChisqDfMax) {
    return Value::error(ErrorCode::Num);
  }
  const double r = (x == 0.0) ? 1.0 : stats::q_gamma(0.5 * df, 0.5 * x);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Shared Newton-Raphson inverter for CHISQ.INV(p, df). Assumes the caller
// has already validated `0 <= p < 1` and `df` in range, and passes the
// already-floored `df`. `p == 0` is handled up-front by the wrappers; this
// routine is invoked only with `p` strictly in (0, 1). Returns NaN on
// non-convergence so callers can surface `#NUM!`.
double ChisqInvCore(double p, double df) noexcept {
  // Wilson-Hilferty transformation for the initial guess. For moderate df
  // this lands within a few percent of the true quantile; for very small
  // df / extreme p we fall back to df/2 if the guess goes negative.
  const double h = 2.0 / (9.0 * df);
  const double z = InverseStandardNormal(p);
  const double cube_arg = 1.0 - h + z * std::sqrt(h);
  double x = df * cube_arg * cube_arg * cube_arg;
  if (!(x > 0.0)) {
    // Covers negative, zero, NaN cases (e.g. df=1 and p near 0).
    x = 0.5 * df;
  }
  constexpr int kMaxIter = 100;
  constexpr double kTol = 1e-10;
  for (int i = 0; i < kMaxIter; ++i) {
    const double cdf = stats::p_gamma(0.5 * df, 0.5 * x);
    if (std::isnan(cdf)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    const double pdf = ChisqPdf(x, df);
    if (pdf <= 0.0 || std::isnan(pdf) || std::isinf(pdf)) {
      // No meaningful Newton step; accept the current x as the best
      // estimate and let the caller decide whether it's close enough.
      return x;
    }
    double step = (cdf - p) / pdf;
    double x_new = x - step;
    // Safeguard against stepping into the forbidden x < 0 half-line.
    // Halve the step until we land inside the positive reals.
    while (x_new <= 0.0) {
      step *= 0.5;
      x_new = x - step;
      if (std::abs(step) < kTol) {
        x_new = 0.5 * x;  // Final fallback: move toward zero.
        break;
      }
    }
    if (std::abs(x_new - x) < kTol * std::max(1.0, std::abs(x))) {
      return x_new;
    }
    x = x_new;
  }
  // Failed to converge after kMaxIter iterations. Return NaN so the
  // wrapper surfaces #NUM!.
  return std::numeric_limits<double>::quiet_NaN();
}

// CHISQ.INV(p, df) - inverse of the left-tailed chi-squared CDF. `p` must
// lie in `[0, 1)`; `p == 1` and `p` outside the unit interval yield `#NUM!`.
Value ChisqInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p_arg = coerce_to_number(args[0]);
  if (!p_arg) {
    return Value::error(p_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  const double p = p_arg.value();
  const double df = std::floor(df_arg.value());
  if (p < 0.0 || p >= 1.0 || df < 1.0 || df > kChisqDfMax) {
    return Value::error(ErrorCode::Num);
  }
  if (p == 0.0) {
    return Value::number(0.0);
  }
  const double r = ChisqInvCore(p, df);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// CHISQ.INV.RT(p, df) - inverse of the right-tailed CDF. `p == 1` means
// quantile at x=0 (mass 1 on the right); `p == 0` means +inf tail and
// surfaces `#NUM!`. Equivalent to `CHISQ.INV(1 - p, df)` modulo the
// closed/open endpoint conventions above.
Value ChisqInvRt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p_arg = coerce_to_number(args[0]);
  if (!p_arg) {
    return Value::error(p_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  const double p = p_arg.value();
  const double df = std::floor(df_arg.value());
  if (p <= 0.0 || p > 1.0 || df < 1.0 || df > kChisqDfMax) {
    return Value::error(ErrorCode::Num);
  }
  if (p == 1.0) {
    return Value::number(0.0);
  }
  const double r = ChisqInvCore(1.0 - p, df);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// EXPON.DIST(x, lambda, cumulative) - exponential distribution PDF or CDF.
// `x < 0` or `lambda <= 0` yields `#NUM!`. PDF: lambda * exp(-lambda*x);
// CDF: 1 - exp(-lambda*x).
Value ExponDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  auto lambda = coerce_to_number(args[1]);
  if (!lambda) {
    return Value::error(lambda.error());
  }
  auto cum = coerce_to_bool(args[2]);
  if (!cum) {
    return Value::error(cum.error());
  }
  if (x.value() < 0.0 || lambda.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = cum.value() ? 1.0 - std::exp(-lambda.value() * x.value())
                               : lambda.value() * std::exp(-lambda.value() * x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ---------------------------------------------------------------------------
// T / F distribution family
// ---------------------------------------------------------------------------
//
// All nine entries (T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T, F.DIST,
// F.DIST.RT, F.INV, F.INV.RT) share a single helper: Student's t CDF maps to
// `regularized_incomplete_beta(df/2, 1/2, df/(df + x*x))` and Snedecor's F
// CDF maps to `regularized_incomplete_beta(d1/2, d2/2, d1*x/(d1*x + d2))`.
// The inverses are Newton-Raphson over the CDF with a distribution-specific
// initial guess; bisection kicks in when the steep tails cause oscillation.

// Shared df cap. Excel 365 rejects df > 1e10 for the T and F families
// (matches the CHISQ cap); keeping the constant in this block makes the
// domain checks readable at each callsite.
constexpr double kTFdfMax = 1.0e10;

// Student's t CDF at `x` with `df` degrees of freedom, computed through the
// regularized incomplete beta. The reflection handles negative `x` via
// P(T <= x) = 1 - P(T <= -x) for x >= 0.
double TDistCdf(double x, double df) noexcept {
  const double t2 = x * x;
  const double y = df / (df + t2);
  const double half = 0.5 * stats::regularized_incomplete_beta(0.5 * df, 0.5, y);
  return (x >= 0.0) ? 1.0 - half : half;
}

// Student's t PDF at `x` with `df` degrees of freedom, computed in log
// space via lgamma to stay stable for large df (tgamma overflows past ~170).
double TDistPdf(double x, double df) noexcept {
  const double log_norm = std::lgamma(0.5 * (df + 1.0)) - std::lgamma(0.5 * df) - 0.5 * std::log(df * kStatsPi);
  const double log_kernel = -0.5 * (df + 1.0) * std::log(1.0 + x * x / df);
  return std::exp(log_norm + log_kernel);
}

// T.DIST(x, deg_freedom, cumulative) - Student's t-distribution PDF or CDF.
// `df` is floored toward -inf and must satisfy `df >= 1`; `x` is
// unrestricted (the distribution is symmetric around 0). The PDF uses
// lgamma to avoid overflow at large df.
Value TDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  auto cum = coerce_to_bool(args[2]);
  if (!cum) {
    return Value::error(cum.error());
  }
  const double x = x_arg.value();
  const double df = std::floor(df_arg.value());
  if (df < 1.0 || df > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  const double r = cum.value() ? TDistCdf(x, df) : TDistPdf(x, df);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// T.DIST.2T(x, deg_freedom) - two-tailed Student's t probability. Excel
// enforces `x >= 0` here (the function is defined as the probability that
// |T| exceeds x in absolute value); negative x yields `#NUM!`.
Value TDist2T(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  const double x = x_arg.value();
  const double df = std::floor(df_arg.value());
  if (x < 0.0 || df < 1.0 || df > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  // Two-tailed probability is `I_y(df/2, 1/2)` with y = df / (df + x*x).
  const double y = df / (df + x * x);
  const double r = stats::regularized_incomplete_beta(0.5 * df, 0.5, y);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// T.DIST.RT(x, deg_freedom) - right-tailed Student's t CDF, `1 - CDF(x)`.
// `x` is unrestricted; `df` must satisfy the same domain as T.DIST.
Value TDistRt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  const double x = x_arg.value();
  const double df = std::floor(df_arg.value());
  if (df < 1.0 || df > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  const double r = 1.0 - TDistCdf(x, df);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Newton-Raphson inverter for the Student's t CDF. Assumes the caller has
// already validated `0 < p < 1` and `df >= 1`. Uses Hill's Cornish-Fisher
// approximation for the initial guess, falling back to bisection when
// Newton oscillates in the steep tails.
double TInvCore(double p, double df) noexcept {
  if (p == 0.5) {
    return 0.0;
  }
  // Hill's approximation: starts from the standard-normal quantile and
  // applies two correction terms in df^-1 and df^-2. Good to ~1e-3 for
  // moderate df; for df = 1 it's off by ~10% near the tails but Newton
  // recovers in a handful of iterations.
  const double z = InverseStandardNormal(p);
  const double h = z * z;
  double x = z * (1.0 + (h + 1.0) / (4.0 * df) + ((5.0 * h + 16.0) * h + 3.0) / (96.0 * df * df));
  // Bisection bracket for fallback: the t quantile at p in (0, 1) always
  // lies inside [-1e6, 1e6] for df >= 1 (the mode is 0; tails are heavy
  // but finite at any p strictly inside the open unit interval).
  double lo = -1.0e6;
  double hi = 1.0e6;
  constexpr int kMaxIter = 100;
  constexpr double kTol = 1e-10;
  for (int i = 0; i < kMaxIter; ++i) {
    const double cdf = TDistCdf(x, df);
    const double err = cdf - p;
    // Tighten the bracket as Newton progresses; if Newton escapes it we
    // fall back to bisection below.
    if (err < 0.0) {
      lo = std::max(lo, x);
    } else {
      hi = std::min(hi, x);
    }
    if (std::abs(err) < kTol) {
      return x;
    }
    const double pdf = TDistPdf(x, df);
    if (pdf <= 0.0 || !std::isfinite(pdf)) {
      // Bisect.
      x = 0.5 * (lo + hi);
      continue;
    }
    double step = err / pdf;
    double x_new = x - step;
    if (x_new <= lo || x_new >= hi || !std::isfinite(x_new)) {
      // Step escaped the bracket; bisect instead. Keeps tail-steepness
      // cases (df = 1, p near 0 or 1) from diverging.
      x_new = 0.5 * (lo + hi);
    }
    if (std::abs(x_new - x) < kTol * std::max(1.0, std::abs(x))) {
      return x_new;
    }
    x = x_new;
  }
  return std::numeric_limits<double>::quiet_NaN();
}

// T.INV(probability, deg_freedom) - inverse of Student's t CDF.
// `p` must lie in the open unit interval; the median is returned exactly
// at `p == 0.5`. `df` floors toward -inf and must satisfy `df >= 1`.
Value TInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p_arg = coerce_to_number(args[0]);
  if (!p_arg) {
    return Value::error(p_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  const double p = p_arg.value();
  const double df = std::floor(df_arg.value());
  if (p <= 0.0 || p >= 1.0 || df < 1.0 || df > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  const double r = TInvCore(p, df);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// T.INV.2T(probability, deg_freedom) - inverse of the two-tailed Student's
// t. Solves `T.DIST.2T(x, df) == p`, equivalent to `T.INV(1 - p/2, df)`.
// Excel accepts `p == 1` here (yielding 0), so the upper bound is >.
Value TInv2T(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p_arg = coerce_to_number(args[0]);
  if (!p_arg) {
    return Value::error(p_arg.error());
  }
  auto df_arg = coerce_to_number(args[1]);
  if (!df_arg) {
    return Value::error(df_arg.error());
  }
  const double p = p_arg.value();
  const double df = std::floor(df_arg.value());
  if (p <= 0.0 || p > 1.0 || df < 1.0 || df > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  if (p == 1.0) {
    return Value::number(0.0);
  }
  const double r = TInvCore(1.0 - 0.5 * p, df);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Snedecor's F CDF at `x >= 0` with `(d1, d2)` degrees of freedom, via
// the regularized incomplete beta on `y = d1*x / (d1*x + d2)`.
double FDistCdf(double x, double d1, double d2) noexcept {
  if (x <= 0.0) {
    return 0.0;
  }
  const double y = (d1 * x) / (d1 * x + d2);
  return stats::regularized_incomplete_beta(0.5 * d1, 0.5 * d2, y);
}

// Snedecor's F PDF at `x > 0` with `(d1, d2)` degrees of freedom,
// computed in log space. The caller handles the x == 0 boundary
// (divergent for d1 < 2, equal to 1 for d1 == 2, and 0 for d1 > 2).
double FDistPdf(double x, double d1, double d2) noexcept {
  const double log_pdf = 0.5 * d1 * std::log(d1) + 0.5 * d2 * std::log(d2) + (0.5 * d1 - 1.0) * std::log(x) -
                         0.5 * (d1 + d2) * std::log(d1 * x + d2) + std::lgamma(0.5 * (d1 + d2)) -
                         std::lgamma(0.5 * d1) - std::lgamma(0.5 * d2);
  return std::exp(log_pdf);
}

// F.DIST(x, d1, d2, cumulative) - Snedecor's F distribution PDF or CDF.
// `d1` and `d2` floor toward -inf and must both satisfy `>= 1`. `x < 0`
// yields `#NUM!`. At `x == 0` the PDF is divergent for `d1 == 1` (Excel
// surfaces `#NUM!`), equals 1 for `d1 == 2`, and 0 for `d1 > 2`.
Value FDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  auto d1_arg = coerce_to_number(args[1]);
  if (!d1_arg) {
    return Value::error(d1_arg.error());
  }
  auto d2_arg = coerce_to_number(args[2]);
  if (!d2_arg) {
    return Value::error(d2_arg.error());
  }
  auto cum = coerce_to_bool(args[3]);
  if (!cum) {
    return Value::error(cum.error());
  }
  const double x = x_arg.value();
  const double d1 = std::floor(d1_arg.value());
  const double d2 = std::floor(d2_arg.value());
  if (x < 0.0 || d1 < 1.0 || d1 > kTFdfMax || d2 < 1.0 || d2 > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  double r;
  if (cum.value()) {
    r = FDistCdf(x, d1, d2);
  } else {
    if (x == 0.0) {
      if (d1 == 1.0) {
        return Value::error(ErrorCode::Num);
      }
      r = (d1 == 2.0) ? 1.0 : 0.0;
    } else {
      r = FDistPdf(x, d1, d2);
    }
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// F.DIST.RT(x, d1, d2) - right-tailed Snedecor's F CDF, `1 - CDF(x)`.
// Same domain as F.DIST; always `x >= 0`.
Value FDistRt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  auto d1_arg = coerce_to_number(args[1]);
  if (!d1_arg) {
    return Value::error(d1_arg.error());
  }
  auto d2_arg = coerce_to_number(args[2]);
  if (!d2_arg) {
    return Value::error(d2_arg.error());
  }
  const double x = x_arg.value();
  const double d1 = std::floor(d1_arg.value());
  const double d2 = std::floor(d2_arg.value());
  if (x < 0.0 || d1 < 1.0 || d1 > kTFdfMax || d2 < 1.0 || d2 > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  const double r = 1.0 - FDistCdf(x, d1, d2);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Newton-Raphson inverter for Snedecor's F CDF. Assumes `0 < p < 1` and
// both degrees of freedom `>= 1`. Uses a bisection sweep on [1e-6, 1e6]
// for the first few steps to land inside the high-curvature region, then
// switches to Newton. The F quantile at p close to 1 is steep and Newton
// alone can oscillate; the bisection warm-up sidesteps that entirely.
double FInvCore(double p, double d1, double d2) noexcept {
  double lo = 1e-10;
  double hi = 1e10;
  // Seed with ~20 bisection iterations to isolate the quantile to a
  // decently small interval. The F CDF is monotone on (0, inf), so
  // sign-checking the error at (lo, mid, hi) identifies the half-interval
  // containing the root.
  constexpr int kBisectIter = 40;
  for (int i = 0; i < kBisectIter; ++i) {
    const double mid = 0.5 * (lo + hi);
    const double cdf = FDistCdf(mid, d1, d2);
    if (cdf < p) {
      lo = mid;
    } else {
      hi = mid;
    }
    if ((hi - lo) < 1e-6 * std::max(1.0, mid)) {
      break;
    }
  }
  double x = 0.5 * (lo + hi);
  constexpr int kNewtonIter = 100;
  constexpr double kTol = 1e-10;
  for (int i = 0; i < kNewtonIter; ++i) {
    const double cdf = FDistCdf(x, d1, d2);
    const double err = cdf - p;
    if (err < 0.0) {
      lo = std::max(lo, x);
    } else {
      hi = std::min(hi, x);
    }
    if (std::abs(err) < kTol) {
      return x;
    }
    const double pdf = FDistPdf(x, d1, d2);
    if (pdf <= 0.0 || !std::isfinite(pdf)) {
      x = 0.5 * (lo + hi);
      continue;
    }
    double step = err / pdf;
    double x_new = x - step;
    if (x_new <= lo || x_new >= hi || !std::isfinite(x_new)) {
      x_new = 0.5 * (lo + hi);
    }
    if (std::abs(x_new - x) < kTol * std::max(1.0, std::abs(x))) {
      return x_new;
    }
    x = x_new;
  }
  return std::numeric_limits<double>::quiet_NaN();
}

// F.INV(probability, d1, d2) - inverse of Snedecor's F CDF. `p` must lie
// strictly inside the open unit interval.
Value FInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p_arg = coerce_to_number(args[0]);
  if (!p_arg) {
    return Value::error(p_arg.error());
  }
  auto d1_arg = coerce_to_number(args[1]);
  if (!d1_arg) {
    return Value::error(d1_arg.error());
  }
  auto d2_arg = coerce_to_number(args[2]);
  if (!d2_arg) {
    return Value::error(d2_arg.error());
  }
  const double p = p_arg.value();
  const double d1 = std::floor(d1_arg.value());
  const double d2 = std::floor(d2_arg.value());
  if (p <= 0.0 || p >= 1.0 || d1 < 1.0 || d1 > kTFdfMax || d2 < 1.0 || d2 > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  const double r = FInvCore(p, d1, d2);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// F.INV.RT(probability, d1, d2) - inverse of the right-tailed F CDF.
// Equivalent to `F.INV(1 - p, d1, d2)` modulo the `p == 1` edge case
// (Excel accepts `p == 1` and returns 0 for the right-tail variant,
// symmetric to CHISQ.INV.RT).
Value FInvRt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p_arg = coerce_to_number(args[0]);
  if (!p_arg) {
    return Value::error(p_arg.error());
  }
  auto d1_arg = coerce_to_number(args[1]);
  if (!d1_arg) {
    return Value::error(d1_arg.error());
  }
  auto d2_arg = coerce_to_number(args[2]);
  if (!d2_arg) {
    return Value::error(d2_arg.error());
  }
  const double p = p_arg.value();
  const double d1 = std::floor(d1_arg.value());
  const double d2 = std::floor(d2_arg.value());
  if (p <= 0.0 || p > 1.0 || d1 < 1.0 || d1 > kTFdfMax || d2 < 1.0 || d2 > kTFdfMax) {
    return Value::error(ErrorCode::Num);
  }
  if (p == 1.0) {
    return Value::number(0.0);
  }
  const double r = FInvCore(1.0 - p, d1, d2);
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

  // Probability-distribution family. Scalar-only: `accepts_ranges` is left
  // at its default `false`, and `propagate_errors` stays `true` so the
  // dispatcher short-circuits error arguments before the impl runs.
  registry.register_function(FunctionDef{"NORM.DIST", 4u, 4u, &NormDist});
  registry.register_function(FunctionDef{"NORM.S.DIST", 2u, 2u, &NormSDist});
  registry.register_function(FunctionDef{"NORM.INV", 3u, 3u, &NormInv});
  registry.register_function(FunctionDef{"NORM.S.INV", 1u, 1u, &NormSInv});
  registry.register_function(FunctionDef{"BINOM.DIST", 4u, 4u, &BinomDist});
  registry.register_function(FunctionDef{"POISSON.DIST", 3u, 3u, &PoissonDist});
  registry.register_function(FunctionDef{"EXPON.DIST", 3u, 3u, &ExponDist});

  // Chi-squared distribution family. All four are scalar-only (no range
  // expansion) and lean on `stats::p_gamma` / `stats::q_gamma` from
  // `eval/stats/special_functions.h`; the inverses close the loop with a
  // Newton-Raphson iteration seeded by Wilson-Hilferty.
  registry.register_function(FunctionDef{"CHISQ.DIST", 3u, 3u, &ChisqDist});
  registry.register_function(FunctionDef{"CHISQ.DIST.RT", 2u, 2u, &ChisqDistRt});
  registry.register_function(FunctionDef{"CHISQ.INV", 2u, 2u, &ChisqInv});
  registry.register_function(FunctionDef{"CHISQ.INV.RT", 2u, 2u, &ChisqInvRt});

  // Student's t distribution family. Scalar-only; all share
  // `stats::regularized_incomplete_beta` for the CDF surface. The inverses
  // use Hill's approximation for the initial guess with a bisection
  // fallback in the steep tails.
  registry.register_function(FunctionDef{"T.DIST", 3u, 3u, &TDist});
  registry.register_function(FunctionDef{"T.DIST.2T", 2u, 2u, &TDist2T});
  registry.register_function(FunctionDef{"T.DIST.RT", 2u, 2u, &TDistRt});
  registry.register_function(FunctionDef{"T.INV", 2u, 2u, &TInv});
  registry.register_function(FunctionDef{"T.INV.2T", 2u, 2u, &TInv2T});

  // Snedecor's F distribution family. Scalar-only; the inverses use a
  // bisection warm-up on [1e-10, 1e10] before switching to Newton to
  // avoid oscillation near the steep upper tail.
  registry.register_function(FunctionDef{"F.DIST", 4u, 4u, &FDist});
  registry.register_function(FunctionDef{"F.DIST.RT", 3u, 3u, &FDistRt});
  registry.register_function(FunctionDef{"F.INV", 3u, 3u, &FInv});
  registry.register_function(FunctionDef{"F.INV.RT", 3u, 3u, &FInvRt});

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
