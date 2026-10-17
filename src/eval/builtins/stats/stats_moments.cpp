//
// Higher-moment descriptive-statistics family:
// GEOMEAN, HARMEAN, DEVSQ, AVEDEV, SKEW, SKEW.P, KURT, STANDARDIZE.
//
// All entries share the direct-argument rule of the VAR / STDEV family
// (`collect_direct_stats`): a direct Bool / numeric Text / omitted slot
// counts, non-numeric Text is #VALUE!, range-sourced non-Numbers are
// dropped by the dispatcher; Errors are short-circuited by the dispatcher. STANDARDIZE is the odd
// one out -- scalar-only, no range expansion -- but it is grouped here
// because it consumes the same `(x, mean, sd)` triple that backs every
// other moment-based function in this TU.
//
// Registration lives in the sibling `stats.cpp` translation unit; the
// shared helpers (`collect_numerics`, `compute_mean_ss`, `mean_of`,
// `centered_and_scaled`, `read_number_triple`, `finite_number_result`)
// are declared in `stats/stats_helpers.h`.

#include <cmath>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/builtins/stats/stats_helpers.h"
#include "eval/stats/flushed_ops.h"
#include "utils/arena.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

namespace {

Expected<std::vector<double>, ErrorCode> collect_nonempty_direct_stats(const Value* args, std::uint32_t arity) {
  auto collected = collect_direct_stats(args, arity);
  if (!collected) {
    return std::move(collected.error());
  }
  if (collected.value().empty()) {
    return ErrorCode::Num;
  }
  return std::move(collected.value());
}

}  // namespace

// GEOMEAN(value, ...) - geometric mean. Every numeric input must be
// strictly positive; a zero or negative value (including a range cell
// coerced via the numeric provenance rule) yields `#NUM!`. The product is
// kept as a mantissa in [1, 2) with a separate binary exponent E, so it can
// neither overflow nor underflow; the root is exp(log(m * 2^rem) * (1/n)) times
// 2^q, with q = trunc(E / n) and rem = E - q * n.
Value GeoMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto collected = collect_nonempty_direct_stats(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  const std::vector<double>& xs = collected.value();
  double mantissa = 1.0;
  long long exponent = 0;
  for (double x : xs) {
    if (x <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
    mantissa = flush_subnormal(mantissa * x);
    int e = 0;
    mantissa = std::frexp(mantissa, &e) * 2.0;
    exponent += e - 1;
  }
  const long long n = static_cast<long long>(xs.size());
  const long long q = exponent / n;
  const long long rem = exponent - q * n;
  const double scaled = std::ldexp(mantissa, static_cast<int>(rem));
  const double r =
      std::exp(flush_subnormal(std::log(scaled) * stats::flushed::guarded_div(1.0, static_cast<double>(n))));
  return finite_number_result(std::ldexp(r, static_cast<int>(q)));
}

// HARMEAN(value, ...) - harmonic mean. Every input must be strictly
// positive; any value <= 0 yields `#NUM!`. `1 / (sum(1/x_i) / n)`, the sum in
// argument order.
Value HarMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto collected = collect_nonempty_direct_stats(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  const std::vector<double>& xs = collected.value();
  double inv_sum = 0.0;
  for (double x : xs) {
    if (x <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
    inv_sum += stats::flushed::guarded_div(1.0, x);
  }
  const double r =
      stats::flushed::guarded_div(1.0, stats::flushed::guarded_div(inv_sum, static_cast<double>(xs.size())));
  return finite_number_result(r);
}

// DEVSQ(value, ...) - sum of squared deviations from the mean,
// `sum((x_i - mean)^2)`. Empty numeric slice yields 0 per Excel (the
// empty sum is zero; there is no division involved).
Value DevSq(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto collected = collect_direct_stats(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  const std::vector<double>& xs = collected.value();
  if (xs.empty()) {
    return Value::number(0.0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  return finite_number_result(ms.ss);
}

// AVEDEV(value, ...) - mean absolute deviation from the mean,
// `sum(|x_i - mean|) / n`. Empty numeric slice yields `#NUM!`.
Value AveDev(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto collected = collect_nonempty_direct_stats(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  const std::vector<double>& xs = collected.value();
  const double mean = mean_of(xs);
  double abs_sum = 0.0;
  for (double x : xs) {
    abs_sum += std::fabs(x - mean);
  }
  const double r = abs_sum / static_cast<double>(xs.size());
  return finite_number_result(r);
}

// Sum of `((x - mean) / scale)^3` over `xs`, shared by SKEW and SKEW.P.
static double cubed_z_sum(const CenteredScaled& stat) {
  double cubed_sum = 0.0;
  for (double x : stat.xs) {
    const double z = (x - stat.mean) / stat.scale;
    cubed_sum += z * z * z;
  }
  return cubed_sum;
}

// SKEW(value, ...) - sample skewness,
// `(n / ((n - 1)(n - 2))) * sum(((x_i - mean) / s)^3)` where `s` is the
// sample stdev. Requires at least 3 distinct non-zero deviations; fewer
// than 3 numeric inputs or zero sample variance yields `#DIV/0!`.
Value Skew(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto stat = centered_and_scaled(args, arity, 3u, 1.0);
  if (!stat) {
    return Value::error(stat.error());
  }
  const double cubed_sum = cubed_z_sum(stat.value());
  const double n = stat.value().n;
  const double coeff = n / ((n - 1.0) * (n - 2.0));
  return finite_number_result(coeff * cubed_sum);
}

// SKEW.P(value, ...) - population skewness,
// `(1 / n) * sum(((x_i - mean) / sigma)^3)` where `sigma` is the
// population stdev. A constant data set (sigma == 0) yields `#DIV/0!`;
// matching Excel, fewer than 3 numeric inputs also yields `#DIV/0!` so
// callers never see a degenerate near-symmetric zero.
Value SkewP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto stat = centered_and_scaled(args, arity, 3u, 0.0);
  if (!stat) {
    return Value::error(stat.error());
  }
  return finite_number_result(cubed_z_sum(stat.value()) / stat.value().n);
}

// KURT(value, ...) - excess kurtosis (Fisher's definition),
// `(n(n+1) / ((n-1)(n-2)(n-3))) * sum(((x - mean) / s)^4)
//   - 3(n-1)^2 / ((n-2)(n-3))`.
// Requires at least 4 numeric inputs (the cubic denominator collapses
// otherwise) and non-zero sample variance; both failures yield `#DIV/0!`.
Value Kurt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto collected = collect_direct_stats(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  const std::vector<double>& xs = collected.value();
  if (xs.size() < 4u) {
    return Value::error(ErrorCode::Div0);
  }
  const double n = static_cast<double>(xs.size());
  const double mean = mean_of(xs);
  const double variance = compute_mean_ss(xs).ss / (n - 1.0);
  if (!(variance > 0.0)) {
    return Value::error(ErrorCode::Div0);
  }
  const double sd = std::sqrt(variance);
  double quartic_sum = 0.0;
  for (double x : xs) {
    const double q = (x - mean) / sd;
    quartic_sum += (q * q) * (q * q);
  }
  const double coeff_a = (n * (n + 1.0)) / (((n - 1.0) * (n - 2.0)) * (n - 3.0));
  const double coeff_b = (3.0 * ((n - 1.0) * (n - 1.0))) / ((n - 2.0) * (n - 3.0));
  return finite_number_result(coeff_a * quartic_sum - coeff_b);
}

// STANDARDIZE(x, mean, standard_dev) - z-score, `(x - mean) / standard_dev`.
// Scalar-only: `accepts_ranges` stays false so the dispatcher coerces each
// argument directly. `standard_dev <= 0` yields `#NUM!`.
Value Standardize(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto input = read_number_triple(args, 0, 1, 2);
  if (!input) {
    return Value::error(input.error());
  }
  const double sd = input.value().third;
  if (sd <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return finite_number_result((input.value().first - input.value().second) / sd);
}

}  // namespace stats_detail
}  // namespace eval
}  // namespace formulon
