// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Higher-moment descriptive-statistics family:
// GEOMEAN, HARMEAN, DEVSQ, AVEDEV, SKEW, SKEW.P, KURT, STANDARDIZE.
//
// All entries share the skip-non-numeric rule of the MEDIAN / VAR family
// (only `Number` kind participates; Text / Bool / Blank are ignored;
// Errors are short-circuited by the dispatcher). STANDARDIZE is the odd
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
#include <vector>

#include "eval/builtins/stats/stats_helpers.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

// GEOMEAN(value, ...) - geometric mean. Every numeric input must be
// strictly positive; a zero or negative value (including a range cell
// coerced via the numeric provenance rule) yields `#NUM!`. Computed in
// log-space to avoid overflow for long data sets: exp(mean(ln(x_i))).
Value GeoMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
  return finite_number_result(r);
}

// HARMEAN(value, ...) - harmonic mean. Every input must be strictly
// positive; any value <= 0 yields `#NUM!`. `n / sum(1/x_i)`.
Value HarMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
  return finite_number_result(r);
}

// DEVSQ(value, ...) - sum of squared deviations from the mean,
// `sum((x_i - mean)^2)`. Empty numeric slice yields 0 per Excel (the
// empty sum is zero; there is no division involved).
Value DevSq(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::number(0.0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  return finite_number_result(ms.ss);
}

// AVEDEV(value, ...) - mean absolute deviation from the mean,
// `sum(|x_i - mean|) / n`. Empty numeric slice yields `#NUM!`.
Value AveDev(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
  return finite_number_result(r);
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
  double cubed_sum = 0.0;
  for (double x : stat.value().xs) {
    const double z = (x - stat.value().mean) / stat.value().scale;
    cubed_sum += z * z * z;
  }
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
  double cubed_sum = 0.0;
  for (double x : stat.value().xs) {
    const double z = (x - stat.value().mean) / stat.value().scale;
    cubed_sum += z * z * z;
  }
  return finite_number_result(cubed_sum / stat.value().n);
}

// KURT(value, ...) - excess kurtosis (Fisher's definition),
// `(n(n+1) / ((n-1)(n-2)(n-3))) * sum(((x - mean) / s)^4)
//   - 3(n-1)^2 / ((n-2)(n-3))`.
// Requires at least 4 numeric inputs (the cubic denominator collapses
// otherwise) and non-zero sample variance; both failures yield `#DIV/0!`.
Value Kurt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto stat = centered_and_scaled(args, arity, 4u, 1.0);
  if (!stat) {
    return Value::error(stat.error());
  }
  double quartic_sum = 0.0;
  for (double x : stat.value().xs) {
    const double z = (x - stat.value().mean) / stat.value().scale;
    quartic_sum += z * z * z * z;
  }
  const double n = stat.value().n;
  const double coeff_a = (n * (n + 1.0)) / ((n - 1.0) * (n - 2.0) * (n - 3.0));
  const double coeff_b = (3.0 * (n - 1.0) * (n - 1.0)) / ((n - 2.0) * (n - 3.0));
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
