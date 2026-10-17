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
#include "numeric_aggregate_kernels.h"
#include "utils/arena.h"
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

/// Double-double value `hi + lo` used to evaluate `y^n` past double precision.
struct DoubleDouble {
  double hi;
  double lo;
};

/// Exact product of two doubles via Veltkamp splitting (no fma dependency).
DoubleDouble two_product(double a, double b) {
  constexpr double kSplit = 134217729.0;  // 2^27 + 1
  const double hi = a * b;
  const double ta = kSplit * a;
  const double a_hi = ta - (ta - a);
  const double a_lo = a - a_hi;
  const double tb = kSplit * b;
  const double b_hi = tb - (tb - b);
  const double b_lo = b - b_hi;
  return {hi, ((a_hi * b_hi - hi) + a_hi * b_lo + a_lo * b_hi) + a_lo * b_lo};
}

DoubleDouble dd_mul(DoubleDouble a, DoubleDouble b) {
  DoubleDouble p = two_product(a.hi, b.hi);
  p.lo += a.hi * b.lo + a.lo * b.hi;
  const double hi = p.hi + p.lo;
  return {hi, p.lo - (hi - p.hi)};
}

/// Correctly rounded n-th root of a positive finite `value`: `pow` seeds the
/// root and one Newton step against `y^n` in double-double lands on the
/// nearest double, which is the root Excel returns.
double nth_root(double value, std::uint32_t n) {
  const double seed = std::pow(value, 1.0 / static_cast<double>(n));
  if (n == 1U || seed == 0.0 || !std::isfinite(seed)) {
    return seed;
  }
  DoubleDouble power{1.0, 0.0};
  DoubleDouble square{seed, 0.0};
  for (std::uint32_t k = n; k != 0U; k >>= 1U) {
    if ((k & 1U) != 0U) {
      power = dd_mul(power, square);
    }
    square = dd_mul(square, square);
  }
  const double residual = (value - power.hi) - power.lo;
  const double root = seed + seed * residual / (static_cast<double>(n) * value);
  return std::isfinite(root) ? root : seed;
}

}  // namespace

// GEOMEAN(value, ...) - geometric mean. Every numeric input must be
// strictly positive; a zero or negative value (including a range cell
// coerced via the numeric provenance rule) yields `#NUM!`. Computed as the
// n-th root of the product accumulated in argument order.
Value GeoMean(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto collected = collect_nonempty_direct_stats(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  const std::vector<double>& xs = collected.value();
  double product = 1.0;
  for (double x : xs) {
    if (x <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
    product *= x;
  }
  const double r = nth_root(product, static_cast<std::uint32_t>(xs.size()));
  return finite_number_result(r);
}

// HARMEAN(value, ...) - harmonic mean. Every input must be strictly
// positive; any value <= 0 yields `#NUM!`. `n / sum(1/x_i)`.
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
  const double variance = numeric_aggregate_kernels::dispersion_variance(xs, true);
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
