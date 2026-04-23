// Copyright 2026 libraz. Licensed under the MIT License.
//
// Auxiliary probability-distribution builtins split out of
// `stats_distributions.cpp` to keep that TU focused on the NORM / BINOM /
// POISSON / CHISQ / T / F core. This file holds:
//
//   * Confidence-interval half-widths (CONFIDENCE.NORM / CONFIDENCE.T)
//   * BINOM.INV / CRITBINOM
//   * FISHER / FISHERINV
//   * GAUSS / PHI
//   * NEGBINOM.DIST / NEGBINOMDIST
//   * BINOM.DIST.RANGE
//
// All entries share the same scalar-only, Excel-semantics conventions as the
// core distribution TU; see the header comment in
// `stats_distributions.cpp` for the full argument-coercion contract.

#include <cmath>
#include <cstdint>
#include <limits>

#include "eval/builtins/stats/stats_helpers.h"
#include "eval/coerce.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

// ---------------------------------------------------------------------------
// Confidence-interval half-widths
// ---------------------------------------------------------------------------

// CONFIDENCE / CONFIDENCE.NORM(alpha, stdev, size) - half-width of the
// (1 - alpha) confidence interval for a sample mean under the normal model:
//   z_{1 - alpha/2} * stdev / sqrt(size).
// `size` is truncated toward zero (Excel floors positive inputs; negative
// inputs are rejected outright). Domain: alpha in (0, 1), stdev > 0,
// size >= 1; any violation surfaces `#NUM!`.
Value ConfidenceNorm(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto alpha_arg = coerce_to_number(args[0]);
  if (!alpha_arg) {
    return Value::error(alpha_arg.error());
  }
  auto sd_arg = coerce_to_number(args[1]);
  if (!sd_arg) {
    return Value::error(sd_arg.error());
  }
  auto size_arg = coerce_to_number(args[2]);
  if (!size_arg) {
    return Value::error(size_arg.error());
  }
  const double alpha = alpha_arg.value();
  const double sd = sd_arg.value();
  const double size_raw = size_arg.value();
  // Reject negative / non-finite sizes before flooring. Excel truncates
  // toward zero for positives but does not silently convert negatives.
  if (std::isnan(size_raw) || std::isinf(size_raw) || size_raw < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double n = std::floor(size_raw);
  if (alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || n < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double z = InverseStandardNormal(1.0 - 0.5 * alpha);
  const double r = z * sd / std::sqrt(n);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// CONFIDENCE.T(alpha, stdev, size) - t-based confidence half-width:
//   t_{1 - alpha/2, size - 1} * stdev / sqrt(size).
// Domain matches CONFIDENCE.NORM plus `size >= 2` (df = size - 1 must be
// >= 1). `size == 1` surfaces `#DIV/0!` per Excel because df collapses to 0.
Value ConfidenceT(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto alpha_arg = coerce_to_number(args[0]);
  if (!alpha_arg) {
    return Value::error(alpha_arg.error());
  }
  auto sd_arg = coerce_to_number(args[1]);
  if (!sd_arg) {
    return Value::error(sd_arg.error());
  }
  auto size_arg = coerce_to_number(args[2]);
  if (!size_arg) {
    return Value::error(size_arg.error());
  }
  const double alpha = alpha_arg.value();
  const double sd = sd_arg.value();
  const double size_raw = size_arg.value();
  if (std::isnan(size_raw) || std::isinf(size_raw) || size_raw < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double n = std::floor(size_raw);
  if (alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || n < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  if (n == 1.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double t = TInvCore(1.0 - 0.5 * alpha, n - 1.0);
  const double r = t * sd / std::sqrt(n);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ---------------------------------------------------------------------------
// BINOM.INV / CRITBINOM
// ---------------------------------------------------------------------------

// BINOM.INV(trials, probability_s, alpha) - smallest integer k in [0, trials]
// with CDF(k) >= alpha. `trials` floors toward -inf. Domain:
// trials >= 0, prob in [0, 1], alpha in [0, 1]. Alpha == 0 short-circuits
// to 0; alpha beyond CDF(trials) by floating-point slop returns trials.
Value BinomInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto trials_arg = coerce_to_number(args[0]);
  if (!trials_arg) {
    return Value::error(trials_arg.error());
  }
  auto prob_arg = coerce_to_number(args[1]);
  if (!prob_arg) {
    return Value::error(prob_arg.error());
  }
  auto alpha_arg = coerce_to_number(args[2]);
  if (!alpha_arg) {
    return Value::error(alpha_arg.error());
  }
  const double n = std::floor(trials_arg.value());
  const double p = prob_arg.value();
  const double alpha = alpha_arg.value();
  if (n < 0.0 || p < 0.0 || p > 1.0 || alpha < 0.0 || alpha > 1.0) {
    return Value::error(ErrorCode::Num);
  }
  if (alpha == 0.0) {
    return Value::number(0.0);
  }
  const auto n_int = static_cast<std::uint64_t>(n);
  double cumulative = 0.0;
  for (std::uint64_t k = 0; k <= n_int; ++k) {
    cumulative += BinomPmf(static_cast<double>(k), n, p);
    if (cumulative >= alpha) {
      return Value::number(static_cast<double>(k));
    }
  }
  // Only reachable when alpha is extremely close to 1 and floating-point
  // roundoff keeps the cumulative sum a hair below it; the correct answer
  // is trials.
  return Value::number(n);
}

// ---------------------------------------------------------------------------
// FISHER / FISHERINV
// ---------------------------------------------------------------------------

// FISHER(x) - Fisher transformation: 0.5 * ln((1 + x) / (1 - x)).
// Domain: |x| < 1; |x| >= 1 surfaces `#NUM!` (the transformation diverges).
Value Fisher(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  const double x = x_arg.value();
  if (!(x > -1.0 && x < 1.0)) {
    return Value::error(ErrorCode::Num);
  }
  const double r = 0.5 * std::log((1.0 + x) / (1.0 - x));
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// FISHERINV(y) - inverse Fisher: (exp(2y) - 1) / (exp(2y) + 1). Defined for
// all finite y; returns `#NUM!` only if the computation overflows to +/-inf
// or produces NaN (i.e. y so large that exp(2y) already saturates).
Value FisherInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto y_arg = coerce_to_number(args[0]);
  if (!y_arg) {
    return Value::error(y_arg.error());
  }
  const double y = y_arg.value();
  const double e2y = std::exp(2.0 * y);
  const double r = (e2y - 1.0) / (e2y + 1.0);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ---------------------------------------------------------------------------
// GAUSS / PHI
// ---------------------------------------------------------------------------

// GAUSS(x) - probability that a standard-normal variable falls in [0, x].
// Equivalent to `NORM.S.DIST(x, TRUE) - 0.5`, i.e.
// `0.5 * erfc(-x / sqrt(2)) - 0.5`. Defined for all finite x.
Value Gauss(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  const double x = x_arg.value();
  const double r = 0.5 * std::erfc(-x / std::sqrt(2.0)) - 0.5;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// PHI(x) - standard-normal PDF: exp(-x^2 / 2) / sqrt(2 * pi). Defined for
// all finite x. Extreme |x| underflows to 0 without triggering `#NUM!`.
Value Phi(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = coerce_to_number(args[0]);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  const double x = x_arg.value();
  const double r = std::exp(-0.5 * x * x) / std::sqrt(2.0 * kStatsPi);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ---------------------------------------------------------------------------
// NEGBINOM.DIST / NEGBINOMDIST
// ---------------------------------------------------------------------------

// Log-space PMF of NegBinom(s, p) at f failures:
//   lgamma(f + s) - lgamma(f + 1) - lgamma(s) + s*log(p) + f*log(1-p).
// `p == 1` is handled separately (all mass at f = 0). `p == 0` must be
// rejected by the caller because s*log(p) diverges to -inf.
static double NegBinomLogPmf(double f, double s, double p) noexcept {
  if (p == 1.0) {
    return f == 0.0 ? 0.0 : -std::numeric_limits<double>::infinity();
  }
  return std::lgamma(f + s) - std::lgamma(f + 1.0) - std::lgamma(s) + s * std::log(p) + f * std::log1p(-p);
}

// NEGBINOM.DIST(number_f, number_s, probability_s, cumulative) - negative
// binomial PMF or CDF. `number_f` and `number_s` floor toward -inf; the
// domain requires f >= 0, s >= 1, and p in (0, 1] (p == 0 is rejected
// because the log-space PMF would collapse to -inf). CDF is an O(f) partial
// sum of PMFs from 0 to f.
Value NegBinomDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto f_arg = coerce_to_number(args[0]);
  if (!f_arg) {
    return Value::error(f_arg.error());
  }
  auto s_arg = coerce_to_number(args[1]);
  if (!s_arg) {
    return Value::error(s_arg.error());
  }
  auto prob_arg = coerce_to_number(args[2]);
  if (!prob_arg) {
    return Value::error(prob_arg.error());
  }
  auto cum = coerce_to_bool(args[3]);
  if (!cum) {
    return Value::error(cum.error());
  }
  const double f = std::floor(f_arg.value());
  const double s = std::floor(s_arg.value());
  const double p = prob_arg.value();
  if (f < 0.0 || s < 1.0 || p <= 0.0 || p > 1.0) {
    return Value::error(ErrorCode::Num);
  }
  // `p == 1` collapses the distribution to a point mass at f = 0. The CDF
  // is 1 everywhere for f >= 0; the PMF is 1 at f = 0 and 0 elsewhere.
  if (p == 1.0) {
    if (cum.value()) {
      return Value::number(1.0);
    }
    return Value::number(f == 0.0 ? 1.0 : 0.0);
  }
  double r;
  if (cum.value()) {
    r = 0.0;
    const auto f_int = static_cast<std::uint64_t>(f);
    for (std::uint64_t i = 0; i <= f_int; ++i) {
      r += std::exp(NegBinomLogPmf(static_cast<double>(i), s, p));
    }
  } else {
    r = std::exp(NegBinomLogPmf(f, s, p));
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// NEGBINOMDIST(number_f, number_s, probability_s) - pre-2010 3-arg
// spelling that always returns the PMF. Builds a synthetic 4-arg invocation
// with `cumulative = FALSE` and delegates to `NegBinomDist`.
Value NegBinomDistLegacy(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  Value synthetic[4] = {args[0], args[1], args[2], Value::boolean(false)};
  return NegBinomDist(synthetic, 4u, arena);
}

// ---------------------------------------------------------------------------
// BINOM.DIST.RANGE
// ---------------------------------------------------------------------------

// BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2]) - probability
// of obtaining between `number_s` and `number_s2` successes (inclusive).
// With 3 args, `number_s2` defaults to `number_s` (single-point PMF query).
// Domain: trials >= 0, prob in [0, 1], 0 <= number_s <= trials, and
// (when supplied) number_s <= number_s2 <= trials. Any violation surfaces
// `#NUM!`. Implementation sums PMFs from `number_s` to `number_s2`.
Value BinomDistRange(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto trials_arg = coerce_to_number(args[0]);
  if (!trials_arg) {
    return Value::error(trials_arg.error());
  }
  auto prob_arg = coerce_to_number(args[1]);
  if (!prob_arg) {
    return Value::error(prob_arg.error());
  }
  auto s_arg = coerce_to_number(args[2]);
  if (!s_arg) {
    return Value::error(s_arg.error());
  }
  const double n = std::floor(trials_arg.value());
  const double p = prob_arg.value();
  const double s1 = std::floor(s_arg.value());
  double s2 = s1;
  if (arity >= 4) {
    auto s2_arg = coerce_to_number(args[3]);
    if (!s2_arg) {
      return Value::error(s2_arg.error());
    }
    s2 = std::floor(s2_arg.value());
  }
  if (n < 0.0 || p < 0.0 || p > 1.0 || s1 < 0.0 || s1 > n || s2 < s1 || s2 > n) {
    return Value::error(ErrorCode::Num);
  }
  const auto s1_int = static_cast<std::uint64_t>(s1);
  const auto s2_int = static_cast<std::uint64_t>(s2);
  double r = 0.0;
  for (std::uint64_t k = s1_int; k <= s2_int; ++k) {
    r += BinomPmf(static_cast<double>(k), n, p);
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace stats_detail
}  // namespace eval
}  // namespace formulon
