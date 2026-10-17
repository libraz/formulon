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
#include "eval/stats/incomplete_beta.h"
#include "eval/stats/inverse_distributions.h"
#include "eval/stats/special_functions.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

// ---------------------------------------------------------------------------
// Confidence-interval half-widths
// ---------------------------------------------------------------------------

// Reads and validates the (alpha, stdev, size) triple shared by CONFIDENCE.NORM
// and CONFIDENCE.T; yields {alpha, stdev, floor(size)}.
static Expected<NumberTriple, ErrorCode> read_confidence_args(const Value* args) {
  auto input = read_number_triple(args, 0, 1, 2);
  if (!input) {
    return std::move(input.error());
  }
  const auto [alpha, sd, size_raw] = input.value();
  // Reject negative / non-finite sizes before flooring. Excel truncates
  // toward zero for positives but does not silently convert negatives.
  if (std::isnan(size_raw) || std::isinf(size_raw) || size_raw < 1.0) {
    return ErrorCode::Num;
  }
  const double n = std::floor(size_raw);
  if (alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || n < 1.0) {
    return ErrorCode::Num;
  }
  return NumberTriple{alpha, sd, n};
}

// CONFIDENCE / CONFIDENCE.NORM(alpha, stdev, size) - half-width of the
// (1 - alpha) confidence interval for a sample mean under the normal model:
//   z_{1 - alpha/2} * stdev / sqrt(size).
// `size` is truncated toward zero (Excel floors positive inputs; negative
// inputs are rejected outright). Domain: alpha in (0, 1), stdev > 0,
// size >= 1; any violation surfaces `#NUM!`.
Value ConfidenceNorm(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto input = read_confidence_args(args);
  if (!input) {
    return Value::error(input.error());
  }
  const auto [alpha, sd, n] = input.value();
  const double z = InverseStandardNormal(1.0 - 0.5 * alpha);
  return finite_number_result(z * sd / std::sqrt(n));
}

// CONFIDENCE.T(alpha, stdev, size) - t-based confidence half-width:
//   t_{1 - alpha/2, size - 1} * stdev / sqrt(size).
// Domain matches CONFIDENCE.NORM plus `size >= 2` (df = size - 1 must be
// >= 1). `size == 1` surfaces `#DIV/0!` per Excel because df collapses to 0.
Value ConfidenceT(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto input = read_confidence_args(args);
  if (!input) {
    return Value::error(input.error());
  }
  const auto [alpha, sd, n] = input.value();
  if (n == 1.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double t = TInvCore(1.0 - 0.5 * alpha, n - 1.0);
  return finite_number_result(t * sd / std::sqrt(n));
}

// ---------------------------------------------------------------------------
// BINOM.INV / CRITBINOM
// ---------------------------------------------------------------------------

// BINOM.INV(trials, probability_s, alpha) - smallest integer k in [0, trials]
// with CDF(k) >= alpha, found by bisection on the binomial CDF. `trials`
// floors toward -inf. Domain: trials >= 0, prob in (0, 1), alpha in (0, 1) --
// Excel rejects both endpoints of either probability with #NUM!.
Value BinomInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto input = read_number_triple(args, 0, 1, 2);
  if (!input) {
    return Value::error(input.error());
  }
  const double n = std::floor(input.value().first);
  const double p = input.value().second;
  const double alpha = input.value().third;
  if (n < 0.0 || p <= 0.0 || p >= 1.0 || alpha <= 0.0 || alpha >= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  return finite_number_result(stats::binom_inv_excel(n, p, alpha));
}

// ---------------------------------------------------------------------------
// FISHER / FISHERINV
// ---------------------------------------------------------------------------

// FISHER(x) - Fisher transformation: 0.5 * ln((1 + x) / (1 - x)).
// Domain: |x| < 1; |x| >= 1 surfaces `#NUM!` (the transformation diverges).
Value Fisher(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = read_number_arg(args, 0);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  const double x = x_arg.value();
  if (!(x > -1.0 && x < 1.0)) {
    return Value::error(ErrorCode::Num);
  }
  return finite_number_result(0.5 * std::log((1.0 + x) / (1.0 - x)));
}

// FISHERINV(y) - inverse Fisher: (exp(2y) - 1) / (exp(2y) + 1). Defined for
// all finite y; once exp(2y) overflows the result saturates to 1.
Value FisherInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto y_arg = read_number_arg(args, 0);
  if (!y_arg) {
    return Value::error(y_arg.error());
  }
  const double y = y_arg.value();
  const double e2y = std::exp(2.0 * y);
  if (std::isinf(e2y)) {
    return Value::number(1.0);
  }
  return finite_number_result((e2y - 1.0) / (e2y + 1.0));
}

// ---------------------------------------------------------------------------
// GAUSS / PHI
// ---------------------------------------------------------------------------

// GAUSS(x) - probability that a standard-normal variable falls in [0, x].
// `NORM.S.DIST(x, TRUE) - 0.5`, except that below 1e-15 the cancellation is
// avoided with `erf(x * -sqrt(1/2)) * -0.5`. Defined for all finite x.
Value Gauss(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = read_number_arg(args, 0);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  const double x = x_arg.value();
  const double r = std::fabs(x) > 1e-15
                       ? stats::normal_cdf(x) - 0.5
                       : flush_subnormal(stats::erf_excel(flush_subnormal(x * -0.7071067811865476)) * -0.5);
  return finite_number_result(r);
}

// PHI(x) - standard-normal PDF: exp(-x^2 / 2) / sqrt(2 * pi). Defined for
// all finite x. Extreme |x| underflows to 0 without triggering `#NUM!`.
Value Phi(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = read_number_arg(args, 0);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  const double x = x_arg.value();
  return finite_number_result(std::exp(-0.5 * x * x) / std::sqrt(2.0 * kStatsPi));
}

// ---------------------------------------------------------------------------
// NEGBINOM.DIST / NEGBINOMDIST
// ---------------------------------------------------------------------------

// NEGBINOM.DIST(number_f, number_s, probability_s, cumulative) - negative
// binomial PMF or CDF. `number_f` and `number_s` floor toward -inf; Excel
// requires f >= 0, s >= 1, and p strictly in (0, 1) -- both p == 0 and
// p == 1 (which collapse the distribution) surface #NUM!. CDF sums PMFs
// from 0 to f up to `kMaxCumulativeTerms` terms, then switches to the
// closed form so the cost never tracks the value of `number_f`. The usual
// shape here is a small `number_s` against a large `number_f`, which the
// closed form resolves exactly at any magnitude; a large `number_s` makes
// the pair balanced and is refused past the shape bound.
Value NegBinomDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto input = read_triple_cumulative(args);
  if (!input) {
    return Value::error(input.error());
  }
  const double f = std::floor(input.value().first);
  const double s = std::floor(input.value().second);
  const double p = input.value().third;
  if (f < 0.0 || s < 1.0 || p <= 0.0 || p >= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = input.value().cumulative ? stats::negbinom_cdf_excel(f, s, p) : stats::negbinom_pmf_excel(f, s, p);
  return finite_number_result(r);
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
// `#NUM!`. Implementation sums PMFs from `number_s` to `number_s2`; once
// the span would exceed `kMaxCumulativeTerms` terms it becomes the
// difference of two closed-form CDFs, so the cost stops tracking the
// width of the span. The summation is kept for short spans because a
// single-point query has to stay bit-identical to the matching `BinomPmf`.
Value BinomDistRange(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto input = read_number_triple(args, 0, 1, 2);
  if (!input) {
    return Value::error(input.error());
  }
  const double n = std::floor(input.value().first);
  const double p = input.value().second;
  const double s1 = std::floor(input.value().third);
  double s2 = s1;
  if (arity >= 4) {
    auto s2_arg = read_number_arg(args, 3);
    if (!s2_arg) {
      return Value::error(s2_arg.error());
    }
    s2 = std::floor(s2_arg.value());
  }
  if (n < 0.0 || p < 0.0 || p > 1.0 || s1 < 0.0 || s1 > n || s2 < s1 || s2 > n) {
    return Value::error(ErrorCode::Num);
  }
  if (s2 - s1 >= kMaxCumulativeTerms) {
    // P(s1 <= X <= s2) = CDF(s2) - CDF(s1 - 1); at s1 == 0 the lower term
    // is empty rather than CDF(-1), which is outside the support. Both
    // shape pairs have to be resolvable, since either one alone would
    // decide the difference.
    if (!beta_shapes_are_resolvable(n - s2, s2 + 1.0) || (s1 > 0.0 && !beta_shapes_are_resolvable(n - s1 + 1.0, s1))) {
      return Value::error(ErrorCode::Num);
    }
    const double upper = BinomCdf(s2, n, p);
    const double lower = (s1 == 0.0) ? 0.0 : BinomCdf(s1 - 1.0, n, p);
    return finite_number_result(upper - lower);
  }
  return finite_number_result(sum_pmf(s1, s2, [&](double k) { return BinomPmf(k, n, p); }));
}

}  // namespace stats_detail
}  // namespace eval
}  // namespace formulon
