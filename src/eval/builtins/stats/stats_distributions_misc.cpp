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
#include "eval/stats/special_functions.h"
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
  auto input = read_number_triple(args, 0, 1, 2);
  if (!input) {
    return Value::error(input.error());
  }
  const double alpha = input.value().first;
  const double sd = input.value().second;
  const double size_raw = input.value().third;
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
  return finite_number_result(z * sd / std::sqrt(n));
}

// CONFIDENCE.T(alpha, stdev, size) - t-based confidence half-width:
//   t_{1 - alpha/2, size - 1} * stdev / sqrt(size).
// Domain matches CONFIDENCE.NORM plus `size >= 2` (df = size - 1 must be
// >= 1). `size == 1` surfaces `#DIV/0!` per Excel because df collapses to 0.
Value ConfidenceT(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto input = read_number_triple(args, 0, 1, 2);
  if (!input) {
    return Value::error(input.error());
  }
  const double alpha = input.value().first;
  const double sd = input.value().second;
  const double size_raw = input.value().third;
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
  return finite_number_result(t * sd / std::sqrt(n));
}

// ---------------------------------------------------------------------------
// BINOM.INV / CRITBINOM
// ---------------------------------------------------------------------------

// Bisection for the smallest integer k in [0, n] with CDF(k) >= alpha,
// on the closed-form CDF. Used once a term-by-term walk would run past
// its ceiling. Each step halves the bracket, so the step count is bounded
// by the width of a double's exponent rather than by `n`.
//
// Maintains `CDF(lo) < alpha <= CDF(hi)`, with `hi` seeded at `n` because
// `CDF(n) == 1` and the caller has already rejected `alpha >= 1`.
//
// The shape check uses the balanced midpoint, which is the least
// resolvable pair the search can reach; refusing there keeps the answer
// inside the family's tolerance rather than off by a step or more.
static Value BinomInvBisect(double n, double p, double alpha) noexcept {
  if (!beta_shapes_are_resolvable(0.5 * n, 0.5 * n)) {
    return Value::error(ErrorCode::Num);
  }
  const double cdf_zero = BinomCdf(0.0, n, p);
  if (std::isnan(cdf_zero)) {
    return Value::error(ErrorCode::Num);
  }
  if (cdf_zero >= alpha) {
    return Value::number(0.0);
  }
  double lo = 0.0;
  double hi = n;
  while (hi - lo > 1.0) {
    const double mid = std::floor(lo + 0.5 * (hi - lo));
    if (!(mid > lo) || !(mid < hi)) {
      // No integer strictly inside the bracket is representable at this
      // magnitude; `hi` is the tightest answer available.
      break;
    }
    const double cdf = BinomCdf(mid, n, p);
    if (std::isnan(cdf)) {
      return Value::error(ErrorCode::Num);
    }
    if (cdf >= alpha) {
      hi = mid;
    } else {
      lo = mid;
    }
  }
  return finite_number_result(hi);
}

// BINOM.INV(trials, probability_s, alpha) - smallest integer k in [0, trials]
// with CDF(k) >= alpha. `trials` floors toward -inf. Domain:
// trials >= 0, prob in [0, 1], alpha in (0, 1) -- Excel rejects alpha == 0
// and alpha == 1 with #NUM!. Alpha very close to 1 may saturate the
// cumulative sum a hair below it due to floating-point roundoff; the correct
// answer is then trials.
//
// The walk stops at `kMaxCumulativeTerms` terms; the bound is on the
// terms actually consumed, not on `trials`, so a large-`trials` quantile
// that still lands early (a small `probability_s`, say) is answered by
// the walk. Past that the quantile is bisected on the closed-form CDF,
// which costs one evaluation per halving of the bracket.
Value BinomInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto input = read_number_triple(args, 0, 1, 2);
  if (!input) {
    return Value::error(input.error());
  }
  const double n = std::floor(input.value().first);
  const double p = input.value().second;
  const double alpha = input.value().third;
  if (n < 0.0 || p < 0.0 || p > 1.0 || alpha <= 0.0 || alpha >= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  // Stepped as a double rather than an integer counter: `trials` is an
  // unconstrained double, and a value past 2^64 would make the cast
  // undefined. Increments are exact well beyond the term ceiling.
  double cumulative = 0.0;
  for (double k = 0.0; k <= n; k += 1.0) {
    if (k >= kMaxCumulativeTerms) {
      return BinomInvBisect(n, p, alpha);
    }
    cumulative += BinomPmf(k, n, p);
    if (cumulative >= alpha) {
      return Value::number(k);
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
// all finite y; returns `#NUM!` only if the computation overflows to +/-inf
// or produces NaN (i.e. y so large that exp(2y) already saturates).
Value FisherInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto y_arg = read_number_arg(args, 0);
  if (!y_arg) {
    return Value::error(y_arg.error());
  }
  const double y = y_arg.value();
  const double e2y = std::exp(2.0 * y);
  return finite_number_result((e2y - 1.0) / (e2y + 1.0));
}

// ---------------------------------------------------------------------------
// GAUSS / PHI
// ---------------------------------------------------------------------------

// GAUSS(x) - probability that a standard-normal variable falls in [0, x].
// Equivalent to `NORM.S.DIST(x, TRUE) - 0.5`, i.e.
// `0.5 * erfc(-x / sqrt(2)) - 0.5`. Defined for all finite x.
Value Gauss(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_arg = read_number_arg(args, 0);
  if (!x_arg) {
    return Value::error(x_arg.error());
  }
  const double x = x_arg.value();
  return finite_number_result(0.5 * std::erfc(-x / std::sqrt(2.0)) - 0.5);
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

// Log-space PMF of NegBinom(s, p) at f failures:
//   lgamma(f + s) - lgamma(f + 1) - lgamma(s) + s*log(p) + f*log(1-p).
// Caller must reject p == 0 (because s*log(p) diverges to -inf) and p == 1
// (Excel surfaces #NUM! for the degenerate distribution).
static double NegBinomLogPmf(double f, double s, double p) noexcept {
  return std::lgamma(f + s) - std::lgamma(f + 1.0) - std::lgamma(s) + s * std::log(p) + f * std::log1p(-p);
}

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
  auto input = read_number_triple(args, 0, 1, 2);
  if (!input) {
    return Value::error(input.error());
  }
  auto cum = read_bool_arg(args, 3);
  if (!cum) {
    return Value::error(cum.error());
  }
  const double f = std::floor(input.value().first);
  const double s = std::floor(input.value().second);
  const double p = input.value().third;
  if (f < 0.0 || s < 1.0 || p <= 0.0 || p >= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  double r;
  if (cum.value()) {
    if (f >= kMaxCumulativeTerms) {
      if (!beta_shapes_are_resolvable(s, f + 1.0)) {
        return Value::error(ErrorCode::Num);
      }
      // P(F <= f) = I_p(s, f + 1).
      r = stats::regularized_incomplete_beta(s, f + 1.0, p);
    } else {
      r = 0.0;
      const auto f_int = static_cast<std::uint64_t>(f);
      for (std::uint64_t i = 0; i <= f_int; ++i) {
        r += std::exp(NegBinomLogPmf(static_cast<double>(i), s, p));
      }
    }
  } else {
    r = std::exp(NegBinomLogPmf(f, s, p));
  }
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
  const auto s1_int = static_cast<std::uint64_t>(s1);
  const auto s2_int = static_cast<std::uint64_t>(s2);
  double r = 0.0;
  for (std::uint64_t k = s1_int; k <= s2_int; ++k) {
    r += BinomPmf(static_cast<double>(k), n, p);
  }
  return finite_number_result(r);
}

}  // namespace stats_detail
}  // namespace eval
}  // namespace formulon
