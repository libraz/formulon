//
// Implementation of the regularized incomplete gamma functions
// `P(a, x)` / `Q(a, x) = 1 - P(a, x)` and the regularized incomplete beta
// function `I_x(a, b) = B(x; a, b) / B(a, b)` used by the Excel
// distribution family (CHISQ.*, T.*, F.*).
//
// Algorithms follow Numerical Recipes in C §6.2 (incomplete gamma) and
// §6.4 (incomplete beta):
//
//  - Gamma `x < a + 1`: power-series expansion of γ(a, x) / Γ(a).
//    Converges fast when `x` is small relative to `a`, badly when `x >> a`.
//  - Gamma `x >= a + 1`: Lentz's modified continued fraction for
//    Γ(a, x) / Γ(a). The dual of the power-series path; converges badly
//    when `x < a + 1`.
//  - Beta `x < (a + 1) / (a + b + 2)`: Lentz's continued fraction on the
//    direct integrand. Otherwise compute `1 - I_{1-x}(b, a)` via the same
//    CF on the reflected arguments; the reflection keeps every call on
//    the fast-converging branch.
//
// Switching at the boundary keeps each path on the convergent side of the
// split. The gamma paths share the prefactor
// `exp(-x + a*log(x) - lgamma(a))`; the beta path uses
// `exp(lgamma(a+b) - lgamma(a) - lgamma(b) + a*log(x) + b*log(1-x))`, both
// of which avoid overflow for the shape parameters encountered by the
// Excel distribution family (df up to 1e10).

#include "eval/stats/special_functions.h"

// Apple's <math.h> (and traditionally other libcs) only declares the
// reentrant lgamma_r/lgammaf_r/lgammal_r family when _REENTRANT is defined
// before the header is first included; without it the plain (non-reentrant,
// signgam-writing) lgamma is all that's visible.
#ifndef _REENTRANT
#define _REENTRANT
#endif

#include <algorithm>
#include <cmath>
#include <limits>

namespace formulon {
namespace eval {
namespace stats {
namespace {

// Absolute ceiling on iterations for any of the three recursions below,
// whatever the shape parameters ask for. The budgets grow with those
// parameters, and a shape large enough to want more work than this has
// already lost more precision in its prefactor than the extra iterations
// could recover -- so the honest result there is NaN rather than a long
// wait for a value nobody should trust. It also keeps the run time a
// property of the recursion rather than of the caller's arguments, and
// keeps the `int` conversions below well defined for any finite double.
constexpr int kMaxCfIterations = 2000000;

// Clamps a computed iteration budget into `[1000, kMaxCfIterations]`. The
// floor covers small shapes, where the budget formula would otherwise
// allow fewer steps than an ordinary case needs. Written so a NaN budget
// lands on the ceiling instead of converting out of range.
int clamp_iterations(double budget) noexcept {
  if (!(budget < static_cast<double>(kMaxCfIterations))) {
    return kMaxCfIterations;
  }
  return std::max(1000, static_cast<int>(budget));
}

// The gamma series and continued fraction need materially more work for a
// large shape parameter around their transition boundary: the terms fall off
// like exp(-n^2 / 2a), and the stopping rule below also waits out the slowly
// decaying tail. Never return a partial sum silently: callers receive NaN on
// a genuine non-convergence and convert it to Excel's #NUM!.
int max_gamma_iterations(double a) noexcept {
  return clamp_iterations(16.0 * std::sqrt(a) + 200.0);
}

// Iteration budget for the beta continued fraction, which follows a
// different law from the gamma pair and so needs its own.
//
// The slowest configuration is `x` at the reflection point with `a` close
// to `b`; there the step count tracks the cube root of `a + b` (about
// 4.3 * cbrt(a + b) for moderate shapes, easing to ~3.3 by a + b = 1e15).
// Twenty times the cube root holds roughly a five-fold margin over that
// across the whole range. Skewed shapes converge in about ten steps at
// any magnitude, so this budget only ever binds near `a == b`.
//
// The same "never return a partial sum silently" rule as the gamma pair
// applies: exhausting this budget yields NaN, not the running value.
int max_beta_iterations(double a, double b) noexcept {
  return clamp_iterations(20.0 * std::cbrt(a + b) + 200.0);
}

// Relative convergence threshold. Tightening this past ~1e-15 runs into
// IEEE-754 round-off and no longer buys accuracy.
constexpr double kEps = 1e-15;

// Lentz's floor for intermediate partial denominators. Prevents division
// by zero when a continued-fraction term vanishes exactly.
constexpr double kFpMin = 1e-300;

// Keeps a Lentz continued-fraction term away from zero.
double clamp_fp_min(double v) noexcept {
  return std::abs(v) < kFpMin ? kFpMin : v;
}

// Result of one series / continued-fraction evaluation. `converged` is
// what keeps a truncated run identifiable at the call site: every
// recursion here reports it, and no caller may read `value` without it.
struct CfResult {
  double value;
  bool converged;
};

// Stirling-series remainder `ln(n!) - ((n + 1/2) ln n - n + ln sqrt(2 pi))`,
// the building block of the saddle-point prefactors below. The series is used
// past 15; below it the definition is evaluated directly, where `ln(n!)` is
// small enough not to cancel.
double stirling_error(double n) noexcept {
  if (n <= 15.0) {
    constexpr double kLnSqrtTwoPi = 0.918938533204672741780329736406;
    return log_gamma(n + 1.0) - (n + 0.5) * std::log(n) + n - kLnSqrtTwoPi;
  }
  constexpr double kS0 = 1.0 / 12.0;
  constexpr double kS1 = 1.0 / 360.0;
  constexpr double kS2 = 1.0 / 1260.0;
  constexpr double kS3 = 1.0 / 1680.0;
  constexpr double kS4 = 1.0 / 1188.0;
  const double nn = n * n;
  if (n > 500.0) {
    return (kS0 - kS1 / nn) / n;
  }
  if (n > 80.0) {
    return (kS0 - (kS1 - kS2 / nn) / nn) / n;
  }
  if (n > 35.0) {
    return (kS0 - (kS1 - (kS2 - kS3 / nn) / nn) / nn) / n;
  }
  return (kS0 - (kS1 - (kS2 - (kS3 - kS4 / nn) / nn) / nn) / nn) / n;
}

// Deviance term `x ln(x / np) + np - x` without the cancellation a direct
// evaluation suffers when `x` is close to `np`.
double deviance_term(double x, double np) noexcept {
  if (std::abs(x - np) < 0.1 * (x + np)) {
    double v = (x - np) / (x + np);
    double s = (x - np) * v;
    if (std::abs(s) < std::numeric_limits<double>::min()) {
      return s;
    }
    double ej = 2.0 * x * v;
    v *= v;
    for (int j = 1; j < 1000; ++j) {
      ej *= v;
      const double s1 = s + ej / static_cast<double>(2 * j + 1);
      if (s1 == s) {
        return s1;
      }
      s = s1;
    }
  }
  return x * std::log(x / np) + np - x;
}

// `x^a e^-x / Gamma(a)`, the factor in front of both incomplete-gamma
// recursions. Past a shape of 15 the three log-space terms of the direct form
// are individually huge and cancel, leaving a relative error that grows with
// `a`; the saddle-point form `sqrt(a / 2 pi) exp(-stirling_error(a) -
// deviance_term(a, x))` keeps every term small and is accurate to a few ulps
// at any shape.
double gamma_prefactor(double a, double x) noexcept {
  constexpr double kSaddlePointShape = 15.0;
  if (a > kSaddlePointShape) {
    constexpr double kTwoPi = 6.283185307179586476925286766559;
    return std::sqrt(a / kTwoPi) * std::exp(-stirling_error(a) - deviance_term(a, x));
  }
  return std::exp(-x + a * std::log(x) - log_gamma(a));
}

// `x^a (1 - x)^b / B(a, b)`, the factor in front of the incomplete-beta
// continued fraction. The direct log-space form sums five terms that grow
// with the shapes and cancel, so its relative error grows like
// `eps * (a + b)`; the binomial saddle-point form
// `n * p^a q^b Gamma(n) / (Gamma(a) Gamma(b))` with `n = a + b` and
// `q = 1 - x` keeps every term small.
double beta_prefactor(double a, double b, double x) noexcept {
  const double n = a + b;
  constexpr double kSaddlePointShape = 30.0;
  if (n > kSaddlePointShape) {
    constexpr double kTwoPi = 6.283185307179586476925286766559;
    const double q = 1.0 - x;
    const double log_binomial =
        stirling_error(n) - stirling_error(a) - stirling_error(b) - deviance_term(a, n * x) - deviance_term(b, n * q);
    return std::sqrt(n / (kTwoPi * a * b)) * std::exp(log_binomial) * (a * b / n);
  }
  return std::exp(log_gamma(n) - log_gamma(a) - log_gamma(b) + a * std::log(x) + b * std::log(1.0 - x));
}

// Series expansion for `P(a, x)`:
// γ(a, x) / Γ(a) = e^(-x) * x^a / Γ(a) * Σ_{n=0..∞} x^n / (a*(a+1)*...*(a+n))
// written iteratively as sum_{n} del_n where del_0 = 1/a and
// del_{n+1} = del_n * x / (a + n + 1). Early-out once the geometric tail left
// behind, `del * x / (ap - x)`, is below `sum * eps`; stopping on `del` alone
// leaves a truncation error of order `eps / (1 - x/ap)`, which is large when
// `x` is close to a big `a`.
//
// For a large shape the series runs for ~sqrt(a) steps and a plain recurrence
// accumulates their round-off (about 1e-11 relative at a = 5e9). Each term is
// therefore carried as an unevaluated sum of two doubles, with the quotient
// `x / ap` split into its rounded value and exact remainder, and the sum is
// compensated; the cost is a few extra operations per step.
CfResult p_gamma_series(double a, double x) noexcept {
  double term_hi = 1.0 / a;
  double term_lo = std::fma(-term_hi, a, 1.0) / a;
  double sum = term_hi;
  double comp = term_lo;
  const int max_iterations = max_gamma_iterations(a);
  for (int n = 1; n <= max_iterations; ++n) {
    const double ap = a + static_cast<double>(n);
    const double ratio_hi = x / ap;
    const double ratio_lo = std::fma(-ratio_hi, ap, x) / ap;
    const double product = term_hi * ratio_hi;
    const double error = std::fma(term_hi, ratio_hi, -product) + (term_hi * ratio_lo + term_lo * ratio_hi);
    term_hi = product + error;
    term_lo = error - (term_hi - product);
    const double total = sum + term_hi;
    comp += (std::abs(sum) >= std::abs(term_hi) ? (sum - total) + term_hi : (term_hi - total) + sum) + term_lo;
    sum = total;
    if (ap > x && term_hi * x < (sum + comp) * kEps * (ap - x)) {
      return {(sum + comp) * gamma_prefactor(a, x), true};
    }
  }
  return {std::numeric_limits<double>::quiet_NaN(), false};
}

// Lentz's modified continued fraction for `Q(a, x)` valid for `x >= a + 1`.
// Γ(a, x) / Γ(a) = e^(-x) * x^a / Γ(a) * (1 / (x + 1 - a - ...))
// evaluated via the standard Lentz recursion on partial numerators
// `an = -i * (i - a)` and partial denominators `b = x + 2i + 1 - a`.
CfResult q_gamma_cf(double a, double x) noexcept {
  double b = x + 1.0 - a;
  double c = 1.0 / kFpMin;
  double d = 1.0 / b;
  double h = d;
  for (int i = 1; i <= max_gamma_iterations(a); ++i) {
    const double an = -static_cast<double>(i) * (static_cast<double>(i) - a);
    b += 2.0;
    d = an * d + b;
    d = clamp_fp_min(d);
    c = b + an / c;
    c = clamp_fp_min(c);
    d = 1.0 / d;
    const double del = d * c;
    h *= del;
    if (std::abs(del - 1.0) < kEps) {
      return {h * gamma_prefactor(a, x), true};
    }
  }
  return {std::numeric_limits<double>::quiet_NaN(), false};
}

// Lentz's modified continued fraction for the incomplete beta integral,
// evaluated at `(a, b, x)` with `x` on the fast-convergence branch (i.e.
// `x < (a + 1) / (a + b + 2)`). The public entry point calls this twice,
// once with `(a, b, x)` and once with `(b, a, 1 - x)`, then combines via
// the standard symmetry reflection.
//
// The recursion uses Numerical Recipes §6.4's partial-numerator pattern:
//   aa = m * (b - m) * x / ((a + 2m - 1) * (a + 2m))           (even step)
//   aa = -(a + m) * (a + b + m) * x / ((a + 2m) * (a + 2m + 1)) (odd step)
// interleaved inside a single iteration of Lentz's scheme.
CfResult beta_cf(double a, double b, double x) noexcept {
  const double qab = a + b;
  const double qap = a + 1.0;
  const double qam = a - 1.0;
  double c = 1.0;
  double d = 1.0 - qab * x / qap;
  d = clamp_fp_min(d);
  d = 1.0 / d;
  double h = d;
  const int max_iterations = max_beta_iterations(a, b);
  for (int m = 1; m <= max_iterations; ++m) {
    const double dm = static_cast<double>(m);
    const double m2 = 2.0 * dm;
    // Even step of Lentz's recursion.
    double aa = dm * (b - dm) * x / ((qam + m2) * (a + m2));
    d = 1.0 + aa * d;
    d = clamp_fp_min(d);
    c = 1.0 + aa / c;
    c = clamp_fp_min(c);
    d = 1.0 / d;
    h *= d * c;
    // Odd step.
    aa = -(a + dm) * (qab + dm) * x / ((a + m2) * (qap + m2));
    d = 1.0 + aa * d;
    d = clamp_fp_min(d);
    c = 1.0 + aa / c;
    c = clamp_fp_min(c);
    d = 1.0 / d;
    const double del = d * c;
    h *= del;
    if (std::abs(del - 1.0) < kEps) {
      return {h, true};
    }
  }
  return {std::numeric_limits<double>::quiet_NaN(), false};
}

// Largest `x` for which the series is preferred over the continued fraction.
// The classical split is `a + 1`; for a large shape the series stays accurate
// well beyond it (its terms carry no cancellation) while the continued
// fraction near the centre accumulates round-off, so the split moves out to
// three standard deviations, where the complement is still >= 1e-3.
double series_limit(double a) noexcept {
  return a < 100.0 ? a + 1.0 : a + 3.0 * std::sqrt(a) + 1.0;
}

}  // namespace

double log_gamma(double x) noexcept {
#if defined(_WIN32)
  // The Microsoft CRT does not expose `signgam` at all, so its `lgamma`
  // has no global side effect and is already thread-safe.
  return std::lgamma(x);
#else
  int sign = 0;
  return ::lgamma_r(x, &sign);
#endif
}

double p_gamma(double a, double x) noexcept {
  if (a <= 0.0 || x < 0.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (x == 0.0) {
    return 0.0;
  }
  if (x < series_limit(a)) {
    return p_gamma_series(a, x).value;
  }
  // Continued-fraction branch computes Q; return 1 - Q.
  const CfResult q = q_gamma_cf(a, x);
  return q.converged ? 1.0 - q.value : q.value;
}

double q_gamma(double a, double x) noexcept {
  if (a <= 0.0 || x < 0.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (x == 0.0) {
    return 1.0;
  }
  if (x < series_limit(a)) {
    const CfResult p = p_gamma_series(a, x);
    return p.converged ? 1.0 - p.value : p.value;
  }
  return q_gamma_cf(a, x).value;
}

double gamma_quantile(double a, double prob, bool upper) noexcept {
  constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();
  if (!(a > 0.0) || !(prob > 0.0 && prob < 1.0)) {
    return kNaN;
  }
  // The smaller tail carries the information: `1 - prob` is exact once
  // `prob > 0.5`, while a tail recovered as `1 - P` is not.
  if (prob > 0.5) {
    prob = 1.0 - prob;
    upper = !upper;
  }
  // Newton on `ln tail(e^t) - ln prob`, a smooth near-linear function of
  // `t = ln x` even when `prob` is far below double epsilon relative to 1.
  // A bracket on the root keeps every step that Newton would throw outside it
  // (or that a vanishing slope cannot define) on a bisection.
  const double sign = upper ? -1.0 : 1.0;
  constexpr double kMaxLogX = 700.0;
  constexpr double kMaxStep = 4.0;
  double lo = -kMaxLogX;
  double hi = kMaxLogX;
  double t = std::log(a);
  for (int i = 0; i < 400; ++i) {
    const double x = std::exp(t);
    const double tail = upper ? q_gamma(a, x) : p_gamma(a, x);
    if (std::isnan(tail)) {
      return kNaN;
    }
    // A tail that underflowed lies below the target on the increasing side.
    const double g = sign * (tail > 0.0 ? std::log(tail) - std::log(prob) : -kMaxLogX);
    if (g < 0.0) {
      lo = t;
    } else {
      hi = t;
    }
    const double slope = tail > 0.0 ? gamma_prefactor(a, x) / tail : 0.0;
    double t_new;
    if (slope > 0.0 && std::isfinite(slope)) {
      t_new = t - std::clamp(g / slope, -kMaxStep, kMaxStep);
    } else {
      t_new = g < 0.0 ? t + kMaxStep : t - kMaxStep;
    }
    if (!(t_new > lo && t_new < hi)) {
      t_new = (lo > -kMaxLogX && hi < kMaxLogX) ? 0.5 * (lo + hi) : std::clamp(t_new, lo, hi);
    }
    if (std::abs(t_new - t) <= 1e-15 * std::max(1.0, std::abs(t))) {
      return std::exp(t_new);
    }
    t = t_new;
  }
  return kNaN;
}

double regularized_incomplete_beta(double a, double b, double x) noexcept {
  if (a <= 0.0 || b <= 0.0 || x < 0.0 || x > 1.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (x == 0.0) {
    return 0.0;
  }
  if (x == 1.0) {
    return 1.0;
  }
  // Shared prefactor: (x^a * (1-x)^b) / (a * B(a, b)), computed in log
  // space to stay stable for large a / b. This is the factor outside the
  // continued fraction in Numerical Recipes eq. (6.4.5).
  const double lg_ab = log_gamma(a + b);
  const double lg_a = log_gamma(a);
  const double lg_b = log_gamma(b);
  const double a_log_x = a * std::log(x);
  const double b_log_1mx = b * std::log(1.0 - x);
  // Those five terms are individually huge for large shapes and cancel
  // almost entirely, so the round-off left in their sum is set by the
  // largest of them rather than by the result. Exponentiating a log that
  // is uncertain by `d` gives a prefactor with relative error about `d`,
  // so once `d` reaches order 1 the result is not a probability at all --
  // and it arrives with no other sign of trouble, because the continued
  // fraction itself converged. Refuse instead: the same rule as a
  // truncated recursion, applied to the other half of the computation.
  const double log_magnitude =
      std::abs(lg_ab) + std::abs(lg_a) + std::abs(lg_b) + std::abs(a_log_x) + std::abs(b_log_1mx);
  if (!(std::numeric_limits<double>::epsilon() * log_magnitude < 1.0)) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  const double bt = beta_prefactor(a, b, x);
  // Reflection point (a+1)/(a+b+2) is the approximate maximum of the
  // integrand; call `beta_cf` on whichever branch keeps x on the fast
  // side, and use the `I_x(a,b) = 1 - I_{1-x}(b,a)` identity otherwise.
  // A truncated continued fraction propagates as NaN rather than as a
  // partial value: past its budget the recursion is nowhere near the
  // answer, and the running value is not a probability at all.
  if (x < (a + 1.0) / (a + b + 2.0)) {
    const CfResult cf = beta_cf(a, b, x);
    return cf.converged ? bt * cf.value / a : cf.value;
  }
  const CfResult cf = beta_cf(b, a, 1.0 - x);
  return cf.converged ? 1.0 - bt * cf.value / b : cf.value;
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon
