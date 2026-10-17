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

#include "eval/stats/horner.h"
#include "utils/subnormal_flush.h"

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

namespace {

// Coefficient tables of the ERF/ERFC rational forms, highest degree first.
constexpr double kErfNumSmall[] = {
    0x1.932767df32911p-13, 0x1.d93a5813e2e53p-10, 0x1.6f289d165b309p-5, 0x1.201d8627a7d2cp-3, 0x1.1df5e70c7956bp+0,
};
constexpr double kErfDenSmall[] = {
    0x1.4f20a47119845p-16, 0x1.609751b5c02acp-11, 0x1.56de5da86ad9cp-7,
    0x1.79f85ba688040p-4,  0x1.d191aca8d2499p-2,  0x1.fada15778ece6p-1,
};
constexpr double kErfcNumMid[] = {
    0x1.decdbb07d9e2fp-16, 0x1.41cfbb83ed0e4p-11, 0x1.a016066459c2dp-8, 0x1.50b66746537eap-5, 0x1.76480f4d97fdap-3,
    0x1.27fc8b2d97507p-1,  0x1.4db9fcfb3c6e7p+0,  0x1.044811191d2cbp+1, 0x1.0223d9f70ed04p+1, 0x1.fffffffffff9fp-1,
};
constexpr double kErfcDenMid[] = {
    0x1.a8542411a723bp-15,
    0x1.1d32ae54a52b7p-10,
    0x1.71934f4e7b285p-7,
    0x1.2ca1bb9df5335p-4,
    0x1.516f295417ca8p-2,
    0x1.0f7f27e5dccd7p+0,
    0x1.3bc730d3c1c66p+1,
    0x1.05350002d8ffap+2,
    0x1.25449c578e4bdp+2,
    0x1.92929479238dbp+1,
    0x1p+0,
};
constexpr double kErfcNumBig[] = {
    0x1.4dc756d761350p-9, 0x1.6ca0f8ecfc798p-6, 0x1.0191f83634b97p-3, 0x1.d936f5138efd9p-2, 0x1.396f0fc8e5c9cp+0,
    0x1.26c409d384171p+1, 0x1.80fb2b5716d1ep+1, 0x1.3da266cd039d3p+1, 0x1.000aebf22de3bp+0,
};
constexpr double kErfcDenBig[] = {
    0x1.27cdbb941913dp-8, 0x1.4324d3ef0f238p-5, 0x1.cd27390103553p-3, 0x1.ad794d2e2c325p-1, 0x1.23c050f5fc4a3p+1,
    0x1.1e2d955b8fd9dp+2, 0x1.93da0f9daba61p+2, 0x1.850e3f0a8dba3p+2, 0x1.ce328ffeb796dp+1, 0x1p+0,
};

// Coefficient tables of the normal quantile: central ERFINV rationals, then
// the two tail rationals in sqrt(-ln q); highest degree first.
constexpr double kInvCentralNum[] = {
    -0x1.c6105b261ae5dp-12, 0x1.c034f2417c47bp-6, -0x1.5444e88a51ceep-2, 0x1.a0cbd20e4ce55p+0,
    -0x1.fce6e6f12a56bp+1,  0x1.492fc431b904dp+2, -0x1.b03599ad6a33ap+1, 0x1.c5bf891b4ef6bp-1,
};
constexpr double kInvCentralDen[] = {
    -0x1.5ae11dcadf003p-9, 0x1.38664c121c0dbp-4, -0x1.570c88781d342p-1, 0x1.59bb87268973cp+1,
    -0x1.708d997e809b9p+2, 0x1.ae761f165bea6p+2, -0x1.049a6a9451cabp+2, 0x1.0000000000000p+0,
};
constexpr double kInvMidNum[] = {
    0x1.08ad0740a6d4ap-2,  -0x1.32e5a7e827056p+1, -0x1.86cf9d01e16b2p+1, 0x1.ef9ffb45a3761p+4,
    -0x1.3aefb9a3ec59fp+5, 0x1.5a1d453954a8dp+2,  0x1.ff5f60cbe8bbap+2,  0x1.199fb95ef4e09p-2,
};
constexpr double kInvMidDen[] = {
    -0x1.ad7d1c36679b0p+3, 0x1.7ea3ac1d395ffp+5, -0x1.59120f7e33292p+5,
    -0x1.a4548a23b3e33p+2, 0x1.ccaeab51a6b87p+3, 0x1.0000000000000p+0,
};
constexpr double kInvOuterNum[] = {
    -0x1.1b8ef95ab8b0ep+11, 0x1.da1d6974302b0p+11, 0x1.ed8c72b7a0424p+12,  -0x1.ccdd3c3131b29p+12,
    -0x1.2c8f855be2320p+11, 0x1.f255c31ac92e5p+11, -0x1.7b1a5e1d9a10cp+10, 0x1.0a97c7664c331p+8,
    -0x1.6abfe9390c9e6p+4,  0x1.8242556d368b6p-1,
};
constexpr double kInvOuterDen[] = {
    -0x1.22d350f181b5ep+11, 0x1.9f372995afd52p+9,  0x1.75b4bf097172ep+13,  -0x1.5df83d4e27971p+12,
    -0x1.bb9141a0cb063p+12, 0x1.a1d6dc9f4d349p+12, -0x1.1a2a27132af74p+11, 0x1.76d842df4b9aap+8,
    -0x1.ec8f636f70044p+4,  0x1.0000000000000p+0,
};
constexpr double kTailNumNear[] = {
    0x1.9615ac0b7ace9p-11, 0x1.744eb6c45ec67p-6, 0x1.ef2abb9b85c37p-3, 0x1.453cc085375b2p+0,
    0x1.d2ecb1a3d02c4p+1,  0x1.713f71462256ap+2, 0x1.2857748cab19bp+2, 0x1.6c665fde9526ap+0,
};
constexpr double kTailDenNear[] = {
    0x1.20d3f686439e4p-30, 0x1.1f18cbfdf2728p-11, 0x1.f207a7eab17bfp-7, 0x1.2f5123394f040p-3,
    0x1.61292f23385c9p-1,  0x1.ad278e6526633p+0,  0x1.06cefbb46a449p+1, 0x1.0000000000000p+0,
};
constexpr double kTailNumFar[] = {
    0x1.afb74d693bf93p-23, 0x1.c6ec6cc59e02ap-16, 0x1.45c1908425345p-10, 0x1.b2b41193b4ee7p-6,
    0x1.2fad9315255cfp-2,  0x1.c8ea6461fa445p+0,  0x1.5daea6e875003p+2,  0x1.aa1b1c13ee526p+2,
};
constexpr double kTailDenFar[] = {
    0x1.269bff1f8c190p-49, 0x1.31446f740b9e0p-23, 0x1.35c2c496374bfp-16, 0x1.9c8bc979dc5d7p-11,
    0x1.e76f93215462ap-7,  0x1.186eb183443fbp-3,  0x1.331d34fc7d77fp-1,  0x1.0000000000000p+0,
};

double erfc_positive(double v) noexcept;

// Odd rational in v^2 below 0.5, `1 - erfc` beyond; `v` is non-negative.
double erf_positive(double v) noexcept {
  if (v < 0.5) {
    const double t = flush_subnormal(v * v);
    return flush_subnormal(flush_subnormal(horner(kErfNumSmall, t) * v) / horner(kErfDenSmall, t));
  }
  return 1.0 - erfc_positive(v);
}

// Rational in v times exp(-v^2), split at 4; `v` is non-negative.
double erfc_positive(double v) noexcept {
  if (v < 0.5) {
    return 1.0 - erf_positive(v);
  }
  if (v > 100.0) {
    return 0.0;
  }
  const double e = std::exp(-flush_subnormal(v * v));
  const double r = v > 4.0 ? flush_subnormal(horner(kErfcNumBig, v) / horner(kErfcDenBig, v))
                           : flush_subnormal(horner(kErfcNumMid, v) / horner(kErfcDenMid, v));
  return flush_subnormal(r * e);
}

// Tail rational in w = sqrt(-ln q) with q the smaller side; `a` is the lower-tail
// probability and `b` its complement.
double tail_quantile(double a, double b) noexcept {
  const bool upper = a > 0.5;
  const double w = std::sqrt(-std::log(upper ? b : a));
  double s = w - 5.0;
  double r;
  if (s > 0.0) {
    r = flush_subnormal(horner(kTailNumFar, s) / horner(kTailDenFar, s));
  } else {
    s = w - 1.6;
    r = flush_subnormal(horner(kTailNumNear, s) / horner(kTailDenNear, s));
  }
  return upper ? r : -r;
}

// Inverse error function in three regions of |y|, tail form beyond 0.9375.
double erfinv_core(double y) noexcept {
  if (y == 0.0) {
    return 0.0;
  }
  const double v = std::fabs(y);
  const double t = flush_subnormal(v * v);
  double r;
  if (t <= 0.5) {
    r = flush_subnormal(flush_subnormal(horner(kInvCentralNum, t) / horner(kInvCentralDen, t)) * v);
  } else if (v > 0.9375) {
    const double a = flush_subnormal((1.0 + v) * 0.5);
    const double b = flush_subnormal((1.0 - v) * 0.5);
    r = flush_subnormal(tail_quantile(a, b) * 0.7071067811865476);
  } else {
    const double w = std::sqrt(-std::log(1.0 - v));
    if (v > 0.85) {
      const double u = v - 0.87890625;
      r = flush_subnormal(horner(kInvOuterNum, u) / horner(kInvOuterDen, u));
    } else {
      r = flush_subnormal(horner(kInvMidNum, t) / horner(kInvMidDen, t));
    }
    r = flush_subnormal(r * w);
  }
  return y < 0.0 ? -r : r;
}

}  // namespace

double erf_excel(double x) noexcept {
  if (x == 0.0) {
    return 0.0;
  }
  const double r = erf_positive(std::fabs(x));
  return x < 0.0 ? -r : r;
}

double erfc_excel(double x) noexcept {
  if (x == 0.0) {
    return 1.0;
  }
  const double r = erfc_positive(std::fabs(x));
  return x < 0.0 ? (-r) + 2.0 : r;
}

double normal_cdf(double z) noexcept {
  return flush_subnormal(erfc_excel(flush_subnormal(z * -0.7071067811865476)) * 0.5);
}

double erfinv_excel(double y) noexcept {
  return erfinv_core(y);
}

double normal_quantile(double p) noexcept {
  if (p < 0.075 || p > 0.9375) {
    return tail_quantile(p, 1.0 - p);
  }
  const double y = flush_subnormal(p * 2.0) - 1.0;
  return flush_subnormal(erfinv_core(y) * 1.4142135623730951);
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon
