// Copyright 2026 libraz. Licensed under the MIT License.
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

#include <cmath>
#include <limits>

namespace formulon {
namespace eval {
namespace stats {
namespace {

// Maximum number of iterations before we give up on convergence in either
// the series or continued-fraction branch. Numerical Recipes suggests 100;
// we round up to 200 for safety on inputs near the transition boundary,
// where convergence is slowest.
constexpr int kMaxIter = 200;

// Relative convergence threshold. Tightening this past ~1e-15 runs into
// IEEE-754 round-off and no longer buys accuracy.
constexpr double kEps = 1e-15;

// Lentz's floor for intermediate partial denominators. Prevents division
// by zero when a continued-fraction term vanishes exactly.
constexpr double kFpMin = 1e-300;

// Series expansion for `P(a, x)` valid for `x < a + 1`.
// γ(a, x) / Γ(a) = e^(-x) * x^a / Γ(a) * Σ_{n=0..∞} x^n / (a*(a+1)*...*(a+n))
// which is written iteratively as sum_{n} del_n where del_0 = 1/a and
// del_{n+1} = del_n * x / (a + n + 1). Early-out when |del| < |sum| * eps.
double p_gamma_series(double a, double x) noexcept {
  double ap = a;
  double sum = 1.0 / a;
  double del = sum;
  for (int n = 1; n <= kMaxIter; ++n) {
    ap += 1.0;
    del *= x / ap;
    sum += del;
    if (std::abs(del) < std::abs(sum) * kEps) {
      break;
    }
  }
  return sum * std::exp(-x + a * std::log(x) - std::lgamma(a));
}

// Lentz's modified continued fraction for `Q(a, x)` valid for `x >= a + 1`.
// Γ(a, x) / Γ(a) = e^(-x) * x^a / Γ(a) * (1 / (x + 1 - a - ...))
// evaluated via the standard Lentz recursion on partial numerators
// `an = -i * (i - a)` and partial denominators `b = x + 2i + 1 - a`.
double q_gamma_cf(double a, double x) noexcept {
  double b = x + 1.0 - a;
  double c = 1.0 / kFpMin;
  double d = 1.0 / b;
  double h = d;
  for (int i = 1; i <= kMaxIter; ++i) {
    const double an = -static_cast<double>(i) * (static_cast<double>(i) - a);
    b += 2.0;
    d = an * d + b;
    if (std::abs(d) < kFpMin) {
      d = kFpMin;
    }
    c = b + an / c;
    if (std::abs(c) < kFpMin) {
      c = kFpMin;
    }
    d = 1.0 / d;
    const double del = d * c;
    h *= del;
    if (std::abs(del - 1.0) < kEps) {
      break;
    }
  }
  return h * std::exp(-x + a * std::log(x) - std::lgamma(a));
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
double beta_cf(double a, double b, double x) noexcept {
  const double qab = a + b;
  const double qap = a + 1.0;
  const double qam = a - 1.0;
  double c = 1.0;
  double d = 1.0 - qab * x / qap;
  if (std::abs(d) < kFpMin) {
    d = kFpMin;
  }
  d = 1.0 / d;
  double h = d;
  for (int m = 1; m <= kMaxIter; ++m) {
    const double dm = static_cast<double>(m);
    const double m2 = 2.0 * dm;
    // Even step of Lentz's recursion.
    double aa = dm * (b - dm) * x / ((qam + m2) * (a + m2));
    d = 1.0 + aa * d;
    if (std::abs(d) < kFpMin) {
      d = kFpMin;
    }
    c = 1.0 + aa / c;
    if (std::abs(c) < kFpMin) {
      c = kFpMin;
    }
    d = 1.0 / d;
    h *= d * c;
    // Odd step.
    aa = -(a + dm) * (qab + dm) * x / ((a + m2) * (qap + m2));
    d = 1.0 + aa * d;
    if (std::abs(d) < kFpMin) {
      d = kFpMin;
    }
    c = 1.0 + aa / c;
    if (std::abs(c) < kFpMin) {
      c = kFpMin;
    }
    d = 1.0 / d;
    const double del = d * c;
    h *= del;
    if (std::abs(del - 1.0) < kEps) {
      break;
    }
  }
  return h;
}

}  // namespace

double p_gamma(double a, double x) noexcept {
  if (a <= 0.0 || x < 0.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (x == 0.0) {
    return 0.0;
  }
  if (x < a + 1.0) {
    return p_gamma_series(a, x);
  }
  // Continued-fraction branch computes Q; return 1 - Q.
  return 1.0 - q_gamma_cf(a, x);
}

double q_gamma(double a, double x) noexcept {
  if (a <= 0.0 || x < 0.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (x == 0.0) {
    return 1.0;
  }
  if (x < a + 1.0) {
    return 1.0 - p_gamma_series(a, x);
  }
  return q_gamma_cf(a, x);
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
  const double bt =
      std::exp(std::lgamma(a + b) - std::lgamma(a) - std::lgamma(b) + a * std::log(x) + b * std::log(1.0 - x));
  // Reflection point (a+1)/(a+b+2) is the approximate maximum of the
  // integrand; call `beta_cf` on whichever branch keeps x on the fast
  // side, and use the `I_x(a,b) = 1 - I_{1-x}(b,a)` identity otherwise.
  if (x < (a + 1.0) / (a + b + 2.0)) {
    return bt * beta_cf(a, b, x) / a;
  }
  return 1.0 - bt * beta_cf(b, a, 1.0 - x) / b;
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon
