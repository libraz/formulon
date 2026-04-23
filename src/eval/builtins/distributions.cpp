// Copyright 2026 libraz. Licensed under the MIT License.
//
// Extended probability-distribution catalog:
//   BETA.DIST / BETA.INV
//   GAMMA / GAMMALN / GAMMALN.PRECISE / GAMMA.DIST / GAMMA.INV
//   WEIBULL.DIST
//   LOGNORM.DIST / LOGNORM.INV
//   HYPGEOM.DIST
//
// Every entry is scalar-only (no range expansion). Arguments are coerced
// via `coerce_to_number` / `coerce_to_bool`; error-typed arguments
// short-circuit via the dispatcher before the impl runs. The CDF routines
// lean on `stats::p_gamma` and `stats::regularized_incomplete_beta` from
// `eval/stats/special_functions.h`; the inverses use a shared
// bracket-then-Newton helper defined locally in this TU.
//
// Numerical notes:
//   * Gamma / log-gamma: std::tgamma / std::lgamma handle the generic
//     case. tgamma is NaN/Inf at non-positive integers (poles); we
//     surface #NUM! explicitly before calling std::tgamma so the
//     behaviour matches Excel regardless of the platform's tgamma
//     NaN vs. Inf convention.
//   * GAMMA.INV reuses the chi-squared inverse identity
//       GAMMA.INV(p, alpha, beta) == beta * chisq_inv(p, 2*alpha) / 2,
//     avoiding a second Newton driver. (chisq_inv lives in stats.cpp;
//     we replicate just the solver here because the stats TU doesn't
//     expose it. The solver itself is a standard Wilson-Hilferty +
//     Newton and is small.)
//   * BETA.INV uses a shared bisection-then-Newton driver on the CDF
//     surface, bracketing on [0, 1].
//   * HYPGEOM.DIST's PMF uses lgamma-based log-combinations so the
//     finite factorials used by large (K, N) stay stable.

#include "eval/builtins/distributions.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <limits>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "eval/stats/special_functions.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Mathematical constant pi; matches std::acos(-1.0) on any IEEE-754
// platform. Used for the lognormal PDF normalisation.
constexpr double kDistPi = 3.14159265358979323846;

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

// Finalises a scalar result: non-finite becomes #NUM!, otherwise wraps in a
// Value::number. Mirrors the convention used throughout the builtins.
Value finalize(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Reads a required numeric argument. Non-finite coerced values surface
// as #NUM! so the impl bodies never have to re-check for NaN / Inf.
Expected<double, ErrorCode> read_number(const Value* args, std::uint32_t index) {
  auto coerced = coerce_to_number(args[index]);
  if (!coerced) {
    return coerced.error();
  }
  const double v = coerced.value();
  if (std::isnan(v) || std::isinf(v)) {
    return ErrorCode::Num;
  }
  return v;
}

// Reads an optional trailing numeric argument, falling back to
// `default_value` when arity <= index. Parallels `read_number` for
// error/non-finite handling.
Expected<double, ErrorCode> read_optional_number(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                 double default_value) {
  if (arity <= index) {
    return default_value;
  }
  return read_number(args, index);
}

// Shared bracket-then-Newton inverter for CDF surfaces on a half-open
// interval. The caller supplies `cdf(x)` and its derivative `pdf(x)`
// along with a bracket `[lo, hi]` known to contain the root (i.e.
// cdf(lo) <= p <= cdf(hi)). The driver runs a bisection warm-up to
// tighten the bracket, then switches to Newton with a bracket-escape
// safeguard that falls back to bisection if a Newton step would land
// outside `[lo, hi]`.
//
// Returns NaN on non-convergence so wrappers can surface #NUM!.
template <typename CdfFn, typename PdfFn>
double BracketThenNewton(CdfFn cdf, PdfFn pdf, double lo, double hi, double p) {
  constexpr int kBisectIter = 40;
  constexpr int kNewtonIter = 100;
  constexpr double kTol = 1e-12;

  // Bisection warm-up: tightens the bracket to ~1e-3 relative width before
  // Newton starts, which keeps tail-steep CDFs from oscillating.
  for (int i = 0; i < kBisectIter; ++i) {
    const double mid = 0.5 * (lo + hi);
    const double c = cdf(mid);
    if (std::isnan(c)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    if (c < p) {
      lo = mid;
    } else {
      hi = mid;
    }
    if ((hi - lo) < 1e-6 * std::max(1.0, std::abs(mid))) {
      break;
    }
  }
  double x = 0.5 * (lo + hi);
  for (int i = 0; i < kNewtonIter; ++i) {
    const double c = cdf(x);
    if (std::isnan(c)) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    const double err = c - p;
    if (err < 0.0) {
      lo = std::max(lo, x);
    } else {
      hi = std::min(hi, x);
    }
    if (std::abs(err) < kTol) {
      return x;
    }
    const double d = pdf(x);
    double x_new;
    if (d <= 0.0 || !std::isfinite(d)) {
      x_new = 0.5 * (lo + hi);
    } else {
      const double step = err / d;
      x_new = x - step;
      if (x_new <= lo || x_new >= hi || !std::isfinite(x_new)) {
        x_new = 0.5 * (lo + hi);
      }
    }
    if (std::abs(x_new - x) < kTol * std::max(1.0, std::abs(x))) {
      return x_new;
    }
    x = x_new;
  }
  return std::numeric_limits<double>::quiet_NaN();
}

// ---------------------------------------------------------------------------
// BETA family
// ---------------------------------------------------------------------------

// Beta PDF on [0, 1] in log space:
//   pdf(x) = exp( (alpha-1)*log(x) + (beta-1)*log(1-x)
//                - lgamma(alpha) - lgamma(beta) + lgamma(alpha+beta) )
// Callers must pre-reject x <= 0 and x >= 1; the 0^(a-1) cases have
// distribution-dependent boundary values that the wrappers handle directly.
double BetaPdfStd(double x, double alpha, double beta_shape) noexcept {
  const double log_pdf = (alpha - 1.0) * std::log(x) + (beta_shape - 1.0) * std::log(1.0 - x) - std::lgamma(alpha) -
                         std::lgamma(beta_shape) + std::lgamma(alpha + beta_shape);
  return std::exp(log_pdf);
}

// BETA.DIST(x, alpha, beta, cumulative, [A], [B]) - beta distribution on
// the optional support [A, B]. The standard [0, 1] form is recovered by
// omitting A / B (defaults A = 0, B = 1). For x == A or x == B the PDF
// hits boundary cases; for x strictly inside (A, B) we rescale via
//   y = (x - A) / (B - A)
// and delegate to the standard beta PDF/CDF, rescaling the PDF by the
// Jacobian 1 / (B - A). #NUM! on: alpha <= 0, beta <= 0, A >= B,
// x < A, x > B.
Value BetaDist(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto x_e = read_number(args, 0);
  if (!x_e) {
    return Value::error(x_e.error());
  }
  auto alpha_e = read_number(args, 1);
  if (!alpha_e) {
    return Value::error(alpha_e.error());
  }
  auto beta_e = read_number(args, 2);
  if (!beta_e) {
    return Value::error(beta_e.error());
  }
  auto cum_e = coerce_to_bool(args[3]);
  if (!cum_e) {
    return Value::error(cum_e.error());
  }
  auto a_e = read_optional_number(args, arity, 4, 0.0);
  if (!a_e) {
    return Value::error(a_e.error());
  }
  auto b_e = read_optional_number(args, arity, 5, 1.0);
  if (!b_e) {
    return Value::error(b_e.error());
  }
  const double x = x_e.value();
  const double alpha = alpha_e.value();
  const double beta_shape = beta_e.value();
  const double a = a_e.value();
  const double b = b_e.value();
  if (alpha <= 0.0 || beta_shape <= 0.0 || a >= b || x < a || x > b) {
    return Value::error(ErrorCode::Num);
  }
  const double span = b - a;
  const double y = (x - a) / span;
  if (cum_e.value()) {
    // CDF boundaries: the regularized incomplete beta already returns 0
    // at y == 0 and 1 at y == 1 exactly, so no special-case is needed.
    const double r = stats::regularized_incomplete_beta(alpha, beta_shape, y);
    return finalize(r);
  }
  // PDF boundaries at y == 0 and y == 1:
  //   alpha < 1 or beta < 1: PDF diverges -> #NUM!
  //   alpha == 1: PDF at y==0 is (beta / span); otherwise 0
  //   beta == 1: PDF at y==1 is (alpha / span); otherwise 0
  //   alpha > 1 and beta > 1: PDF is 0 at the boundary
  if (y == 0.0) {
    if (alpha < 1.0) {
      return Value::error(ErrorCode::Num);
    }
    if (alpha == 1.0) {
      return finalize(beta_shape / span);
    }
    return finalize(0.0);
  }
  if (y == 1.0) {
    if (beta_shape < 1.0) {
      return Value::error(ErrorCode::Num);
    }
    if (beta_shape == 1.0) {
      return finalize(alpha / span);
    }
    return finalize(0.0);
  }
  const double r = BetaPdfStd(y, alpha, beta_shape) / span;
  return finalize(r);
}

// BETA.INV(p, alpha, beta, [A], [B]) - inverse of BETA.DIST's CDF. Solves
//   I_y(alpha, beta) == p
// for y in (0, 1), then rescales via x = A + y * (B - A). #NUM! on
// p <= 0, p >= 1, alpha <= 0, beta <= 0, A >= B.
Value BetaInv(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto p_e = read_number(args, 0);
  if (!p_e) {
    return Value::error(p_e.error());
  }
  auto alpha_e = read_number(args, 1);
  if (!alpha_e) {
    return Value::error(alpha_e.error());
  }
  auto beta_e = read_number(args, 2);
  if (!beta_e) {
    return Value::error(beta_e.error());
  }
  auto a_e = read_optional_number(args, arity, 3, 0.0);
  if (!a_e) {
    return Value::error(a_e.error());
  }
  auto b_e = read_optional_number(args, arity, 4, 1.0);
  if (!b_e) {
    return Value::error(b_e.error());
  }
  const double p = p_e.value();
  const double alpha = alpha_e.value();
  const double beta_shape = beta_e.value();
  const double a = a_e.value();
  const double b = b_e.value();
  if (p <= 0.0 || p >= 1.0 || alpha <= 0.0 || beta_shape <= 0.0 || a >= b) {
    return Value::error(ErrorCode::Num);
  }
  // Invert on the [0, 1] support. The bracket-then-Newton driver stays
  // inside (0, 1) strictly because the Beta PDF is zero at the endpoints
  // for alpha/beta >= 1 and the safeguarded step never escapes the
  // bracket; for alpha < 1 or beta < 1 the PDF diverges at one endpoint
  // but bisection alone still converges.
  const auto cdf = [alpha, beta_shape](double y) { return stats::regularized_incomplete_beta(alpha, beta_shape, y); };
  const auto pdf = [alpha, beta_shape](double y) {
    if (y <= 0.0 || y >= 1.0) {
      return 0.0;
    }
    return BetaPdfStd(y, alpha, beta_shape);
  };
  const double y = BracketThenNewton(cdf, pdf, 1e-12, 1.0 - 1e-12, p);
  if (std::isnan(y)) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(a + y * (b - a));
}

// ---------------------------------------------------------------------------
// GAMMA family
// ---------------------------------------------------------------------------

// GAMMA(x) - Γ(x). Non-positive integers are poles (std::tgamma returns
// NaN or +Inf depending on platform); we reject them explicitly so the
// Excel-visible error code is consistent. At x == 0 Excel also surfaces
// #NUM! (pole from the left).
Value Gamma(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_e = read_number(args, 0);
  if (!x_e) {
    return Value::error(x_e.error());
  }
  const double x = x_e.value();
  // Non-positive integers are poles; also catch x == 0.
  if (x <= 0.0 && std::floor(x) == x) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::tgamma(x);
  return finalize(r);
}

// GAMMALN(x) / GAMMALN.PRECISE(x) - ln Γ(x). Both share this impl; the
// .PRECISE variant is Excel 2010+'s canonical spelling but behaves
// identically to the legacy name.
Value Gammaln(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_e = read_number(args, 0);
  if (!x_e) {
    return Value::error(x_e.error());
  }
  const double x = x_e.value();
  if (x <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(std::lgamma(x));
}

// Gamma PDF at x >= 0 with shape `alpha` and scale `beta`:
//   pdf(x) = 1 / (beta^alpha * Γ(alpha)) * x^(alpha-1) * exp(-x/beta)
// Evaluated in log space for stability at large alpha; callers reject
// x < 0, alpha <= 0, beta <= 0 up front.
double GammaPdf(double x, double alpha, double beta_scale) noexcept {
  const double log_pdf =
      -alpha * std::log(beta_scale) - std::lgamma(alpha) + (alpha - 1.0) * std::log(x) - x / beta_scale;
  return std::exp(log_pdf);
}

// GAMMA.DIST(x, alpha, beta, cumulative). #NUM! on x < 0, alpha <= 0,
// beta <= 0. At x == 0 the PDF is:
//   alpha < 1: divergent -> #NUM!
//   alpha == 1: 1 / beta (exponential rate)
//   alpha > 1: 0
// The CDF at x == 0 is 0 by definition.
Value GammaDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_e = read_number(args, 0);
  if (!x_e) {
    return Value::error(x_e.error());
  }
  auto alpha_e = read_number(args, 1);
  if (!alpha_e) {
    return Value::error(alpha_e.error());
  }
  auto beta_e = read_number(args, 2);
  if (!beta_e) {
    return Value::error(beta_e.error());
  }
  auto cum_e = coerce_to_bool(args[3]);
  if (!cum_e) {
    return Value::error(cum_e.error());
  }
  const double x = x_e.value();
  const double alpha = alpha_e.value();
  const double beta_scale = beta_e.value();
  if (x < 0.0 || alpha <= 0.0 || beta_scale <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (cum_e.value()) {
    if (x == 0.0) {
      return finalize(0.0);
    }
    return finalize(stats::p_gamma(alpha, x / beta_scale));
  }
  if (x == 0.0) {
    if (alpha < 1.0) {
      return Value::error(ErrorCode::Num);
    }
    if (alpha == 1.0) {
      return finalize(1.0 / beta_scale);
    }
    return finalize(0.0);
  }
  return finalize(GammaPdf(x, alpha, beta_scale));
}

// Inverse standard-normal CDF: rational Acklam approximation + two
// Halley-method polish iterations. Duplicated from stats.cpp because
// that TU keeps the helper internal; the replica is the same algorithm
// (same coefficients) and is used only as a Wilson-Hilferty initial
// guess for GAMMA.INV, not as a user-visible function.
double InverseStandardNormal(double p) {
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
  for (int i = 0; i < 2; ++i) {
    const double phi_cdf = 0.5 * std::erfc(-z / std::sqrt(2.0));
    const double phi_pdf = std::exp(-0.5 * z * z) / std::sqrt(2.0 * kDistPi);
    if (phi_pdf == 0.0) {
      break;
    }
    const double e = phi_cdf - p;
    const double u = e / phi_pdf;
    z -= u * (1.0 + 0.5 * z * u);
  }
  return z;
}

// GAMMA.INV(p, alpha, beta) - inverse of GAMMA.DIST's CDF. Mathematically
// equivalent to `beta * chisq_inv(p, 2*alpha) / 2`, i.e. a scaled
// chi-squared inverse; we implement the Wilson-Hilferty seed + Newton
// driver directly here (on the gamma surface) to avoid calling into
// stats.cpp for the helper. The two paths agree to ~1e-12.
Value GammaInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p_e = read_number(args, 0);
  if (!p_e) {
    return Value::error(p_e.error());
  }
  auto alpha_e = read_number(args, 1);
  if (!alpha_e) {
    return Value::error(alpha_e.error());
  }
  auto beta_e = read_number(args, 2);
  if (!beta_e) {
    return Value::error(beta_e.error());
  }
  const double p = p_e.value();
  const double alpha = alpha_e.value();
  const double beta_scale = beta_e.value();
  if (p < 0.0 || p >= 1.0 || alpha <= 0.0 || beta_scale <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (p == 0.0) {
    return finalize(0.0);
  }
  // Wilson-Hilferty transformation seeds Newton from within a few percent
  // of the true quantile: invert the normal approximation to the
  // cube-root-scaled chi-squared. `df = 2*alpha` is the chi-squared
  // equivalent; the scale factor `beta/2` recovers the gamma scale.
  const double df = 2.0 * alpha;
  const double h = 2.0 / (9.0 * df);
  const double z = InverseStandardNormal(p);
  const double cube_arg = 1.0 - h + z * std::sqrt(h);
  double chi2 = df * cube_arg * cube_arg * cube_arg;
  if (!(chi2 > 0.0)) {
    chi2 = 0.5 * df;  // Safeguard: negative / NaN -> fall back to mean.
  }
  double x = 0.5 * beta_scale * chi2;
  // Newton on GAMMA.DIST's CDF: f(x) = p_gamma(alpha, x/beta) - p.
  // f'(x) = PDF(x) = GammaPdf(x, alpha, beta).
  constexpr int kMaxIter = 100;
  constexpr double kTol = 1e-12;
  for (int i = 0; i < kMaxIter; ++i) {
    const double cdf = stats::p_gamma(alpha, x / beta_scale);
    if (std::isnan(cdf)) {
      return Value::error(ErrorCode::Num);
    }
    const double pdf = GammaPdf(x, alpha, beta_scale);
    if (pdf <= 0.0 || !std::isfinite(pdf)) {
      break;
    }
    double step = (cdf - p) / pdf;
    double x_new = x - step;
    // Safeguard: Newton can overshoot into x <= 0 on steep tails; halve
    // the step until we land inside the positive reals.
    while (x_new <= 0.0) {
      step *= 0.5;
      x_new = x - step;
      if (std::abs(step) < kTol) {
        x_new = 0.5 * x;
        break;
      }
    }
    if (std::abs(x_new - x) < kTol * std::max(1.0, std::abs(x))) {
      return finalize(x_new);
    }
    x = x_new;
  }
  // Final convergence check: even if the Newton loop ran out of
  // iterations, `x` is usually close; surface #NUM! only if the CDF
  // residual is large.
  const double residual = std::abs(stats::p_gamma(alpha, x / beta_scale) - p);
  if (residual > 1e-6) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(x);
}

// ---------------------------------------------------------------------------
// WEIBULL.DIST
// ---------------------------------------------------------------------------

// WEIBULL.DIST(x, alpha, beta, cumulative). The Weibull distribution with
// shape `alpha` and scale `beta`:
//   CDF: 1 - exp(-(x/beta)^alpha)
//   PDF: (alpha/beta) * (x/beta)^(alpha-1) * exp(-(x/beta)^alpha)
// #NUM! on x < 0, alpha <= 0, beta <= 0.
//
// Boundary at x == 0: Mac Excel 365 returns exactly 0 for the PDF at
// x == 0 regardless of alpha -- including the `alpha == 1` exponential
// case (where the mathematical limit is 1/beta) and the `alpha < 1`
// divergent case (where the mathematical limit is +Inf). Excel's
// contract for this entry is "density at the boundary is zero"; verified
// against the golden oracle. The CDF at x == 0 is 0 as expected.
Value WeibullDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_e = read_number(args, 0);
  if (!x_e) {
    return Value::error(x_e.error());
  }
  auto alpha_e = read_number(args, 1);
  if (!alpha_e) {
    return Value::error(alpha_e.error());
  }
  auto beta_e = read_number(args, 2);
  if (!beta_e) {
    return Value::error(beta_e.error());
  }
  auto cum_e = coerce_to_bool(args[3]);
  if (!cum_e) {
    return Value::error(cum_e.error());
  }
  const double x = x_e.value();
  const double alpha = alpha_e.value();
  const double beta_scale = beta_e.value();
  if (x < 0.0 || alpha <= 0.0 || beta_scale <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double t = x / beta_scale;
  const double t_pow = std::pow(t, alpha);
  if (cum_e.value()) {
    return finalize(1.0 - std::exp(-t_pow));
  }
  // PDF boundary: Mac Excel surfaces exactly 0 at x == 0 regardless of
  // alpha, overriding the mathematical limits of 1/beta (alpha == 1)
  // and +Inf (alpha < 1).
  if (x == 0.0) {
    return finalize(0.0);
  }
  const double r = (alpha / beta_scale) * std::pow(t, alpha - 1.0) * std::exp(-t_pow);
  return finalize(r);
}

// ---------------------------------------------------------------------------
// LOGNORM family
// ---------------------------------------------------------------------------

// LOGNORM.DIST(x, mean, sd, cumulative). The log-normal distribution
// parameterised by (mean, sd) on the log scale:
//   CDF: Phi((ln x - mean) / sd)
//   PDF: 1 / (x * sd * sqrt(2*pi)) * exp(-(ln x - mean)^2 / (2 sd^2))
// #NUM! on x <= 0, sd <= 0.
Value LognormDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_e = read_number(args, 0);
  if (!x_e) {
    return Value::error(x_e.error());
  }
  auto mean_e = read_number(args, 1);
  if (!mean_e) {
    return Value::error(mean_e.error());
  }
  auto sd_e = read_number(args, 2);
  if (!sd_e) {
    return Value::error(sd_e.error());
  }
  auto cum_e = coerce_to_bool(args[3]);
  if (!cum_e) {
    return Value::error(cum_e.error());
  }
  const double x = x_e.value();
  const double mean = mean_e.value();
  const double sd = sd_e.value();
  if (x <= 0.0 || sd <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double z = (std::log(x) - mean) / sd;
  if (cum_e.value()) {
    // P(X <= x) = Phi(z) = 0.5 * erfc(-z / sqrt(2)).
    return finalize(0.5 * std::erfc(-z / std::sqrt(2.0)));
  }
  const double r = std::exp(-0.5 * z * z) / (x * sd * std::sqrt(2.0 * kDistPi));
  return finalize(r);
}

// LOGNORM.INV(p, mean, sd) - inverse log-normal CDF. Computed as
//   exp(mean + sd * InverseStandardNormal(p))
// #NUM! on p <= 0, p >= 1, sd <= 0.
Value LognormInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto p_e = read_number(args, 0);
  if (!p_e) {
    return Value::error(p_e.error());
  }
  auto mean_e = read_number(args, 1);
  if (!mean_e) {
    return Value::error(mean_e.error());
  }
  auto sd_e = read_number(args, 2);
  if (!sd_e) {
    return Value::error(sd_e.error());
  }
  const double p = p_e.value();
  const double mean = mean_e.value();
  const double sd = sd_e.value();
  if (p <= 0.0 || p >= 1.0 || sd <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double z = InverseStandardNormal(p);
  return finalize(std::exp(mean + sd * z));
}

// ---------------------------------------------------------------------------
// HYPGEOM.DIST
// ---------------------------------------------------------------------------

// Log-combination log(C(n, k)) = lgamma(n+1) - lgamma(k+1) - lgamma(n-k+1).
// Callers pre-validate 0 <= k <= n so the lgamma arguments stay positive.
double LogCombination(double n, double k) noexcept {
  return std::lgamma(n + 1.0) - std::lgamma(k + 1.0) - std::lgamma(n - k + 1.0);
}

// Hypergeometric PMF at k successes drawn from a population of N with K
// successes, sampling n without replacement:
//   PMF(k) = C(K, k) * C(N-K, n-k) / C(N, n)
// Computed via log-combinations to avoid overflow of the individual
// binomial coefficients for even moderately large N.
double HypgeomPmf(double k, double n, double K, double N) noexcept {
  const double log_pmf = LogCombination(K, k) + LogCombination(N - K, n - k) - LogCombination(N, n);
  return std::exp(log_pmf);
}

// HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative).
// Maps to the hypergeometric PMF / CDF with:
//   k = sample_s, n = number_sample, K = population_s, N = number_pop.
// Excel floors all four counts toward -inf.
//
// Domain split (verified against the Mac Excel 365 oracle):
//   * Any negative count (k < 0, n < 0, K < 0, N < 0) surfaces `#NUM!`.
//   * Malformed arrangements where `n > N` or `K > N` surface `#NUM!`.
//   * Otherwise, when `k` is outside the support
//     `[max(0, n + K - N), min(n, K)]` but non-negative, the PMF is 0
//     and the CDF is 0 (below support) or 1 (above support). Excel
//     treats the infeasible-but-non-negative branch as measure-zero
//     rather than malformed input.
Value HypgeomDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto k_e = read_number(args, 0);
  if (!k_e) {
    return Value::error(k_e.error());
  }
  auto n_e = read_number(args, 1);
  if (!n_e) {
    return Value::error(n_e.error());
  }
  auto big_k_e = read_number(args, 2);
  if (!big_k_e) {
    return Value::error(big_k_e.error());
  }
  auto big_n_e = read_number(args, 3);
  if (!big_n_e) {
    return Value::error(big_n_e.error());
  }
  auto cum_e = coerce_to_bool(args[4]);
  if (!cum_e) {
    return Value::error(cum_e.error());
  }
  const double k = std::floor(k_e.value());
  const double n = std::floor(n_e.value());
  const double big_k = std::floor(big_k_e.value());
  const double big_n = std::floor(big_n_e.value());
  // Malformed-input check. Mac Excel rejects any negative count
  // (including `k < 0`) with #NUM!, alongside the containment checks
  // `n <= N` and `K <= N`. Per-oracle: `=HYPGEOM.DIST(-1, 4, 8, 20, FALSE)`
  // surfaces `#NUM!`, even though k = -1 is trivially outside the
  // support. k < 0 is treated as a malformed argument, not an
  // infeasible support point.
  if (k < 0.0 || n < 0.0 || big_k < 0.0 || big_n < 0.0 || n > big_n || big_k > big_n) {
    return Value::error(ErrorCode::Num);
  }
  // Support window. k >= 0 here; the infeasible branch covers
  // k > min(n, K) and k < n + K - N (when that's positive).
  const double k_min = std::max(0.0, n + big_k - big_n);
  const double k_max = std::min(n, big_k);
  if (cum_e.value()) {
    // CDF is 0 below the support and 1 at / above k_max.
    if (k < k_min) {
      return finalize(0.0);
    }
    if (k >= k_max) {
      return finalize(1.0);
    }
    double acc = 0.0;
    const auto start = static_cast<std::uint64_t>(k_min);
    const auto stop = static_cast<std::uint64_t>(k);
    for (std::uint64_t i = start; i <= stop; ++i) {
      acc += HypgeomPmf(static_cast<double>(i), n, big_k, big_n);
    }
    return finalize(acc);
  }
  // PMF branch: 0 outside the support, otherwise the log-combination form.
  if (k < k_min || k > k_max) {
    return finalize(0.0);
  }
  return finalize(HypgeomPmf(k, n, big_k, big_n));
}

// ---------------------------------------------------------------------------
// Legacy (pre-2010) distribution spellings.
//
// LOGNORMDIST / HYPGEOMDIST / BETADIST predate the 2010 overhaul that
// introduced LOGNORM.DIST / HYPGEOM.DIST / BETA.DIST with explicit
// cumulative flags. The legacy forms inject the fixed cumulative bit
// and forward to the canonical impl so the math stays in one place.
// BETAINV and LOGINV / GAMMA* / WEIBULL have signature-identical
// .NEW forms and therefore need no wrapper (they register the same
// pointer under the legacy name).

// LOGNORMDIST(x, mean, sd) - 3-arg legacy: cumulative CDF only.
Value LognormDistLegacy(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  Value synthetic[4] = {args[0], args[1], args[2], Value::boolean(true)};
  return LognormDist(synthetic, 4u, arena);
}

// HYPGEOMDIST(sample_s, number_sample, population_s, number_pop) - 4-arg
// legacy: non-cumulative PMF only.
Value HypgeomDistLegacy(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  Value synthetic[5] = {args[0], args[1], args[2], args[3], Value::boolean(false)};
  return HypgeomDist(synthetic, 5u, arena);
}

// BETADIST(x, alpha, beta, [A], [B]) - 3-5 arg legacy: cumulative CDF
// only. Injects `cumulative = TRUE` at slot 3 before calling BETA.DIST,
// then forwards the optional A / B bounds (which default to [0, 1]).
Value BetaDistLegacy(const Value* args, std::uint32_t arity, Arena& arena) {
  Value synthetic[6] = {args[0], args[1], args[2], Value::boolean(true), Value::blank(), Value::blank()};
  std::uint32_t new_arity = 4u;
  if (arity >= 4u) {
    synthetic[4] = args[3];
    new_arity = 5u;
  }
  if (arity >= 5u) {
    synthetic[5] = args[4];
    new_arity = 6u;
  }
  return BetaDist(synthetic, new_arity, arena);
}

}  // namespace

void register_distribution_builtins(FunctionRegistry& registry) {
  // Scalar-only: every entry below coerces its args through `coerce_to_number`
  // / `coerce_to_bool`, so `accepts_ranges` stays at the default `false`.
  registry.register_function(FunctionDef{"BETA.DIST", 4u, 6u, &BetaDist});
  registry.register_function(FunctionDef{"BETA.INV", 3u, 5u, &BetaInv});

  registry.register_function(FunctionDef{"GAMMA", 1u, 1u, &Gamma});
  // GAMMALN is the pre-2010 name; GAMMALN.PRECISE is Excel 2010+'s
  // canonical spelling. Both use the same impl and are exposed under
  // both names so either formula parses.
  registry.register_function(FunctionDef{"GAMMALN", 1u, 1u, &Gammaln});
  registry.register_function(FunctionDef{"GAMMALN.PRECISE", 1u, 1u, &Gammaln});
  registry.register_function(FunctionDef{"GAMMA.DIST", 4u, 4u, &GammaDist});
  registry.register_function(FunctionDef{"GAMMA.INV", 3u, 3u, &GammaInv});

  registry.register_function(FunctionDef{"WEIBULL.DIST", 4u, 4u, &WeibullDist});

  registry.register_function(FunctionDef{"LOGNORM.DIST", 4u, 4u, &LognormDist});
  registry.register_function(FunctionDef{"LOGNORM.INV", 3u, 3u, &LognormInv});

  registry.register_function(FunctionDef{"HYPGEOM.DIST", 5u, 5u, &HypgeomDist});

  // Legacy (pre-2010) spellings. BETAINV, GAMMADIST, GAMMAINV, WEIBULL,
  // and LOGINV are signature-identical to the .NEW form and share the
  // impl pointer. LOGNORMDIST / HYPGEOMDIST / BETADIST route through
  // wrappers above because their legacy signatures drop or shift the
  // cumulative flag.
  registry.register_function(FunctionDef{"BETADIST", 3u, 5u, &BetaDistLegacy});
  registry.register_function(FunctionDef{"BETAINV", 3u, 5u, &BetaInv});
  registry.register_function(FunctionDef{"GAMMADIST", 4u, 4u, &GammaDist});
  registry.register_function(FunctionDef{"GAMMAINV", 3u, 3u, &GammaInv});
  registry.register_function(FunctionDef{"WEIBULL", 4u, 4u, &WeibullDist});
  registry.register_function(FunctionDef{"LOGNORMDIST", 3u, 3u, &LognormDistLegacy});
  registry.register_function(FunctionDef{"LOGINV", 3u, 3u, &LognormInv});
  registry.register_function(FunctionDef{"HYPGEOMDIST", 4u, 4u, &HypgeomDistLegacy});
}

}  // namespace eval
}  // namespace formulon
