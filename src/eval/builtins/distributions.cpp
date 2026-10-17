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
// lean on `stats::gamma_ratio_excel` and `stats::regularized_incomplete_beta` from
// `eval/stats/incomplete_gamma.h` and `eval/stats/special_functions.h`; the inverses use a shared
// bracket-then-Newton helper defined locally in this TU.
//
// Numerical notes:
//   * Gamma / log-gamma: std::tgamma / stats::log_gamma handle the generic
//     case. tgamma is NaN/Inf at non-positive integers (poles); we
//     surface #NUM! explicitly before calling std::tgamma so the
//     behaviour matches Excel regardless of the platform's tgamma
//     NaN vs. Inf convention.
//   * GAMMA.INV follows the chi-squared inverse identity
//       GAMMA.INV(p, alpha, beta) == beta * chisq_inv(p, 2*alpha) / 2,
//     but runs its own Wilson-Hilferty seed + Newton driver: the
//     chi-squared solver (`ChisqInvCore` in
//     `eval/builtins/stats/stats_distributions.cpp`) is TU-local. The
//     inverse standard-normal seed is NOT replicated — it comes from
//     `stats_detail::InverseStandardNormal`, the single definition all
//     quantile callers share.
//   * BETA.INV uses a shared bisection-then-Newton driver on the CDF
//     surface, bracketing on [0, 1].
//   * HYPGEOM.DIST's PMF uses lgamma-based log-combinations so the
//     finite factorials used by large (K, N) stay stable.

#include "eval/builtins/distributions.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>

#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/builtins/stats/stats_helpers.h"
#include "eval/coerce.h"
#include "eval/excel_math.h"
#include "eval/function_registry.h"
#include "eval/stats/incomplete_beta.h"
#include "eval/stats/incomplete_gamma.h"
#include "eval/stats/inverse_distributions.h"
#include "eval/stats/special_functions.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::NumberTriple;

// Shared kPi alias kept for readability at the lognormal PDF normalisation
// site; the value lives in `numeric_helpers.h`.
constexpr double kDistPi = builtins_detail::kPi;

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

// Finalises a scalar result: non-finite becomes #NUM!, otherwise wraps in a
// Value::number. Thin alias kept so the impl bodies read naturally.
inline Value finalize(double r) {
  return builtins_detail::to_finite_value(r);
}

// Reads a required numeric argument. Non-finite coerced values surface
// as #NUM! so the impl bodies never have to re-check for NaN / Inf.
inline Expected<double, ErrorCode> read_number(const Value* args, std::uint32_t index) {
  return builtins_detail::read_required_number(args, index);
}

// Reads an optional trailing numeric argument, falling back to
// `default_value` when arity <= index. Parallels `read_number` for
// error/non-finite handling.
inline Expected<double, ErrorCode> read_optional_number(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                        double default_value) {
  return builtins_detail::read_optional_number(args, arity, index, default_value);
}

inline Expected<NumberTriple, ErrorCode> read_number_triple(const Value* args) {
  return builtins_detail::read_number_triple(args, 0, 1, 2);
}

using stats_detail::read_triple_cumulative;
using stats_detail::TripleCumulative;

// Reads the (x, alpha, beta, cumulative) arguments of GAMMA.DIST / WEIBULL.DIST
// and applies their shared domain check: x >= 0, alpha > 0, beta > 0.
Expected<TripleCumulative, ErrorCode> read_scaled_cumulative_triple(const Value* args) {
  auto parsed = read_triple_cumulative(args);
  if (!parsed) {
    return std::move(parsed.error());
  }
  const TripleCumulative& p = parsed.value();
  if (p.first < 0.0 || p.second <= 0.0 || p.third <= 0.0) {
    return ErrorCode::Num;
  }
  return parsed.value();
}

// Shared bracket-then-Newton inverter for CDF surfaces on a half-open
// interval. The caller supplies `cdf(x)` and its derivative `pdf(x)`
// along with a bracket `[lo, hi]` known to contain the root (i.e.
// cdf(lo) <= p <= cdf(hi)). The driver runs a bisection warm-up to
// tighten the bracket, then switches to Newton with a bracket-escape
// safeguard that falls back to bisection if a Newton step would land
// outside `[lo, hi]`.
//
// Convergence policy: termination is decided in *x-space* using a
// relative bound `kTol * max(kTol, |x_new|)`, not on the CDF residual
// `|cdf(x) - p|`. For tail-tiny probabilities (e.g. p ~ 1e-24 with a
// true root x ~ 1e-12) the residual can fall well below `kTol` while
// `x` is still many orders of magnitude away from the root, so a
// residual-based stop returns gibberish. The `max(kTol, |x_new|)`
// floor keeps the relative test meaningful when the answer sits at
// the support boundary (|x| -> 0) without demanding sub-1e-24
// precision, which would loop forever against double-precision noise.
//
// Returns NaN on non-convergence so wrappers can surface #NUM!.
template <typename CdfFn, typename PdfFn>
double BracketThenNewton(CdfFn cdf, PdfFn pdf, double lo, double hi, double p) {
  constexpr int kBisectIter = 40;
  constexpr int kNewtonIter = 100;
  constexpr double kTol = 1e-12;

  // Bisection warm-up: tightens the bracket before Newton starts, which
  // keeps tail-steep CDFs from oscillating.
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
    // Stop when the bracket is tight enough either in relative or
    // absolute terms. The relative bound (`1e-6 * |mid|`) lets Newton
    // take over fast for mid-range answers; the absolute bound
    // (`kTol`) ensures we don't stop bisecting prematurely when the
    // answer is at the support boundary (mid -> 0).
    if ((hi - lo) < std::max(kTol, 1e-6 * std::abs(mid))) {
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
    // Relative-x convergence: 12 sig digits, with an absolute floor of
    // kTol to handle answers genuinely at the support boundary (where
    // |x| -> 0). The original `kTol * std::max(1.0, |x|)` floor of 1.0
    // turned this into an absolute kTol check for `|x| < 1`, which (for
    // tail-tiny answers like x ~ 1e-12) made the relative test
    // vacuously easy and let Newton terminate at the wrong scale.
    if (std::abs(x_new - x) < kTol * std::max(kTol, std::abs(x_new))) {
      return x_new;
    }
    x = x_new;
  }
  return std::numeric_limits<double>::quiet_NaN();
}

// ---------------------------------------------------------------------------
// BETA family
// ---------------------------------------------------------------------------

// BETA.DIST(x, alpha, beta, cumulative, [A], [B]) - beta distribution on
// the optional support [A, B]. The standard [0, 1] form is recovered by
// omitting A / B (defaults A = 0, B = 1). For x == A or x == B the PDF
// hits boundary cases; for x strictly inside (A, B) we rescale via
//   y = (x - A) / (B - A)
// and delegate to the standard beta PDF/CDF, rescaling the PDF by the
// Jacobian 1 / (B - A). #NUM! on: alpha <= 0, beta <= 0, A >= B,
// x < A, x > B.
Value BetaDist(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto parsed = read_triple_cumulative(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  auto a_e = read_optional_number(args, arity, 4, 0.0);
  if (!a_e) {
    return Value::error(a_e.error());
  }
  auto b_e = read_optional_number(args, arity, 5, 1.0);
  if (!b_e) {
    return Value::error(b_e.error());
  }
  const double x = parsed.value().first;
  const double alpha = parsed.value().second;
  const double beta_shape = parsed.value().third;
  const double a = a_e.value();
  const double b = b_e.value();
  if (alpha <= 0.0 || beta_shape <= 0.0 || a >= b || x < a || x > b) {
    return Value::error(ErrorCode::Num);
  }
  const double span = b - a;
  const double y = (x - a) / span;
  if (parsed.value().cumulative) {
    // CDF boundaries: the regularized incomplete beta already returns 0
    // at y == 0 and 1 at y == 1 exactly, so no special-case is needed.
    return finalize(stats::beta_dist_cdf_excel(x, alpha, beta_shape, a, b));
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
  return finalize(stats::beta_dist_pdf_excel(x, alpha, beta_shape, a, b));
}

// BETA.INV(p, alpha, beta, [A], [B]) - inverse of BETA.DIST's CDF. Solves
//   I_y(alpha, beta) == p
// for y in (0, 1), then rescales via x = A + y * (B - A). #NUM! on
// p <= 0, p >= 1, alpha <= 0, beta <= 0, A >= B.
Value BetaInv(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto parsed = read_number_triple(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  auto a_e = read_optional_number(args, arity, 3, 0.0);
  if (!a_e) {
    return Value::error(a_e.error());
  }
  auto b_e = read_optional_number(args, arity, 4, 1.0);
  if (!b_e) {
    return Value::error(b_e.error());
  }
  const double p = parsed.value().first;
  const double alpha = parsed.value().second;
  const double beta_shape = parsed.value().third;
  const double a = a_e.value();
  const double b = b_e.value();
  if (p <= 0.0 || p >= 1.0 || alpha <= 0.0 || beta_shape <= 0.0 || a >= b) {
    return Value::error(ErrorCode::Num);
  }
  const double r = stats::beta_inv_excel(p, alpha, beta_shape, a, b);
  if (!std::isfinite(r)) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(r);
}

// ---------------------------------------------------------------------------
// GAMMA family
// ---------------------------------------------------------------------------

// Coefficients of the GAMMA kernel: the Stirling series in 1/x for |x| > 10 and the
// rational approximation (numerator over denominator) on [1, 2] for |x| <= 10.
constexpr double kGammaStirling[] = {
    -0.05905576189139824,   0.001080329535785209,   0.01280672566761614,   -0.00032503252556783166,
    -0.003828876997130955,  0.00014409790832040022, 0.0016789974413441745, -0.00010343581816521184,
    -0.0011843328747073879, 0.00013945627516731714, 0.0015680784434401332, -0.00045894418724279834,
    -0.0053626543209876545, 0.006944444444444444,   0.16666666666666666,   2.0};
constexpr double kGammaNumerator[] = {-41.32975575688121, -224.55734984993717, -1766.0581782060663, -11262.704073349521,
                                      -39779.10693069725, -183439.1689740763,  -345373.8010894497,  -896026.3354389172};
constexpr double kGammaDenominator[] = {
    -319.17974684033175, 6642.776155506226,  -49294.268969809906, 118905.79578918342,
    204704.893587802,    -862516.6292015722, -896037.449404572,   1.0};
constexpr double kEulerGamma = 0.5772156649015329;
constexpr double kSqrtPi = 1.772453850905516;
constexpr double kSqrtHalfPi = 1.2533141373155003;

template <std::size_t N>
double horner_series(const double (&coefficients)[N], double t) noexcept {
  double acc = coefficients[0];
  for (std::size_t i = 1; i < N; ++i) {
    acc = acc * t + coefficients[i];
  }
  return acc;
}

// base^e for a non-negative integer e: square-and-multiply with the product
// taken as `base * result` on each set bit.
double integer_power(double base, std::uint64_t e) noexcept {
  double result = 1.0;
  while (true) {
    if ((e & 1U) != 0U) {
      result = base * result;
    }
    if (e < 2U) {
      return result;
    }
    e >>= 1U;
    base = base * base;
  }
}

// Gamma function. NaN marks the #NUM! domain (poles and overflow).
double gamma_kernel(double x) noexcept {
  const double nan = std::numeric_limits<double>::quiet_NaN();
  if (x > 171.6243769563027) {
    return nan;
  }
  const bool integral = std::trunc(x) - x == 0.0;
  if (x < -171.03093456965283) {
    return integral ? nan : 0.0;
  }
  if (integral) {
    if (x <= 0.0) {
      return nan;
    }
    if (!(x > 88.0)) {
      double factorial = 1.0;
      for (auto k = static_cast<std::int64_t>(x) - 1; k > 0; --k) {
        factorial = static_cast<double>(k) * factorial;
      }
      return factorial;
    }
  }
  const double ax = std::fabs(x);
  const auto whole = static_cast<std::uint64_t>(std::trunc(ax));
  const double frac = ax - static_cast<double>(whole);
  if (frac == 0.5) {
    // Half-integers: the product of i/2 over odd i below 2n, times sqrt(pi).
    double result = 1.0;
    double count = 1.0;
    for (std::uint64_t i = 1; i < 2 * whole; ++i) {
      if ((i & 1U) != 0U) {
        result = (count / 2.0) * result;
      }
      count = count + 1.0;
    }
    result = result * kSqrtPi;
    if (x < 0.0) {
      result = (-kDistPi) / result;
      result = result / ax;
      if ((whole & 1U) != 0U) {
        result = -result;
      }
    }
    return result;
  }
  if (ax <= 2.220446049250313e-16) {
    return 1.0 / x - kEulerGamma;
  }
  if (ax <= 10.0) {
    double y = x;
    double product = 1.0;
    while (y > 2.0) {
      y = y - 1.0;
      product = product * y;
    }
    while (y < 1.0) {
      product = product / y;
      y = y + 1.0;
    }
    return (horner_series(kGammaNumerator, y) / horner_series(kGammaDenominator, y)) * product;
  }
  // Stirling: Gamma(|x|) = sqrt(pi/2) * exp(-|x|) * sqrt(1/|x|) * |x|^|x| * S(1/|x|),
  // with |x|^|x| split as (|x|^(m + r))^2 for |x| / 2 = m + r.
  const double t = 1.0 / ax;
  const double series = horner_series(kGammaStirling, t);
  const double root = std::sqrt(t);
  const double e = std::exp(-ax) * root;
  const double half = ax * 0.5;
  const double m = std::trunc(half);
  const double r = half - m;
  double power = integer_power(ax, static_cast<std::uint64_t>(m));
  power = std::exp(std::log(ax) * r) * power;
  double v = e * power;
  v = v * power;
  v = v * kSqrtHalfPi;
  double result = v * series;
  if (x < 0.0) {
    result = (-kDistPi) / result;
    const double f = frac > 0.5 ? 1.0 - frac : frac;
    double sine = std::sin(kDistPi * f);
    sine = sine * ax;
    result = result / sine;
    if ((whole & 1U) != 0U) {
      result = -result;
    }
  }
  return result;
}

// GAMMA(x) - the gamma function; poles (non-positive integers) and overflow are #NUM!.
Value Gamma(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x_e = read_number(args, 0);
  if (!x_e) {
    return Value::error(x_e.error());
  }
  return finalize(gamma_kernel(x_e.value()));
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
  return finalize(stats::lgamma_excel(x));
}

// Gamma PDF at x > 0 with shape `alpha` and scale `beta`:
// `dpois_raw(x / beta, alpha) * (alpha / x)`. Callers reject x < 0,
// alpha <= 0, beta <= 0 up front.
double GammaPdf(double x, double alpha, double beta_scale) noexcept {
  return flush_subnormal(stats::dpois_raw(flush_subnormal(x / beta_scale), alpha) * flush_subnormal(alpha / x));
}

// GAMMA.DIST(x, alpha, beta, cumulative). #NUM! on x < 0, alpha <= 0,
// beta <= 0. At x == 0 the PDF is:
//   alpha <= 1: #NUM! (Excel rejects the exponential rate at alpha == 1 too)
//   alpha > 1: 0
// The CDF at x == 0 is 0 by definition.
Value GammaDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto parsed = read_scaled_cumulative_triple(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const double x = parsed.value().first;
  const double alpha = parsed.value().second;
  const double beta_scale = parsed.value().third;
  if (parsed.value().cumulative) {
    return finalize(stats::gamma_ratio_excel(alpha, flush_subnormal(x / beta_scale), false));
  }
  if (x == 0.0) {
    if (alpha <= 1.0) {
      return Value::error(ErrorCode::Num);
    }
    return finalize(0.0);
  }
  return finalize(GammaPdf(x, alpha, beta_scale));
}

// GAMMA.INV(p, alpha, beta) - inverse of GAMMA.DIST's CDF, `beta` times the
// unit-scale gamma quantile (`stats::gamma_inv_excel`).
Value GammaInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto parsed = read_number_triple(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [p, alpha, beta_scale] = parsed.value();
  if (p < 0.0 || p >= 1.0 || alpha <= 0.0 || beta_scale <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (p == 0.0) {
    return finalize(0.0);
  }
  const double x = stats::gamma_inv_excel(p, alpha, beta_scale);
  if (std::isnan(x)) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(x);
}

// ---------------------------------------------------------------------------
// WEIBULL.DIST
// ---------------------------------------------------------------------------

// WEIBULL.DIST(x, alpha, beta, cumulative). The Weibull distribution with
// shape `alpha` and scale `beta`, evaluated with a plain exp(log(b) * e) power:
//   CDF: 0 - expm1(-(x/beta)^alpha)
//   PDF: ((alpha / beta^alpha) * x^(alpha-1)) * exp(-(x/beta)^alpha)
// #NUM! on x < 0, alpha <= 0, beta <= 0; the density at x == 0 is 0.
Value WeibullDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto parsed = read_scaled_cumulative_triple(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const double x = parsed.value().first;
  const double alpha = parsed.value().second;
  const double beta_scale = parsed.value().third;
  const double scaled_power = std::exp(std::log(x / beta_scale) * alpha);
  if (parsed.value().cumulative) {
    return finalize(0.0 - excel_expm1(-scaled_power));
  }
  // At x == 0 the x^(alpha-1) factor is 0 for every alpha, so the density there is 0.
  const double x_power = x != 0.0 ? std::exp(std::log(x) * (alpha - 1.0)) : 0.0;
  const double r = ((alpha / std::exp(std::log(beta_scale) * alpha)) * x_power) * std::exp(-scaled_power);
  return finalize(r);
}

// ---------------------------------------------------------------------------
// LOGNORM family
// ---------------------------------------------------------------------------

// Error-free sum: `sum` is a + b rounded, `error` the exact remainder.
struct TwoSum {
  double sum;
  double error;
};

inline TwoSum two_sum(double a, double b) noexcept {
  const double sum = a + b;
  const double bb = sum - a;
  return TwoSum{sum, (a - (sum - bb)) + (b - bb)};
}

// Log-normal density exp(-(ln sigma + z + ln sqrt(2 pi) + ln x)) with
// z = (ln x - mu)^2 / (2 sigma^2). The exponent is summed with error terms
// kept, and the rounding errors re-enter as a second exp factor.
double lognormal_pdf(double x, double mean, double sd) noexcept {
  constexpr double kLogSqrt2Pi = 0.9189385332046728;
  const double d = std::log(x) - mean;
  const double z = (d * d) / (2.0 * (sd * sd));
  const TwoSum first = two_sum(std::log(sd), z);
  const TwoSum second = two_sum(first.sum, kLogSqrt2Pi);
  const double log_x = std::log(x);
  // The final sum swaps the roles of its operands in the error term.
  const double total = log_x + second.sum;
  const double bb = total - second.sum;
  const double third_error = (second.sum - (total - bb)) + (log_x - bb);
  const double error = (first.error + second.error) + third_error;
  return std::exp(-total) * std::exp(error);
}

// LOGNORM.DIST(x, mean, sd, cumulative). The log-normal distribution
// parameterised by (mean, sd) on the log scale:
//   CDF: Phi((ln x - mean) / sd)
//   PDF: see `lognormal_pdf`
// #NUM! on x <= 0, sd <= 0.
Value LognormDist(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto parsed = read_triple_cumulative(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const double x = parsed.value().first;
  const double mean = parsed.value().second;
  const double sd = parsed.value().third;
  if (x <= 0.0 || sd <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double z = (std::log(x) - mean) / sd;
  if (parsed.value().cumulative) {
    return finalize(stats::normal_cdf(z));
  }
  return finalize(lognormal_pdf(x, mean, sd));
}

// LOGNORM.INV(p, mean, sd) - inverse log-normal CDF. Computed as
//   exp(mean + sd * stats_detail::InverseStandardNormal(p))
// #NUM! on p <= 0, p >= 1, sd <= 0.
Value LognormInv(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto parsed = read_number_triple(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [p, mean, sd] = parsed.value();
  if (p <= 0.0 || p >= 1.0 || sd <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double z = stats_detail::InverseStandardNormal(p);
  return finalize(std::exp(mean + sd * z));
}

// ---------------------------------------------------------------------------
// HYPGEOM.DIST
// ---------------------------------------------------------------------------

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
  double counts[4];
  for (std::uint32_t i = 0; i < 4U; ++i) {
    auto count_e = read_number(args, i);
    if (!count_e) {
      return Value::error(count_e.error());
    }
    counts[i] = count_e.value();
  }
  auto cum_e = coerce_to_bool(args[4]);
  if (!cum_e) {
    return Value::error(cum_e.error());
  }
  const double k = std::floor(counts[0]);
  const double n = std::floor(counts[1]);
  const double big_k = std::floor(counts[2]);
  const double big_n = std::floor(counts[3]);
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
    // The cdf is a term-by-term recurrence: a span past the term cap is rejected rather than walked.
    if (k - k_min >= stats_detail::kMaxCumulativeTerms) {
      return Value::error(ErrorCode::Num);
    }
    return finalize(stats::hypgeom_cdf_excel(k, n, big_k, big_n));
  }
  // PMF branch: 0 outside the support, otherwise the log-combination form.
  if (k < k_min || k > k_max) {
    return finalize(0.0);
  }
  return finalize(stats::hypgeom_pmf_excel(k, n, big_k, big_n));
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
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"BETA.DIST", 4u, 6u, &BetaDist},
      {"BETA.INV", 3u, 5u, &BetaInv},
      {"GAMMA", 1u, 1u, &Gamma},
      {"GAMMALN", 1u, 1u, &Gammaln},
      {"GAMMALN.PRECISE", 1u, 1u, &Gammaln},
      {"GAMMA.DIST", 4u, 4u, &GammaDist},
      {"GAMMA.INV", 3u, 3u, &GammaInv},
      {"WEIBULL.DIST", 4u, 4u, &WeibullDist},
      {"LOGNORM.DIST", 4u, 4u, &LognormDist},
      {"LOGNORM.INV", 3u, 3u, &LognormInv},
      {"HYPGEOM.DIST", 5u, 5u, &HypgeomDist},
      {"BETADIST", 3u, 5u, &BetaDistLegacy},
      {"BETAINV", 3u, 5u, &BetaInv},
      {"GAMMADIST", 4u, 4u, &GammaDist},
      {"GAMMAINV", 3u, 3u, &GammaInv},
      {"WEIBULL", 4u, 4u, &WeibullDist},
      {"LOGNORMDIST", 3u, 3u, &LognormDistLegacy},
      {"LOGINV", 3u, 3u, &LognormInv},
      {"HYPGEOMDIST", 4u, 4u, &HypgeomDistLegacy},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
