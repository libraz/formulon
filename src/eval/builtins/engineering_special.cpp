//
// Implementation of Excel's ERF / BESSEL engineering built-ins.
//
// Error-function family (4): ERF (1 or 2 arg), ERF.PRECISE (1 arg),
// ERFC (1 arg), ERFC.PRECISE (1 arg). All delegate to the C library
// `std::erf` / `std::erfc` on finite real doubles; a NaN result (from NaN
// input) is reported as #NUM!.
//
// Bessel family (4): BESSELJ, BESSELY, BESSELI, BESSELK, all `(x, n)`
// where `n` is a non-negative integer order (truncated toward zero).
//
// Mac Excel 365 produces Bessel values that match the Numerical Recipes
// in C (2nd ed., 1992) Chapter 6.5 / 6.6 routines bit-for-bit (or within
// 1 ULP) on the entire oracle corpus. Probes against Mac confirm divergence
// from libm `::jn` / `::yn` and from the canonical I_n power series; the
// only formulation that reproduces Mac is the published 6.5/6.6
// Chebyshev / rational-polynomial / Miller-recurrence package, whose
// coefficients are widely tabulated (Abramowitz & Stegun, NIST DLMF,
// netlib TOMS) and are reproduced here verbatim.
//
//   * BESSELJ uses POSIX `::jn` (libm). Mac probes show no observable
//     drift from libm for the J_n cases in our corpus, so the shorter,
//     library-backed path is retained.
//   * BESSELY uses Y_0(x) / Y_1(x) computed from rational polynomials
//     (small x, with the J_n log correction) and an asymptotic
//     expansion (large x), then forward recurrence
//     Y_{n+1} = (2n/x) Y_n - Y_{n-1} for n >= 2. Forward recurrence is
//     numerically stable upward for the second-kind Y_n.
//   * BESSELI uses I_0(x) / I_1(x) computed from Chebyshev-style
//     polynomials (small |x|) and an asymptotic expansion (large |x|),
//     then Miller's downward recurrence for n >= 2 with renormalisation
//     against I_0 to control overflow.
//   * BESSELK uses K_0(x) / K_1(x) computed from polynomial
//     approximations (with the I_n log correction) for x <= 2 and an
//     asymptotic expansion for x > 2, then forward recurrence
//     K_{n+1} = (2n/x) K_n + K_{n-1} for n >= 2.
//
// Domain rules for BESSEL* (matching Excel's API):
//   * n < 0                                   -> #NUM!
//   * BESSELY / BESSELK with x == 0           -> #NUM! (singular at 0)
//   * BESSELY / BESSELK with x < 0            -> #NUM! (defined only on
//                                                positive reals by Excel)
//   * BESSELJ / BESSELI accept any real x.

#include "eval/builtins/engineering_special.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Shared coercion helpers
// ---------------------------------------------------------------------------

// Coerces an ERF / BESSEL numeric argument. Bool is rejected with #VALUE!
// (Excel 365 rejects TRUE/FALSE for the ERF and BESSEL families), Text is
// parsed via coerce_to_number, Blank -> 0. NaN / Inf are left to the caller
// to interpret (most call sites treat NaN as #NUM! at the result stage).
Expected<double, ErrorCode> coerce_real_arg(const Value& v) {
  if (v.kind() == ValueKind::Bool) {
    return ErrorCode::Value;
  }
  return coerce_to_number(v);
}

// Coerces a BESSEL order argument. Must be a finite number >= 0 after
// truncation toward zero; returns the integer order. Negative or non-finite
// -> #NUM!. Bool is rejected with #VALUE! to match Excel 365.
Expected<int, ErrorCode> coerce_bessel_order(const Value& v) {
  if (v.kind() == ValueKind::Bool) {
    return ErrorCode::Value;
  }
  auto n = coerce_to_number(v);
  if (!n) {
    return n.error();
  }
  const double d = n.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  const double t = std::trunc(d);
  if (t < 0.0) {
    return ErrorCode::Num;
  }
  // Upper bound: guard against absurd orders that would overflow the
  // recurrence. Excel caps orders well below this; 2^30 is a defensive
  // ceiling that still accommodates any realistic engineering query.
  if (t > static_cast<double>((1 << 30))) {
    return ErrorCode::Num;
  }
  return static_cast<int>(t);
}

// ---------------------------------------------------------------------------
// ERF family
// ---------------------------------------------------------------------------

// ERF accepts either 1 or 2 arguments:
//   * 1-arg: erf(x).
//   * 2-arg: erf(upper) - erf(lower).
// NaN result -> #NUM!.
Value Erf(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto a = coerce_real_arg(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  if (arity == 1) {
    const double r = std::erf(a.value());
    if (std::isnan(r)) {
      return Value::error(ErrorCode::Num);
    }
    return Value::number(r);
  }
  // 2-arg: treat args[0] as lower, args[1] as upper. Excel's documented
  // shape is `ERF(lower, [upper])` with the result being erf(upper) -
  // erf(lower).
  auto b = coerce_real_arg(args[1]);
  if (!b) {
    return Value::error(b.error());
  }
  const double lower = a.value();
  const double upper = b.value();
  const double r = std::erf(upper) - std::erf(lower);
  if (std::isnan(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ERF.PRECISE: strict 1-arg erf(x). The registry enforces arity (min=max=1)
// so an attempt to pass a second argument surfaces #VALUE! from the
// dispatcher before this impl runs.
Value ErfPrecise(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto a = coerce_real_arg(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  const double r = std::erf(a.value());
  if (std::isnan(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ERFC(x) = 1 - erf(x). Uses std::erfc for better precision in the tail.
Value Erfc(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto a = coerce_real_arg(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  const double r = std::erfc(a.value());
  if (std::isnan(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ERFC.PRECISE: numerically identical to ERFC. Excel keeps the separate
// name for parallelism with ERF / ERF.PRECISE.
Value ErfcPrecise(const Value* args, std::uint32_t arity, Arena& arena) {
  return Erfc(args, arity, arena);
}

// ---------------------------------------------------------------------------
// BESSEL helpers (Numerical Recipes 6.5 / 6.6)
//
// All polynomial / rational coefficients below are the published
// tables that drive Mac Excel's bit-exact output. They appear in NR2 6.5
// (BESSJ0/BESSJ1/BESSY0/BESSY1) and 6.6 (BESSI0/BESSI1/BESSK0/BESSK1) and
// are also tabulated in Abramowitz & Stegun 9.4 / 9.8 and on netlib.
// ---------------------------------------------------------------------------

// J_0(x): rational polynomial for |x| < 8, asymptotic for |x| >= 8.
// Used by Y_0 (log-term correction). Not exposed for BESSELJ itself --
// libm `::jn` is sufficient for J_n in our oracle corpus.
double bessj0(double x) {
  const double ax = std::fabs(x);
  if (ax < 8.0) {
    const double y = x * x;
    const double r =
        57568490574.0 +
        y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * (-184.9052456)))));
    const double s =
        57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y * 1.0))));
    return r / s;
  }
  const double z = 8.0 / ax;
  const double y = z * z;
  const double p = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
  const double q =
      -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * (-0.934935152e-7))));
  const double xx = ax - 0.785398164;
  return std::sqrt(0.636619772 / ax) * (std::cos(xx) * p - z * std::sin(xx) * q);
}

// J_1(x): rational polynomial for |x| < 8, asymptotic for |x| >= 8.
// Used by Y_1 (log-term correction).
double bessj1(double x) {
  const double ax = std::fabs(x);
  double ans;
  if (ax < 8.0) {
    const double y = x * x;
    const double r =
        x * (72362614232.0 +
             y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.48260 + y * (-30.16036606))))));
    const double s =
        144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y * 1.0))));
    ans = r / s;
  } else {
    const double z = 8.0 / ax;
    const double y = z * z;
    const double p = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * (-0.240337019e-6))));
    const double q =
        0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
    const double xx = ax - 2.356194491;
    ans = std::sqrt(0.636619772 / ax) * (std::cos(xx) * p - z * std::sin(xx) * q);
    if (x < 0.0) {
      ans = -ans;
    }
  }
  return ans;
}

// Y_0(x), x > 0. Rational polynomial + (2/pi) J_0(x) ln(x) for x < 8;
// asymptotic expansion for x >= 8.
double bessy0(double x) {
  if (x < 8.0) {
    const double y = x * x;
    const double r = -2957821389.0 +
                     y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))));
    const double s =
        40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.26470 + y * (226.1030244 + y * 1.0))));
    return r / s + 0.636619772 * bessj0(x) * std::log(x);
  }
  const double z = 8.0 / x;
  const double y = z * z;
  const double p = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
  const double q =
      -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * (-0.934945152e-7))));
  const double xx = x - 0.785398164;
  return std::sqrt(0.636619772 / x) * (std::sin(xx) * p + z * std::cos(xx) * q);
}

// Y_1(x), x > 0. Rational polynomial + (2/pi)(J_1(x) ln(x) - 1/x) for
// x < 8; asymptotic expansion for x >= 8.
double bessy1(double x) {
  if (x < 8.0) {
    const double y = x * x;
    const double r =
        x * (-0.4900604943e13 +
             y * (0.1275274390e13 +
                  y * (-0.5153438139e11 + y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))));
    const double s =
        0.2499580570e14 +
        y * (0.4244419664e12 +
             y * (0.3733650367e10 + y * (0.2245904002e8 + y * (0.1020426050e6 + y * (0.3549632885e3 + y * 1.0)))));
    return r / s + 0.636619772 * (bessj1(x) * std::log(x) - 1.0 / x);
  }
  const double z = 8.0 / x;
  const double y = z * z;
  const double p = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * (-0.240337019e-6))));
  const double q =
      0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
  const double xx = x - 2.356194491;
  return std::sqrt(0.636619772 / x) * (std::sin(xx) * p + z * std::cos(xx) * q);
}

// Y_n(x), x > 0, via forward recurrence Y_{n+1} = (2n/x) Y_n - Y_{n-1}.
// Forward recurrence is numerically stable upward for the second-kind
// solution (Y_n grows in magnitude with n at fixed small x).
double bessy(int n, double x) {
  if (n == 0) {
    return bessy0(x);
  }
  if (n == 1) {
    return bessy1(x);
  }
  const double tox = 2.0 / x;
  double bym = bessy0(x);
  double by = bessy1(x);
  for (int j = 1; j < n; ++j) {
    const double byp = static_cast<double>(j) * tox * by - bym;
    bym = by;
    by = byp;
  }
  return by;
}

// I_0(|x|): Chebyshev polynomial for |x| < 3.75, asymptotic for |x| >= 3.75.
// I_0 is even, so we operate on |x|.
double bessi0(double x) {
  const double ax = std::fabs(x);
  if (ax < 3.75) {
    double y = x / 3.75;
    y = y * y;
    return 1.0 +
           y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.0360768 + y * 0.0045813)))));
  }
  const double y = 3.75 / ax;
  return (std::exp(ax) / std::sqrt(ax)) *
         (0.39894228 +
          y * (0.01328592 +
               y * (0.00225319 +
                    y * (-0.00157565 +
                         y * (0.00916281 +
                              y * (-0.02057706 + y * (0.02635537 + y * (-0.01647633 + y * 0.00392377))))))));
}

// I_1(x): Chebyshev polynomial for |x| < 3.75, asymptotic for |x| >= 3.75.
// I_1 is odd; the small-x branch returns ax * P((x/3.75)^2) and the sign
// is restored from the original x.
double bessi1(double x) {
  const double ax = std::fabs(x);
  double ans;
  if (ax < 3.75) {
    double y = x / 3.75;
    y = y * y;
    ans = ax * (0.5 + y * (0.87890594 +
                           y * (0.51498869 + y * (0.15084934 + y * (0.02658733 + y * (0.00301532 + y * 0.00032411))))));
  } else {
    const double y = 3.75 / ax;
    ans = 0.39894228 +
          y * (-0.03988024 +
               y * (-0.00362018 +
                    y * (0.00163801 + y * (-0.01031555 +
                                           y * (0.02282967 + y * (-0.02895312 + y * (0.01787654 - y * 0.00420059)))))));
    ans *= std::exp(ax) / std::sqrt(ax);
  }
  return x < 0.0 ? -ans : ans;
}

// I_n(x), n >= 2, via Miller's downward recurrence starting at
// m = 2 * (n + sqrt(40 n)) and normalising against I_0(x). Renormalises
// the running terms when the magnitude exceeds 1e10 to keep the
// recurrence in IEEE-754 range. I_n is even when n is even and odd when
// n is odd, so the sign is restored from the original x and parity of n.
double bessi(int n, double x) {
  if (n == 0) {
    return bessi0(x);
  }
  if (n == 1) {
    return bessi1(x);
  }
  if (x == 0.0) {
    return 0.0;
  }
  const double ax = std::fabs(x);
  const double tox = 2.0 / ax;
  double bip = 0.0;
  double bi = 1.0;
  double ans = 0.0;
  // Starting index well above n, per NR's stability heuristic.
  const int m = 2 * (n + static_cast<int>(std::sqrt(40.0 * static_cast<double>(n))));
  for (int j = m; j > 0; --j) {
    const double bim = bip + static_cast<double>(j) * tox * bi;
    bip = bi;
    bi = bim;
    if (std::fabs(bi) > 1e10) {
      ans *= 1e-10;
      bip *= 1e-10;
      bi *= 1e-10;
    }
    if (j == n) {
      ans = bip;
    }
  }
  ans *= bessi0(x) / bi;
  return (x < 0.0 && (n & 1) != 0) ? -ans : ans;
}

// K_0(x), x > 0. Polynomial approximation + I_0(x) ln(x/2) correction
// for x <= 2; asymptotic expansion exp(-x)/sqrt(x) * P(2/x) for x > 2.
double bessk0(double x) {
  if (x <= 2.0) {
    const double y = x * x / 4.0;
    return -std::log(x / 2.0) * bessi0(x) +
           (-0.57721566 +
            y * (0.42278420 +
                 y * (0.23069756 + y * (0.3488590e-1 + y * (0.262698e-2 + y * (0.10750e-3 + y * 0.74e-5))))));
  }
  const double y = 2.0 / x;
  return (std::exp(-x) / std::sqrt(x)) *
         (1.25331414 +
          y * (-0.7832358e-1 +
               y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 + y * (-0.251540e-2 + y * 0.53208e-3))))));
}

// K_1(x), x > 0. Polynomial approximation + I_1(x) ln(x/2) correction
// for x <= 2; asymptotic expansion for x > 2.
double bessk1(double x) {
  if (x <= 2.0) {
    const double y = x * x / 4.0;
    return std::log(x / 2.0) * bessi1(x) +
           (1.0 / x) *
               (1.0 + y * (0.15443144 +
                           y * (-0.67278579 +
                                y * (-0.18156897 + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * (-0.4686e-4)))))));
  }
  const double y = 2.0 / x;
  return (std::exp(-x) / std::sqrt(x)) *
         (1.25331414 +
          y * (0.23498619 +
               y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 + y * (0.325614e-2 + y * (-0.68245e-3)))))));
}

// K_n(x), x > 0, via forward recurrence K_{n+1} = (2n/x) K_n + K_{n-1}.
// Forward recurrence is numerically stable upward for K_n.
double bessk(int n, double x) {
  if (n == 0) {
    return bessk0(x);
  }
  if (n == 1) {
    return bessk1(x);
  }
  const double tox = 2.0 / x;
  double bkm = bessk0(x);
  double bk = bessk1(x);
  for (int j = 1; j < n; ++j) {
    const double bkp = bkm + static_cast<double>(j) * tox * bk;
    bkm = bk;
    bk = bkp;
  }
  return bk;
}

// ---------------------------------------------------------------------------
// BESSEL family entry points
// ---------------------------------------------------------------------------

/// BESSELJ(x, n) = J_n(x), defined on all real x. Delegates to libm `::jn`,
/// which agrees with Mac Excel 365 across the oracle corpus to within 1 ULP.
Value BesselJ(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_real_arg(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  auto n = coerce_bessel_order(args[1]);
  if (!n) {
    return Value::error(n.error());
  }
  const double r = ::jn(n.value(), x.value());
  if (std::isnan(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

/// BESSELY(x, n) = Y_n(x), singular at x=0 and undefined (by Excel) for x<0.
/// Computed via the NR2 6.5 polynomial / asymptotic split; libm `::yn`
/// drifts from Mac by up to 1e-5 relative on small x and is not used.
Value BesselY(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_real_arg(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  auto n = coerce_bessel_order(args[1]);
  if (!n) {
    return Value::error(n.error());
  }
  const double xv = x.value();
  if (xv <= 0.0 || std::isnan(xv)) {
    return Value::error(ErrorCode::Num);
  }
  const double r = bessy(n.value(), xv);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

/// BESSELI(x, n) = I_n(x), defined on all real x. Computed via the NR2 6.6
/// Chebyshev approximations of I_0 / I_1 plus Miller's downward recurrence
/// for n >= 2. The earlier power-series implementation accumulated 1e-5
/// relative error past x ~= 30 and is replaced.
Value BesselI(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_real_arg(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  auto n = coerce_bessel_order(args[1]);
  if (!n) {
    return Value::error(n.error());
  }
  const double xv = x.value();
  if (std::isnan(xv) || std::isinf(xv)) {
    return Value::error(ErrorCode::Num);
  }
  const double r = bessi(n.value(), xv);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

/// BESSELK(x, n) = K_n(x). Singular at x=0 and (by Excel) undefined for x<0.
/// Computed via NR2 6.6 polynomial / asymptotic K_0 / K_1 plus the forward
/// recurrence for n >= 2.
Value BesselK(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_real_arg(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  auto n = coerce_bessel_order(args[1]);
  if (!n) {
    return Value::error(n.error());
  }
  const double xv = x.value();
  if (xv <= 0.0 || std::isnan(xv)) {
    return Value::error(ErrorCode::Num);
  }
  const double r = bessk(n.value(), xv);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace

void register_engineering_special_builtins(FunctionRegistry& registry) {
  // ERF family. ERF is the sole 1-or-2-arg entry; the rest are strictly
  // 1-arg. The dispatcher enforces arity, so passing a second argument to
  // ERF.PRECISE / ERFC / ERFC.PRECISE surfaces #VALUE! before the impl runs.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"ERF", 1u, 2u, &Erf},
      {"ERF.PRECISE", 1u, 1u, &ErfPrecise},
      {"ERFC", 1u, 1u, &Erfc},
      {"ERFC.PRECISE", 1u, 1u, &ErfcPrecise},
      // BESSEL family. All are strict 2-arg (x, n).
      {"BESSELJ", 2u, 2u, &BesselJ},
      {"BESSELY", 2u, 2u, &BesselY},
      {"BESSELI", 2u, 2u, &BesselI},
      {"BESSELK", 2u, 2u, &BesselK},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
