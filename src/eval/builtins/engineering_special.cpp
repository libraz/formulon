// Copyright 2026 libraz. Licensed under the MIT License.
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
//   * BESSELJ / BESSELY delegate to POSIX `::jn` / `::yn` (available on
//     glibc / libc++ / Emscripten's musl).
//   * BESSELI is computed from the power series
//         I_n(x) = (x/2)^n * sum_{k>=0} (x/2)^(2k) / (k! (n+k)!)
//     with incremental term update and a 200-iteration safety cap.
//   * BESSELK is computed from the Abramowitz & Stegun 9.8.5-9.8.8
//     rational polynomial approximations for K_0 / K_1 on (0, 2] and
//     (2, +inf), then lifted to K_n via the forward recurrence
//     K_{n+1}(x) = (2n/x) K_n(x) + K_{n-1}(x). The A&S approximations are
//     specified to hold |epsilon| < ~1.9e-7; BESSELI here is accurate to
//     ~1e-13 in practice.
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
  // erf(lower). The spec for this milestone phrases the 2-arg form as
  // "the integral from lower to upper", matching that convention.
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
// BESSELJ / BESSELY (POSIX delegation)
// ---------------------------------------------------------------------------

// BESSELJ(x, n) = J_n(x), defined on all real x.
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

// BESSELY(x, n) = Y_n(x), singular at x=0 and undefined (by Excel) for x<0.
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
  const double r = ::yn(n.value(), xv);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ---------------------------------------------------------------------------
// BESSELI (power series)
// ---------------------------------------------------------------------------

// Computes I_n(x) via the canonical power series:
//   I_n(x) = (x/2)^n * sum_{k>=0} (x/2)^(2k) / (k! (n+k)!)
// `term_0 = 1 / n!`, `term_{k+1} = term_k * (x/2)^2 / ((k+1)(n+k+1))`.
// Terminates when the last term is below 1e-18 * |sum_so_far|, or after
// 200 iterations. Returns NaN if the sum does not converge (should not
// happen for any finite x within the engineering regime).
double bessel_i_series(int n, double x) {
  const double half = x / 2.0;
  const double half_sq = half * half;
  double factorial_n = 1.0;
  for (int i = 2; i <= n; ++i) {
    factorial_n *= static_cast<double>(i);
  }
  double term = 1.0 / factorial_n;
  double sum = term;
  for (int k = 0; k < 200; ++k) {
    term = term * half_sq / (static_cast<double>(k + 1) * static_cast<double>(n + k + 1));
    sum += term;
    if (std::fabs(term) < 1e-18 * std::fabs(sum)) {
      break;
    }
  }
  // Multiply by (x/2)^n. `std::pow` handles n=0 as 1 and preserves sign of
  // `half` for odd n, which gives the correct I_n(-x) = (-1)^n I_n(x).
  const double pref = std::pow(half, n);
  return pref * sum;
}

// BESSELI(x, n) = I_n(x). Defined on all real x.
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
  const double r = bessel_i_series(n.value(), xv);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ---------------------------------------------------------------------------
// BESSELK (A&S 9.8.5-9.8.8 + recurrence)
// ---------------------------------------------------------------------------

// Polynomial approximation of I_0(x) for |x| <= 3.75 (A&S 9.8.1). Used
// inside the small-x branch of K_0 / K_1.
double i0_small(double x) {
  const double t = x / 3.75;
  const double t2 = t * t;
  return 1.0 +
         t2 * (3.5156229 + t2 * (3.0899424 + t2 * (1.2067492 + t2 * (0.2659732 + t2 * (0.0360768 + t2 * 0.0045813)))));
}

// Polynomial approximation of I_1(x) / x for |x| <= 3.75 (A&S 9.8.3).
// Used inside the small-x branch of K_1.
double i1_over_x_small(double x) {
  const double t = x / 3.75;
  const double t2 = t * t;
  return 0.5 + t2 * (0.87890594 +
                     t2 * (0.51498869 + t2 * (0.15084934 + t2 * (0.02658733 + t2 * (0.00301532 + t2 * 0.00032411)))));
}

// K_0(x), x > 0. A&S 9.8.5 for 0 < x <= 2; A&S 9.8.6 for x > 2.
double bessel_k0(double x) {
  if (x <= 2.0) {
    const double y = x * x / 4.0;
    const double i0 = i0_small(x);
    return -std::log(x / 2.0) * i0 +
           (-0.57721566 +
            y * (0.42278420 +
                 y * (0.23069756 + y * (0.03488590 + y * (0.00262698 + y * (0.00010750 + y * 0.00000740))))));
  }
  const double y = 2.0 / x;
  return (std::exp(-x) / std::sqrt(x)) *
         (1.25331414 +
          y * (-0.07832358 +
               y * (0.02189568 + y * (-0.01062446 + y * (0.00587872 + y * (-0.00251540 + y * 0.00053208))))));
}

// K_1(x), x > 0. A&S 9.8.7 for 0 < x <= 2; A&S 9.8.8 for x > 2.
double bessel_k1(double x) {
  if (x <= 2.0) {
    const double y = x * x / 4.0;
    const double i1 = x * i1_over_x_small(x);
    return std::log(x / 2.0) * i1 +
           (1.0 / x) * (1.0 + y * (0.15443144 +
                                   y * (-0.67278579 +
                                        y * (-0.18156897 + y * (-0.01919402 + y * (-0.00110404 + y * -0.00004686))))));
  }
  const double y = 2.0 / x;
  return (std::exp(-x) / std::sqrt(x)) *
         (1.25331414 +
          y * (0.23498619 +
               y * (-0.03655620 + y * (0.01504268 + y * (-0.00780353 + y * (0.00325614 + y * -0.00068245))))));
}

// K_n(x), x > 0, via the forward recurrence
//   K_{n+1}(x) = (2n / x) K_n(x) + K_{n-1}(x).
double bessel_k(int n, double x) {
  if (n == 0) {
    return bessel_k0(x);
  }
  if (n == 1) {
    return bessel_k1(x);
  }
  double km1 = bessel_k0(x);
  double k = bessel_k1(x);
  for (int i = 1; i < n; ++i) {
    const double kp1 = (2.0 * static_cast<double>(i) / x) * k + km1;
    km1 = k;
    k = kp1;
  }
  return k;
}

// BESSELK(x, n) = K_n(x). Singular at x=0 and (by Excel) undefined for x<0.
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
  const double r = bessel_k(n.value(), xv);
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
  registry.register_function(FunctionDef{"ERF", 1u, 2u, &Erf});
  registry.register_function(FunctionDef{"ERF.PRECISE", 1u, 1u, &ErfPrecise});
  registry.register_function(FunctionDef{"ERFC", 1u, 1u, &Erfc});
  registry.register_function(FunctionDef{"ERFC.PRECISE", 1u, 1u, &ErfcPrecise});

  // BESSEL family. All are strict 2-arg (x, n).
  registry.register_function(FunctionDef{"BESSELJ", 2u, 2u, &BesselJ});
  registry.register_function(FunctionDef{"BESSELY", 2u, 2u, &BesselY});
  registry.register_function(FunctionDef{"BESSELI", 2u, 2u, &BesselI});
  registry.register_function(FunctionDef{"BESSELK", 2u, 2u, &BesselK});
}

}  // namespace eval
}  // namespace formulon
