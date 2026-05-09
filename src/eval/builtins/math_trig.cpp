// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of Formulon's transcendental math built-in functions:
// EXP, LN, LOG, LOG10, PI, RADIANS, DEGREES, SIN, COS, TAN, ASIN, ACOS,
// ATAN, ATAN2, plus the reciprocal trig family (SEC, CSC, COT, ACOT) and
// the hyperbolic family (SINH, COSH, TANH, ASINH, ACOSH, ATANH, SECH,
// CSCH, COTH, ACOTH). Each impl follows the same recipe as the rest of
// the builtin catalog: coerce arguments via `eval/coerce.h`, propagate
// the left-most coercion error, and return a `Value`. Every function
// returns `#NUM!` for any non-finite result; trigonometric inputs are
// radians.

#include "eval/builtins/math_trig.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Excel uses an internally-stored value of pi rounded to ~15 significant
// digits. The hard-coded constant here is the same double precision value
// that `std::acos(-1.0)` would yield on any IEEE 754 system, which keeps
// `RADIANS(180) == kPi` exact.
static constexpr double kPi = 3.14159265358979323846;

// Shared kernel for unary numeric built-ins that follow the
// "coerce arg → std::*(x) → reject NaN/Inf as #NUM!" pattern. Used by
// EXP, the trigonometric primaries (SIN/COS/TAN/ATAN), and the entire
// hyperbolic family (SINH/COSH/TANH/ASINH). Functions with extra domain
// guards (LN/LOG/LOG10/ASIN/ACOS/ACOSH/ATANH) keep their own bodies
// because the guard can short-circuit before invoking the math function.
using DoubleFn = double (*)(double);
using DomainPredicate = bool (*)(double);

inline Value finite_math_number(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

inline Value apply_unary_math(DoubleFn fn, const Value* args) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  return finite_math_number(fn(x.value()));
}

inline Value apply_guarded_unary_math(DoubleFn fn, DomainPredicate domain, const Value* args) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (!domain(x.value())) {
    return Value::error(ErrorCode::Num);
  }
  return finite_math_number(fn(x.value()));
}

bool positive_domain(double x) {
  return x > 0.0;
}

bool closed_unit_domain(double x) {
  return x >= -1.0 && x <= 1.0;
}

bool at_least_one_domain(double x) {
  return x >= 1.0;
}

bool open_unit_domain(double x) {
  return x > -1.0 && x < 1.0;
}

bool outside_closed_unit_domain(double x) {
  return x < -1.0 || x > 1.0;
}

// EXP(x) - e raised to x. Overflow (e.g. EXP(1000)) produces +Inf, which is
// caught by the finite-check and surfaces as `#NUM!`. EXP(0) == 1.
Value Exp(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::exp, args);
}

// LN(x) - natural logarithm. Excel rejects `x <= 0` with `#NUM!`.
Value Ln(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_guarded_unary_math(&std::log, &positive_domain, args);
}

// LOG(x, [base]) - logarithm with an optional base (default 10). Excel
// quirks pinned here:
//   - `x <= 0`            -> `#NUM!`
//   - `base <= 0`         -> `#NUM!` (would-be `ln(base)` on a non-positive
//                            value already fails before the divide)
//   - `base == 1`         -> `#DIV/0!` (the divisor `ln(1)` is zero; Excel
//                            surfaces this distinct error code rather than
//                            `#NUM!`)
Value Log(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (x.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  double base = 10.0;
  if (arity >= 2) {
    auto parsed = coerce_to_number(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    base = parsed.value();
  }
  if (base <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double denom = std::log(base);
  if (denom == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = std::log(x.value()) / denom;
  return finite_math_number(r);
}

// LOG10(x) - base-10 logarithm. `x <= 0` -> `#NUM!`.
Value Log10(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_guarded_unary_math(&std::log10, &positive_domain, args);
}

// PI() - the constant pi. Zero-argument; the registry's arity check rejects
// any call with arguments before this body runs.
Value Pi(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::number(kPi);
}

// RADIANS(degrees) - degrees-to-radians conversion. RADIANS(0) == 0,
// RADIANS(180) == pi.
Value Radians(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = x.value() * kPi / 180.0;
  return finite_math_number(r);
}

// DEGREES(radians) - radians-to-degrees conversion. DEGREES(pi) == 180.
// Order of operations matches Mac Excel 365 / IronCalc: divide first, then
// multiply. The mathematically equivalent `x * 180.0 / kPi` form differs by
// 1 ULP for some inputs (e.g. 12345678900); see Mac probe 2026-05-02.
Value Degrees(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = x.value() / kPi * 180.0;
  return finite_math_number(r);
}

// SIN(x) - sine in radians. Excel imposes no domain restriction; only
// non-finite results (essentially impossible for finite input) surface
// `#NUM!`.
Value Sin(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::sin, args);
}

// COS(x) - cosine in radians.
Value Cos(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::cos, args);
}

// TAN(x) - tangent in radians. Excel quirk pin: even at the pole
// `TAN(PI/2)` the return value is a very large but FINITE number (because
// PI/2 in double precision differs slightly from the mathematical pole),
// so we do NOT pre-reject pole-adjacent inputs - only true Inf/NaN is
// reported as `#NUM!`.
Value Tan(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::tan, args);
}

// ASIN(x) - arcsine. Domain [-1, 1]; outside -> `#NUM!`. Result in
// [-pi/2, pi/2] radians.
Value Asin(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_guarded_unary_math(&std::asin, &closed_unit_domain, args);
}

// ACOS(x) - arccosine. Domain [-1, 1]; outside -> `#NUM!`. Result in
// [0, pi] radians.
Value Acos(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_guarded_unary_math(&std::acos, &closed_unit_domain, args);
}

// ATAN(x) - arctangent. No domain restriction. Result in (-pi/2, pi/2)
// radians.
Value Atan(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::atan, args);
}

// ATAN2(x, y) - two-argument arctangent. Excel's argument order is
// `(x, y)`, the OPPOSITE of C's `std::atan2(y, x)`. We pass them through
// swapped so callers see Excel semantics. When BOTH x and y are zero,
// Excel returns `#DIV/0!` even though `std::atan2(0, 0)` is defined as 0.
// Result in (-pi, pi] radians.
Value Atan2(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  auto y = coerce_to_number(args[1]);
  if (!y) {
    return Value::error(y.error());
  }
  if (x.value() == 0.0 && y.value() == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = std::atan2(y.value(), x.value());
  return finite_math_number(r);
}

// --- Hyperbolic functions ----------------------------------------------
//
// All inputs are unrestricted reals except where noted. The ASINH / ATANH
// inverses constrain their domains (ATANH on `(-1, 1)`, ACOSH on `[1, +inf)`).
// Every function guards the final double against NaN / Inf (which, for the
// forward hyperbolics SINH / COSH, can arise from overflow at `|x|` beyond
// roughly 710) and surfaces `#NUM!` in that case, mirroring ACOS / ASIN.

// SINH(x) - hyperbolic sine. Overflow (|x| beyond ~710) yields +/-Inf from
// std::sinh, caught by the finite-check and reported as `#NUM!`.
Value Sinh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::sinh, args);
}

// COSH(x) - hyperbolic cosine. Overflow yields +Inf -> `#NUM!`.
Value Cosh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::cosh, args);
}

// TANH(x) - hyperbolic tangent. Asymptotic to +/-1; no domain restriction.
Value Tanh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::tanh, args);
}

// ASINH(x) - inverse hyperbolic sine. Domain: all reals.
Value Asinh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_unary_math(&std::asinh, args);
}

// ACOSH(x) - inverse hyperbolic cosine. Domain: `[1, +inf)`. `x < 1` is
// outside the domain and Excel reports `#NUM!`; this mirrors the ACOS
// domain check above.
Value Acosh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_guarded_unary_math(&std::acosh, &at_least_one_domain, args);
}

// ATANH(x) - inverse hyperbolic tangent. Domain: `(-1, 1)` EXCLUSIVE. At
// `|x| == 1` the result is +/-Inf; `|x| > 1` produces NaN. Both are folded
// to `#NUM!` - matching Excel, which rejects the closed endpoints as well.
Value Atanh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return apply_guarded_unary_math(&std::atanh, &open_unit_domain, args);
}

// --- Reciprocal trigonometric functions ---------------------------------
//
// These are one-liners on top of the primary trig functions. The only
// interesting edge case is the divisor: `SEC(PI/2)` and `CSC(0)` etc.
// technically sit at poles. For inputs where the primary trig function
// returns *exactly* zero (e.g. `SIN(0) == 0`), we surface `#DIV/0!`,
// matching Excel. For inputs that only approach zero (`COS(PI/2)` is a
// very small but non-zero double), the reciprocal is finite; this matches
// the TAN(PI/2) quirk pinned above.

// SEC(x) - secant, `1 / cos(x)`. `cos(x) == 0` -> `#DIV/0!`.
Value Sec(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double c = std::cos(x.value());
  if (c == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = 1.0 / c;
  return finite_math_number(r);
}

// CSC(x) - cosecant, `1 / sin(x)`. `sin(x) == 0` -> `#DIV/0!`.
Value Csc(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double s = std::sin(x.value());
  if (s == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = 1.0 / s;
  return finite_math_number(r);
}

// COT(x) - cotangent, `cos(x) / sin(x)` (equivalently `1 / tan(x)`).
// `sin(x) == 0` -> `#DIV/0!`. Implemented as `cos / sin` rather than
// `1 / tan` because `tan(x)` can round exactly to 0 at multiples of PI
// even when `sin(x)` is non-zero on the same argument, which would miss
// the divide-by-zero case.
Value Cot(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double s = std::sin(x.value());
  if (s == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = std::cos(x.value()) / s;
  return finite_math_number(r);
}

// ACOT(x) - inverse cotangent, `PI/2 - atan(x)`. Range `(0, PI)`. No
// domain restriction. Note Excel's ACOT does NOT follow the `atan(1/x)`
// definition on `x < 0`; `PI/2 - atan(x)` is the correct formula.
Value Acot(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = kPi / 2.0 - std::atan(x.value());
  return finite_math_number(r);
}

// SECH(x) - hyperbolic secant, `1 / cosh(x)`. `cosh` is always >= 1, so
// the divisor never hits zero; very large `|x|` drives the quotient to 0
// (not Inf), which is a valid return - Excel yields 0 for `SECH(1000)`.
Value Sech(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double c = std::cosh(x.value());
  if (std::isinf(c)) {
    // Overflow in cosh: 1 / Inf == 0 in IEEE arithmetic, but we would
    // rather surface an explicit zero to match Excel's observed output.
    return Value::number(0.0);
  }
  const double r = 1.0 / c;
  return finite_math_number(r);
}

// CSCH(x) - hyperbolic cosecant, `1 / sinh(x)`. `sinh(0) == 0` -> `#DIV/0!`.
// Large `|x|` drives sinh to +/-Inf, which yields 0 (matching Excel).
Value Csch(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double s = std::sinh(x.value());
  if (s == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  if (std::isinf(s)) {
    return Value::number(0.0);
  }
  const double r = 1.0 / s;
  return finite_math_number(r);
}

// COTH(x) - hyperbolic cotangent, `cosh(x) / sinh(x)`. `sinh(0) == 0`
// -> `#DIV/0!`; asymptotic to +/-1 as `|x| -> inf`.
Value Coth(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double s = std::sinh(x.value());
  if (s == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  if (std::isinf(s)) {
    // sinh and cosh overflow together; the ratio tends to +/-1.
    return Value::number((x.value() > 0.0) ? 1.0 : -1.0);
  }
  const double r = std::cosh(x.value()) / s;
  return finite_math_number(r);
}

// ACOTH(x) - inverse hyperbolic cotangent, `atanh(1 / x)`. Domain is
// `|x| > 1` strictly; `|x| <= 1` yields `#NUM!`. `std::atanh(1/x)` returns
// +/-Inf at the endpoints, so the explicit guard provides a cleaner
// contract on the boundary.
Value Acoth(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (!outside_closed_unit_domain(x.value())) {
    return Value::error(ErrorCode::Num);
  }
  return finite_math_number(std::atanh(1.0 / x.value()));
}

}  // namespace

void register_math_trig_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"EXP", 1u, 1u, &Exp},   {"LN", 1u, 1u, &Ln},           {"LOG", 1u, 2u, &Log},         {"LOG10", 1u, 1u, &Log10},
      {"PI", 0u, 0u, &Pi},     {"RADIANS", 1u, 1u, &Radians}, {"DEGREES", 1u, 1u, &Degrees}, {"SIN", 1u, 1u, &Sin},
      {"COS", 1u, 1u, &Cos},   {"TAN", 1u, 1u, &Tan},         {"ASIN", 1u, 1u, &Asin},       {"ACOS", 1u, 1u, &Acos},
      {"ATAN", 1u, 1u, &Atan}, {"ATAN2", 2u, 2u, &Atan2},     {"SINH", 1u, 1u, &Sinh},       {"COSH", 1u, 1u, &Cosh},
      {"TANH", 1u, 1u, &Tanh}, {"ASINH", 1u, 1u, &Asinh},     {"ACOSH", 1u, 1u, &Acosh},     {"ATANH", 1u, 1u, &Atanh},
      {"SEC", 1u, 1u, &Sec},   {"CSC", 1u, 1u, &Csc},         {"COT", 1u, 1u, &Cot},         {"ACOT", 1u, 1u, &Acot},
      {"SECH", 1u, 1u, &Sech}, {"CSCH", 1u, 1u, &Csch},       {"COTH", 1u, 1u, &Coth},       {"ACOTH", 1u, 1u, &Acoth},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
