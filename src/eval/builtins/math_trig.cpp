// Copyright 2026 libraz. Licensed under the MIT License.
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

// EXP(x) - e raised to x. Overflow (e.g. EXP(1000)) produces +Inf, which is
// caught by the finite-check and surfaces as `#NUM!`. EXP(0) == 1.
Value Exp(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::exp(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// LN(x) - natural logarithm. Excel rejects `x <= 0` with `#NUM!`.
Value Ln(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (x.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::log(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// LOG10(x) - base-10 logarithm. `x <= 0` -> `#NUM!`.
Value Log10(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (x.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::log10(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// DEGREES(radians) - radians-to-degrees conversion. DEGREES(pi) == 180.
Value Degrees(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = x.value() * 180.0 / kPi;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// SIN(x) - sine in radians. Excel imposes no domain restriction; only
// non-finite results (essentially impossible for finite input) surface
// `#NUM!`.
Value Sin(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::sin(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// COS(x) - cosine in radians.
Value Cos(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::cos(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// TAN(x) - tangent in radians. Excel quirk pin: even at the pole
// `TAN(PI/2)` the return value is a very large but FINITE number (because
// PI/2 in double precision differs slightly from the mathematical pole),
// so we do NOT pre-reject pole-adjacent inputs - only true Inf/NaN is
// reported as `#NUM!`.
Value Tan(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::tan(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ASIN(x) - arcsine. Domain [-1, 1]; outside -> `#NUM!`. Result in
// [-pi/2, pi/2] radians.
Value Asin(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (x.value() < -1.0 || x.value() > 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::asin(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ACOS(x) - arccosine. Domain [-1, 1]; outside -> `#NUM!`. Result in
// [0, pi] radians.
Value Acos(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (x.value() < -1.0 || x.value() > 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::acos(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ATAN(x) - arctangent. No domain restriction. Result in (-pi/2, pi/2)
// radians.
Value Atan(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::atan(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::sinh(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// COSH(x) - hyperbolic cosine. Overflow yields +Inf -> `#NUM!`.
Value Cosh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::cosh(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// TANH(x) - hyperbolic tangent. Asymptotic to +/-1; no domain restriction.
Value Tanh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::tanh(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ASINH(x) - inverse hyperbolic sine. Domain: all reals.
Value Asinh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  const double r = std::asinh(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ACOSH(x) - inverse hyperbolic cosine. Domain: `[1, +inf)`. `x < 1` is
// outside the domain and Excel reports `#NUM!`; this mirrors the ACOS
// domain check above.
Value Acosh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (x.value() < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::acosh(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ATANH(x) - inverse hyperbolic tangent. Domain: `(-1, 1)` EXCLUSIVE. At
// `|x| == 1` the result is +/-Inf; `|x| > 1` produces NaN. Both are folded
// to `#NUM!` - matching Excel, which rejects the closed endpoints as well.
Value Atanh(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = coerce_to_number(args[0]);
  if (!x) {
    return Value::error(x.error());
  }
  if (x.value() <= -1.0 || x.value() >= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::atanh(x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
  if (x.value() >= -1.0 && x.value() <= 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::atanh(1.0 / x.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace

void register_math_trig_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"EXP", 1u, 1u, &Exp});
  registry.register_function(FunctionDef{"LN", 1u, 1u, &Ln});
  registry.register_function(FunctionDef{"LOG", 1u, 2u, &Log});
  registry.register_function(FunctionDef{"LOG10", 1u, 1u, &Log10});
  registry.register_function(FunctionDef{"PI", 0u, 0u, &Pi});
  registry.register_function(FunctionDef{"RADIANS", 1u, 1u, &Radians});
  registry.register_function(FunctionDef{"DEGREES", 1u, 1u, &Degrees});
  registry.register_function(FunctionDef{"SIN", 1u, 1u, &Sin});
  registry.register_function(FunctionDef{"COS", 1u, 1u, &Cos});
  registry.register_function(FunctionDef{"TAN", 1u, 1u, &Tan});
  registry.register_function(FunctionDef{"ASIN", 1u, 1u, &Asin});
  registry.register_function(FunctionDef{"ACOS", 1u, 1u, &Acos});
  registry.register_function(FunctionDef{"ATAN", 1u, 1u, &Atan});
  registry.register_function(FunctionDef{"ATAN2", 2u, 2u, &Atan2});
  // Hyperbolic family.
  registry.register_function(FunctionDef{"SINH", 1u, 1u, &Sinh});
  registry.register_function(FunctionDef{"COSH", 1u, 1u, &Cosh});
  registry.register_function(FunctionDef{"TANH", 1u, 1u, &Tanh});
  registry.register_function(FunctionDef{"ASINH", 1u, 1u, &Asinh});
  registry.register_function(FunctionDef{"ACOSH", 1u, 1u, &Acosh});
  registry.register_function(FunctionDef{"ATANH", 1u, 1u, &Atanh});
  // Reciprocal trig + inverse cotangent.
  registry.register_function(FunctionDef{"SEC", 1u, 1u, &Sec});
  registry.register_function(FunctionDef{"CSC", 1u, 1u, &Csc});
  registry.register_function(FunctionDef{"COT", 1u, 1u, &Cot});
  registry.register_function(FunctionDef{"ACOT", 1u, 1u, &Acot});
  // Reciprocal hyperbolic + inverse coth.
  registry.register_function(FunctionDef{"SECH", 1u, 1u, &Sech});
  registry.register_function(FunctionDef{"CSCH", 1u, 1u, &Csch});
  registry.register_function(FunctionDef{"COTH", 1u, 1u, &Coth});
  registry.register_function(FunctionDef{"ACOTH", 1u, 1u, &Acoth});
}

}  // namespace eval
}  // namespace formulon
