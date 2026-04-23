// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's transcendental math built-in functions:
// EXP, LN, LOG, LOG10, PI, RADIANS, DEGREES, SIN, COS, TAN, ASIN, ACOS,
// ATAN, and ATAN2. Each impl follows the same recipe as the rest of the
// builtin catalog: coerce arguments via `eval/coerce.h`, propagate the
// left-most coercion error, and return a `Value`. Every function returns
// `#NUM!` for any non-finite result; trigonometric inputs are radians.

#include "eval/builtins_math_trig.h"

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
}

}  // namespace eval
}  // namespace formulon
