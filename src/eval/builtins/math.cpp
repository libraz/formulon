// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's arithmetic and rounding built-in functions:
// ABS, SIGN, INT, TRUNC, SQRT, MOD, POWER, ROUND, ROUNDDOWN, and ROUNDUP.
// Each impl follows the same recipe as the rest of the builtin catalog:
// coerce arguments via `eval/coerce.h`, propagate the left-most coercion
// error, and return a `Value`.

#include "eval/builtins/math.h"

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

// --- Single-number transforms -------------------------------------------

// ABS(value) - absolute value. Coerces the single argument to a number.
Value Abs(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  return Value::number(std::fabs(coerced.value()));
}

// SIGN(value) - returns -1, 0, or +1 depending on the sign of the input.
// `SIGN(-0.0)` returns 0 (the +/- distinction on zero is not preserved).
Value Sign(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  if (x > 0.0) {
    return Value::number(1.0);
  }
  if (x < 0.0) {
    return Value::number(-1.0);
  }
  return Value::number(0.0);
}

// INT(value) - floor toward negative infinity. Excel's documented behavior:
// `INT(2.7) = 2`, `INT(-2.7) = -3`. Implemented with `std::floor`, NOT
// `std::trunc` (the latter would round toward zero and break the negative
// case).
Value Int_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  return Value::number(std::floor(coerced.value()));
}

// Helper: read the optional `digits` argument of TRUNC / ROUND-family.
// Returns the integer count of decimal places, an `ErrorCode` if the
// argument cannot be coerced or is non-finite (NaN/Inf), or 0 when no
// second argument is supplied.
Expected<int, ErrorCode> read_digits(const Value* args, std::uint32_t arity, std::uint32_t index) {
  if (arity <= index) {
    return 0;
  }
  auto coerced = coerce_to_number(args[index]);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return static_cast<int>(std::trunc(d));
}

// TRUNC(value, digits?) - truncate toward zero. With no second arg or
// `digits = 0`, equivalent to `std::trunc(value)`. With `digits != 0`, the
// value is scaled by `10^digits`, truncated, then rescaled. `digits` may be
// negative (e.g. `TRUNC(1234.5, -1) = 1230`). A non-finite scale factor
// (caused by very large `|digits|`) yields `#NUM!`.
Value Trunc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto value = coerce_to_number(args[0]);
  if (!value) {
    return Value::error(value.error());
  }
  auto digits = read_digits(args, arity, 1);
  if (!digits) {
    return Value::error(digits.error());
  }
  const double factor = std::pow(10.0, digits.value());
  if (std::isnan(factor) || std::isinf(factor)) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::trunc(value.value() * factor) / factor;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// SQRT(value) - square root. Negative input -> `#NUM!`.
Value Sqrt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  if (x < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::sqrt(x);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// --- Two-argument numeric -----------------------------------------------

// MOD(n, d) - Excel's modulo. The result has the SIGN OF THE DIVISOR, not
// the C `%` semantics. Formula: `n - d * floor(n / d)`. So `MOD(-7, 3) = 2`,
// `MOD(7, -3) = -2`. `MOD(n, 0)` -> `#DIV/0!`. `std::fmod` is intentionally
// avoided here because it inherits C semantics (sign of dividend).
Value Mod(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n = coerce_to_number(args[0]);
  if (!n) {
    return Value::error(n.error());
  }
  auto d = coerce_to_number(args[1]);
  if (!d) {
    return Value::error(d.error());
  }
  if (d.value() == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = n.value() - d.value() * std::floor(n.value() / d.value());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// POWER(base, exp) - shares the `apply_pow` helper with the `^` operator
// so the two paths cannot diverge on edge cases (negative-base with a
// fractional exponent, overflow, `0^0`, etc.).
Value Power(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto base = coerce_to_number(args[0]);
  if (!base) {
    return Value::error(base.error());
  }
  auto exp = coerce_to_number(args[1]);
  if (!exp) {
    return Value::error(exp.error());
  }
  auto r = apply_pow(base.value(), exp.value());
  if (!r) {
    return Value::error(r.error());
  }
  return Value::number(r.value());
}

// --- Rounding -----------------------------------------------------------
//
// All three take `(value, digits)`. `digits` may be negative. The three
// rounding modes deliberately use distinct formulas (not a single param-
// terised function): the modes have different behaviour and inlining the
// formula keeps each impl trivially auditable.

// ROUND - round half away from zero. `std::round` matches this on every
// supported platform (it is mandated by C++11). `ROUND(2.5, 0) = 3`,
// `ROUND(-2.5, 0) = -3`.
Value Round(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto value = coerce_to_number(args[0]);
  if (!value) {
    return Value::error(value.error());
  }
  auto digits = read_digits(args, 2, 1);
  if (!digits) {
    return Value::error(digits.error());
  }
  const double factor = std::pow(10.0, digits.value());
  if (std::isnan(factor) || std::isinf(factor)) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::round(value.value() * factor) / factor;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ROUNDDOWN - always toward zero. `ROUNDDOWN(2.99, 0) = 2`,
// `ROUNDDOWN(-2.99, 0) = -2`.
Value RoundDown(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto value = coerce_to_number(args[0]);
  if (!value) {
    return Value::error(value.error());
  }
  auto digits = read_digits(args, 2, 1);
  if (!digits) {
    return Value::error(digits.error());
  }
  const double factor = std::pow(10.0, digits.value());
  if (std::isnan(factor) || std::isinf(factor)) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::trunc(value.value() * factor) / factor;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ROUNDUP - always away from zero. `ROUNDUP(2.01, 0) = 3`,
// `ROUNDUP(-2.01, 0) = -3`. Positive inputs use `std::ceil`, negative
// inputs use `std::floor`; zero round-trips through either branch.
Value RoundUp(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto value = coerce_to_number(args[0]);
  if (!value) {
    return Value::error(value.error());
  }
  auto digits = read_digits(args, 2, 1);
  if (!digits) {
    return Value::error(digits.error());
  }
  const double factor = std::pow(10.0, digits.value());
  if (std::isnan(factor) || std::isinf(factor)) {
    return Value::error(ErrorCode::Num);
  }
  const double scaled = value.value() * factor;
  const double r = (value.value() > 0.0) ? std::ceil(scaled) / factor : std::floor(scaled) / factor;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace

void register_math_builtins(FunctionRegistry& registry) {
  // Single-number transforms.
  registry.register_function(FunctionDef{"ABS", 1u, 1u, &Abs});
  registry.register_function(FunctionDef{"SIGN", 1u, 1u, &Sign});
  registry.register_function(FunctionDef{"INT", 1u, 1u, &Int_});
  registry.register_function(FunctionDef{"TRUNC", 1u, 2u, &Trunc});
  registry.register_function(FunctionDef{"SQRT", 1u, 1u, &Sqrt});

  // Two-argument numeric.
  registry.register_function(FunctionDef{"MOD", 2u, 2u, &Mod});
  registry.register_function(FunctionDef{"POWER", 2u, 2u, &Power});

  // Rounding (all take value + digits).
  registry.register_function(FunctionDef{"ROUND", 2u, 2u, &Round});
  registry.register_function(FunctionDef{"ROUNDDOWN", 2u, 2u, &RoundDown});
  registry.register_function(FunctionDef{"ROUNDUP", 2u, 2u, &RoundUp});
}

}  // namespace eval
}  // namespace formulon
