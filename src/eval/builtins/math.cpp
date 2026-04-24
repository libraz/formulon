// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's arithmetic and rounding built-in functions:
// ABS, SIGN, INT, TRUNC, SQRT, MOD, POWER, ROUND, ROUNDDOWN, ROUNDUP,
// EVEN, ODD, and QUOTIENT. Each impl follows the same recipe as the rest
// of the builtin catalog: coerce arguments via `eval/coerce.h`, propagate
// the left-most coercion error, and return a `Value`.

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
//
// IEEE 754 arithmetic on decimal-input expressions often lands 1-2 ULPs
// below the mathematical .5 boundary (e.g.
// `1.05 * (0.0284 + 0.0046) - 0.0284` = 0.006249999999999999, so
// `value * 10000 = 62.499999999999986`). Mac Excel 365 compensates by
// snapping these near-half values back to the decimal-correct side.
// We mirror that with a tiny relative nudge (`2 * 2^-52 * |scaled|`,
// ~9e-16 per unit) — large enough to absorb the typical ULP-level error
// but far below the distance between any two genuinely distinct decimal
// halves at the same scale, so values that are truly off the boundary
// are unaffected.
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
  const double scaled = value.value() * factor;
  const double bias = std::copysign(std::fabs(scaled) * 2.0e-14, scaled);
  const double r = std::round(scaled + bias) / factor;
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

// --- Significance-aware rounding (legacy CEILING / FLOOR / MROUND) ------
//
// The legacy forms share one awkward quirk: if `number` and `significance`
// have opposite signs (and neither is zero), Excel returns #NUM!. The
// modern `*.MATH` variants (below) drop this rule and use `|significance|`
// unconditionally. `MROUND` inherits the opposite-sign #NUM! rule.

// CEILING / FLOOR / MROUND treat empty-string text as `#VALUE!` rather
// than coercing it to zero (which is what the general arithmetic rule in
// `coerce_to_number` does). This matters because the subsequent `s == 0`
// branch would otherwise convert `FLOOR(10, "")` to `#DIV/0!` -- Excel's
// observed behaviour is `#VALUE!`.
inline bool is_empty_text(const Value& v) {
  if (v.kind() != ValueKind::Text) return false;
  const std::string_view t = v.as_text();
  for (char c : t) {
    if (c != ' ' && c != '\t') return false;
  }
  return true;
}

// Snap a floating-point quotient to its nearest integer when the deviation
// is within ~a few ULPs. CEILING / FLOOR would otherwise return `n - s`
// for inputs like `FLOOR(7.1, 0.1)` because 7.1 / 0.1 in IEEE-754 is
// 70.999... rather than exactly 71; Excel recognises this as an exact
// multiple and returns `n` itself. The tolerance is small enough that
// genuinely small-but-nonzero quotients (e.g. `CEILING(1e-12, 1)` where
// the quotient is 1e-12, far more than 1 ULP away from 0) are not
// mistaken for exact integers.
inline double snap_to_integer(double q) {
  const double nearest = std::round(q);
  // Tolerance scales with |q| because FP error in `n / s` is proportional
  // to |q|. A relative tolerance of 2e-15 (~9 ULPs) is tight enough that
  // legitimate non-integer quotients like `9999999999999 / 0.21`
  // (fractional part 0.143) are not snapped, while absorbing the 5-ULP
  // error in `7.1 / 0.1`. No `max(1.0, ...)` floor: snapping a tiny
  // non-zero quotient to 0 would corrupt `CEILING(1e-12, 1) -> 1`.
  if (std::fabs(q - nearest) < 2e-15 * std::fabs(q)) {
    return nearest;
  }
  return q;
}

inline double signum(double x) {
  if (x > 0.0) {
    return 1.0;
  }
  if (x < 0.0) {
    return -1.0;
  }
  return 0.0;
}

// CEILING(number, significance) - legacy: nearest multiple of
// `|significance|` in the direction determined by sign matching:
//
//   * number == 0  -> 0
//   * significance == 0 -> 0 (no #DIV/0!, unlike FLOOR)
//   * sign(number) == sign(significance)  -> round AWAY from zero
//     (magnitude ceil on the positive half-line).
//   * sign(number) != sign(significance)  -> round toward +infinity
//     (math ceil on the signed value, matching CEILING.MATH defaults).
//
// Excel 365 Mac no longer returns #NUM! on sign mismatch; that behaviour
// was the pre-365 legacy rule. See the corresponding oracle suite for the
// exact values used as reference.
Value Ceiling(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  if (is_empty_text(args[0]) || is_empty_text(args[1])) {
    return Value::error(ErrorCode::Value);
  }
  auto number = coerce_to_number(args[0]);
  if (!number) {
    return Value::error(number.error());
  }
  auto significance = coerce_to_number(args[1]);
  if (!significance) {
    return Value::error(significance.error());
  }
  const double n = number.value();
  const double s = significance.value();
  if (n == 0.0) {
    return Value::number(0.0);
  }
  if (s == 0.0) {
    // Legacy CEILING: significance of zero yields zero (no #DIV/0!).
    return Value::number(0.0);
  }
  const double abs_s = std::fabs(s);
  // Matching signs: magnitude away-from-zero, then restore sign.
  // Mismatched signs: math ceiling on the signed value. `snap_to_integer`
  // absorbs the floating-point noise that would otherwise make e.g.
  // `CEILING(-7.1, 0.1)` return `-7` instead of `-7.1`.
  const double r = (signum(n) == signum(s))
                       ? signum(n) * std::ceil(snap_to_integer(std::fabs(n) / abs_s)) * abs_s
                       : std::ceil(snap_to_integer(n / abs_s)) * abs_s;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// FLOOR(number, significance) - legacy: nearest multiple of
// `|significance|` in the direction determined by sign matching. Mirror
// image of CEILING above, with two differences:
//
//   * significance == 0 -> #DIV/0! (Excel-documented quirk).
//   * Direction in mismatched-sign case is toward -infinity (math floor).
//
// As with CEILING, Excel 365 Mac no longer treats a sign mismatch as
// #NUM!; this impl tracks the current oracle.
Value Floor(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  if (is_empty_text(args[0]) || is_empty_text(args[1])) {
    return Value::error(ErrorCode::Value);
  }
  auto number = coerce_to_number(args[0]);
  if (!number) {
    return Value::error(number.error());
  }
  auto significance = coerce_to_number(args[1]);
  if (!significance) {
    return Value::error(significance.error());
  }
  const double n = number.value();
  const double s = significance.value();
  if (n == 0.0) {
    return Value::number(0.0);
  }
  if (s == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double abs_s = std::fabs(s);
  // Matching signs: magnitude toward-zero, then restore sign.
  // Mismatched signs: math floor on the signed value. `snap_to_integer`
  // absorbs the floating-point noise that would otherwise make e.g.
  // `FLOOR(7.1, 0.1)` return `7` instead of `7.1`.
  const double r = (signum(n) == signum(s))
                       ? signum(n) * std::floor(snap_to_integer(std::fabs(n) / abs_s)) * abs_s
                       : std::floor(snap_to_integer(n / abs_s)) * abs_s;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// MROUND(number, multiple) - nearest multiple of `|multiple|` to `number`,
// with ties rounded away from zero. Opposite-signed inputs yield #NUM!;
// `multiple = 0` returns 0 (Excel's documented quirk).
Value MRound(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto number = coerce_to_number(args[0]);
  if (!number) {
    return Value::error(number.error());
  }
  auto multiple = coerce_to_number(args[1]);
  if (!multiple) {
    return Value::error(multiple.error());
  }
  const double n = number.value();
  const double m = multiple.value();
  if (m == 0.0) {
    return Value::number(0.0);
  }
  if (n != 0.0 && signum(n) != signum(m)) {
    return Value::error(ErrorCode::Num);
  }
  const double abs_m = std::fabs(m);
  // `std::floor(x + 0.5)` implements round-half-away-from-zero on the
  // positive half-line; the outer `signum(n)` restores the sign.
  const double r = signum(n) * std::floor(std::fabs(n) / abs_m + 0.5) * abs_m;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// --- CEILING.MATH / FLOOR.MATH ------------------------------------------
//
// Modern variants: `significance` is always used as `|significance|`
// (negative values do NOT trigger #NUM!), and the optional `mode` arg
// only affects NEGATIVE inputs:
//   * mode = 0 (default): both round toward +infinity for CEILING.MATH,
//     toward -infinity for FLOOR.MATH.
//   * mode != 0: both round AWAY from zero
//     (CEILING.MATH: toward -infinity for negatives;
//      FLOOR.MATH: toward +infinity i.e. toward zero for negatives).

Value CeilingMath(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto number = coerce_to_number(args[0]);
  if (!number) {
    return Value::error(number.error());
  }
  double significance = 1.0;
  if (arity >= 2) {
    auto coerced = coerce_to_number(args[1]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    significance = coerced.value();
  }
  bool away_from_zero = false;
  if (arity >= 3) {
    auto coerced = coerce_to_number(args[2]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    away_from_zero = coerced.value() != 0.0;
  }
  const double n = number.value();
  if (n == 0.0) {
    return Value::number(0.0);
  }
  if (significance == 0.0) {
    return Value::number(0.0);
  }
  const double abs_s = std::fabs(significance);
  const double scaled = n / abs_s;
  // Positive inputs: always ceil. Negative inputs: ceil (toward +inf) for
  // default mode, floor (away from zero) when mode != 0.
  const double rounded = (n > 0.0 || !away_from_zero) ? std::ceil(scaled) : std::floor(scaled);
  const double r = rounded * abs_s;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value FloorMath(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto number = coerce_to_number(args[0]);
  if (!number) {
    return Value::error(number.error());
  }
  double significance = 1.0;
  if (arity >= 2) {
    auto coerced = coerce_to_number(args[1]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    significance = coerced.value();
  }
  bool toward_zero = false;
  if (arity >= 3) {
    auto coerced = coerce_to_number(args[2]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    toward_zero = coerced.value() != 0.0;
  }
  const double n = number.value();
  if (n == 0.0) {
    return Value::number(0.0);
  }
  if (significance == 0.0) {
    return Value::number(0.0);
  }
  const double abs_s = std::fabs(significance);
  const double scaled = n / abs_s;
  // Positive inputs: always floor. Negative inputs: floor (toward -inf)
  // for default mode, ceil (toward zero) when mode != 0.
  const double rounded = (n > 0.0 || !toward_zero) ? std::floor(scaled) : std::ceil(scaled);
  const double r = rounded * abs_s;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// --- Parity-aware rounding ----------------------------------------------
//
// EVEN and ODD both round AWAY from zero to the nearest integer of the
// required parity. ODD has a single documented special case: `ODD(0) = 1`
// (because 0 has no odd neighbour in the "nearest" direction, Excel
// promotes it to +1). All other zero / non-integer inputs behave as a
// normal away-from-zero rounding followed by a parity fix-up.

// EVEN(x) - nearest even integer, rounded AWAY from zero.
//   `EVEN(1.5) = 2`, `EVEN(3) = 4`, `EVEN(-1.5) = -2`, `EVEN(-2.1) = -4`,
//   `EVEN(0) = 0`.
Value Even(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  if (x == 0.0) {
    return Value::number(0.0);
  }
  // Round AWAY from zero to the nearest integer first, then step one away
  // again if that integer happens to be odd. `std::ceil / std::floor` on
  // the signed value implements the away-from-zero step.
  const double away = (x > 0.0) ? std::ceil(x) : std::floor(x);
  // Parity check via std::fmod: `|away| mod 2 == 0` -> already even.
  const double parity = std::fmod(std::fabs(away), 2.0);
  const double r = (parity == 0.0) ? away : away + ((x > 0.0) ? 1.0 : -1.0);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// ODD(x) - nearest odd integer, rounded AWAY from zero. `ODD(0) = 1` is the
// documented quirk; otherwise symmetric to EVEN.
//   `ODD(1.5) = 3`, `ODD(2) = 3`, `ODD(-1.5) = -3`, `ODD(-2) = -3`.
Value Odd(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  if (x == 0.0) {
    return Value::number(1.0);
  }
  const double away = (x > 0.0) ? std::ceil(x) : std::floor(x);
  const double parity = std::fmod(std::fabs(away), 2.0);
  // `|away| mod 2 == 1` -> already odd. Otherwise step one away from zero.
  const double r = (parity != 0.0) ? away : away + ((x > 0.0) ? 1.0 : -1.0);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// QUOTIENT(numerator, denominator) - integer division, truncated TOWARD
// zero (not floor-division). Matches Excel: `QUOTIENT(-10, 3) = -3`.
// `denominator == 0` -> `#DIV/0!`. Very large quotients fall through the
// finite-check and surface `#NUM!`.
Value Quotient(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto num = coerce_to_number(args[0]);
  if (!num) {
    return Value::error(num.error());
  }
  auto den = coerce_to_number(args[1]);
  if (!den) {
    return Value::error(den.error());
  }
  if (den.value() == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = std::trunc(num.value() / den.value());
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
  registry.register_function(FunctionDef{"QUOTIENT", 2u, 2u, &Quotient});

  // Rounding (all take value + digits).
  registry.register_function(FunctionDef{"ROUND", 2u, 2u, &Round});
  registry.register_function(FunctionDef{"ROUNDDOWN", 2u, 2u, &RoundDown});
  registry.register_function(FunctionDef{"ROUNDUP", 2u, 2u, &RoundUp});

  // Parity-aware rounding.
  registry.register_function(FunctionDef{"EVEN", 1u, 1u, &Even});
  registry.register_function(FunctionDef{"ODD", 1u, 1u, &Odd});

  // Significance-aware rounding.
  registry.register_function(FunctionDef{"CEILING", 2u, 2u, &Ceiling});
  registry.register_function(FunctionDef{"FLOOR", 2u, 2u, &Floor});
  registry.register_function(FunctionDef{"MROUND", 2u, 2u, &MRound});
  registry.register_function(FunctionDef{"CEILING.MATH", 1u, 3u, &CeilingMath});
  registry.register_function(FunctionDef{"FLOOR.MATH", 1u, 3u, &FloorMath});
}

}  // namespace eval
}  // namespace formulon
