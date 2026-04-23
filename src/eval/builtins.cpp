// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's built-in functions. Each function follows
// the same recipe:
//
//   1. Coerce each argument via the helpers in `eval/coerce.h`.
//   2. Propagate the left-most coercion error.
//   3. Compute the result and finalize (interning text, checking finiteness).
//
// `IF`, `IFERROR`, and `IFNA` are intentionally absent: they short-circuit
// and are special-cased in the tree walker before the registry is consulted.

#include "eval/builtins.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <string>
#include <vector>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "eval/text_ops.h"
#include "eval/utf8_length.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// SUM(value, ...) --------------------------------------------------------
// Excel's SUM coerces each argument to a number; non-coercible text yields
// #VALUE! and any error among the inputs propagates left-to-right.
Value Sum(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    total += coerced.value();
  }
  if (std::isnan(total) || std::isinf(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

// CONCAT(value, ...) / CONCATENATE(value, ...) ---------------------------
// Both spellings share an implementation. Each argument is rendered via
// `coerce_to_text`; left-most error wins. The joined result is interned in
// the call's arena so the returned Value remains readable for the caller.
Value Concat(const Value* args, std::uint32_t arity, Arena& arena) {
  std::string joined;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_text(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    joined.append(coerced.value());
  }
  const std::string_view interned = arena.intern(joined);
  return Value::text(interned);
}

// TRUE() / FALSE() -------------------------------------------------------
// Both are zero-argument constants. Excel rejects any argument with #VALUE!,
// which the registry's arity check enforces (min=max=0). The body simply
// returns the corresponding boolean.
Value True_(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(true);
}

Value False_(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(false);
}

// NOT(value) -------------------------------------------------------------
// Coerces the single argument to bool and negates. Errors propagate (the
// dispatcher already short-circuits on argument errors before invoking
// this body, so by the time we run the input is non-error). A coercion
// failure (e.g. non-numeric text) surfaces as #VALUE!.
Value Not(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_bool(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  return Value::boolean(!coerced.value());
}

// AND(value, ...) / OR(value, ...) ---------------------------------------
// Both functions evaluate every argument (Excel does not logically
// short-circuit AND / OR; the only short-circuit is the dispatcher's
// left-most-error rule, which fires before this body runs). Each argument
// is coerced via `coerce_to_bool`; the first coercion failure surfaces as
// #VALUE! (or #NUM! for non-finite numeric inputs). AND returns true iff
// every argument coerces to true; OR returns true iff any does.
Value And_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  bool result = true;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_bool(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (!coerced.value()) {
      result = false;
    }
  }
  return Value::boolean(result);
}

Value Or_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  bool result = false;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_bool(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value()) {
      result = true;
    }
  }
  return Value::boolean(result);
}

// LEN(text) --------------------------------------------------------------
// Excel reports length in UTF-16 code units, which differs from byte length
// for any non-ASCII codepoint. We coerce the argument to text, then count
// units via the standalone helper.
Value Len(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_text(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  return Value::number(static_cast<double>(utf16_units_in(coerced.value())));
}

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

// --- Aggregates ---------------------------------------------------------

// MIN(value, ...) - smallest of the coerced numbers. The Excel "skip text
// in cell-references" rule does NOT apply here: a literal non-numeric
// argument coerces via `coerce_to_number` and surfaces `#VALUE!` on
// failure. The caller's pre-evaluation has already short-circuited any
// argument that was itself an error.
Value Min(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  // arity >= 1 by registry contract (min_arity = 1).
  auto first = coerce_to_number(args[0]);
  if (!first) {
    return Value::error(first.error());
  }
  double best = first.value();
  for (std::uint32_t i = 1; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value() < best) {
      best = coerced.value();
    }
  }
  if (std::isnan(best) || std::isinf(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

// MAX(value, ...) - symmetric to MIN.
Value Max(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto first = coerce_to_number(args[0]);
  if (!first) {
    return Value::error(first.error());
  }
  double best = first.value();
  for (std::uint32_t i = 1; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value() > best) {
      best = coerced.value();
    }
  }
  if (std::isnan(best) || std::isinf(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

// AVERAGE(value, ...) - arithmetic mean. With `min_arity = 1` enforced
// at the registry, the divisor is always at least 1, so there is no
// divide-by-zero edge case.
Value Average(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    total += coerced.value();
  }
  const double r = total / static_cast<double>(arity);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// PRODUCT(value, ...) - product of all args. Overflow to Inf -> `#NUM!`.
Value Product(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 1.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    total *= coerced.value();
  }
  if (std::isnan(total) || std::isinf(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

// --- Counting aggregators -----------------------------------------------
//
// COUNT / COUNTA / COUNTBLANK. Unlike SUM / AVERAGE / MIN / MAX / PRODUCT
// these three are registered with `propagate_errors = false`: Excel's COUNT
// family is specified in terms of which cells to "count" rather than which
// values to coerce, so an error inside a range must not short-circuit the
// whole call. That opt-out means the impls see Error-typed values in their
// args array directly and must skip them explicitly.
//
// Accepted divergence (range-vs-direct parity):
//   =COUNT(1, 1/0) in Excel is #DIV/0! because the error propagates as a
//   direct argument; =COUNT(A1:A2) where one cell holds #DIV/0! silently
//   skips the error and counts only the numerics. We do not distinguish
//   the two call shapes - any error anywhere is skipped - so direct-arg
//   callers get the range-shape behaviour. The same simplification already
//   applies to text / bool values appearing inside SUM-family ranges (see
//   FunctionDef::accepts_ranges comment in function_registry.h). A true
//   range-context concept is deferred.

// COUNT(value, ...) - count of Number values. Booleans, text (even
// numeric-looking text like "5"), blanks, and errors are all skipped.
// Direct-arg booleans are also skipped (Excel's COUNT never counts
// booleans even when they appear outside a range).
Value Count(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (args[i].is_number()) {
      total += 1.0;
    }
  }
  return Value::number(total);
}

// COUNTA(value, ...) - count of non-Blank values. Numbers, booleans, text
// (including the empty string produced by a formula returning ""), and
// errors are all counted. Only the Blank scalar is skipped.
Value CountA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (!args[i].is_blank()) {
      total += 1.0;
    }
  }
  return Value::number(total);
}

// COUNTBLANK(value, ...) - count of Blank scalars and Text values whose
// contents are exactly "". Numbers (including 0), booleans (including
// FALSE), non-empty text, and errors are all skipped. The public Excel 365
// signature accepts a single range; we accept variadic for symmetry with
// the sibling aggregators - a single A1:B2 ref still expands to many
// scalar args via the dispatcher.
Value CountBlank(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    const Value& v = args[i];
    if (v.is_blank() || (v.is_text() && v.as_text().empty())) {
      total += 1.0;
    }
  }
  return Value::number(total);
}

// --- Exponential / logarithmic / trigonometric --------------------------
//
// Every function in this section coerces its inputs through
// `coerce_to_number` (Bool -> 0/1, Text -> parsed, Blank -> 0, Error ->
// propagated by the dispatcher's left-most-error rule) and returns
// `#NUM!` for any non-finite result. Trigonometric inputs are radians.

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

// --- Text ---------------------------------------------------------------
//
// Every text builtin coerces its inputs via `coerce_to_text` /
// `coerce_to_number`. Errors among the inputs already short-circuit through
// the dispatcher's left-most-error rule before we get here. Length and
// position arithmetic uses Excel's UTF-16 unit semantics via
// `eval/text_ops.h`; the result text (when any) is interned into the
// caller's arena so the returned `Value::text` payload is readable.

// Excel caps the result of REPT (and a handful of related text functions)
// at 32,767 UTF-16 units. We reuse the same constant in REPT's overflow
// guard.
constexpr std::uint64_t kExcelTextCapUnits = 32767u;

// Helper: read a numeric arg as an `int` via `std::trunc`. Returns `#VALUE!`
// on coercion failure, `#NUM!` on non-finite input. Used by LEFT/RIGHT/MID/
// REPT/FIND/SEARCH/SUBSTITUTE for their integer-typed parameters.
Expected<int, ErrorCode> read_int_arg(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return static_cast<int>(std::trunc(d));
}

// UPPER(text) / LOWER(text) - ASCII case fold. Multi-byte UTF-8 bytes are
// preserved verbatim (see `text_ops::to_upper_ascii` for the contract).
Value Upper(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  return Value::text(arena.intern(to_upper_ascii(text.value())));
}

Value Lower(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  return Value::text(arena.intern(to_lower_ascii(text.value())));
}

// TRIM(text) - removes leading and trailing ASCII spaces (0x20) and
// collapses runs of internal ASCII spaces to a single space. Other
// whitespace-like bytes (tabs, newlines, Unicode whitespace) are preserved
// verbatim - this matches Excel exactly.
Value Trim(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  std::string out;
  out.reserve(src.size());
  bool pending_space = false;
  bool seen_non_space = false;
  for (char c : src) {
    if (c == ' ') {
      if (seen_non_space) {
        pending_space = true;
      }
      continue;
    }
    if (pending_space) {
      out.push_back(' ');
      pending_space = false;
    }
    out.push_back(c);
    seen_non_space = true;
  }
  return Value::text(arena.intern(out));
}

// LEFT(text, [n]) - first `n` UTF-16 units. Default n=1. n<0 -> `#VALUE!`.
Value Left(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  int n = 1;
  if (arity >= 2) {
    auto parsed = read_int_arg(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    n = parsed.value();
  }
  if (n < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (n == 0) {
    return Value::text({});
  }
  return Value::text(arena.intern(utf16_substring(text.value(), 0u, static_cast<std::uint32_t>(n))));
}

// RIGHT(text, [n]) - last `n` UTF-16 units. Default n=1. n<0 -> `#VALUE!`.
Value Right(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  int n = 1;
  if (arity >= 2) {
    auto parsed = read_int_arg(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    n = parsed.value();
  }
  if (n < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (n == 0) {
    return Value::text({});
  }
  const std::uint32_t total = utf16_units_in(text.value());
  const auto take = static_cast<std::uint32_t>(n);
  const std::uint32_t start = take >= total ? 0u : total - take;
  return Value::text(arena.intern(utf16_substring(text.value(), start, take)));
}

// MID(text, start_num, num_chars) - 1-based slice in UTF-16 units. Excel
// returns `""` when `start_num` is past the end. `start_num<1` or
// `num_chars<0` -> `#VALUE!`.
Value Mid(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto start = read_int_arg(args[1]);
  if (!start) {
    return Value::error(start.error());
  }
  auto length = read_int_arg(args[2]);
  if (!length) {
    return Value::error(length.error());
  }
  if (start.value() < 1 || length.value() < 0) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t total = utf16_units_in(text.value());
  const auto start_unit = static_cast<std::uint32_t>(start.value() - 1);
  if (start_unit >= total) {
    return Value::text({});
  }
  if (length.value() == 0) {
    return Value::text({});
  }
  return Value::text(
      arena.intern(utf16_substring(text.value(), start_unit, static_cast<std::uint32_t>(length.value()))));
}

// REPT(text, n) - repeat. n<0 -> `#VALUE!`. Excel caps the result length at
// 32,767 UTF-16 units; exceeding the cap also surfaces as `#VALUE!`.
Value Rept(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto count = read_int_arg(args[1]);
  if (!count) {
    return Value::error(count.error());
  }
  if (count.value() < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (count.value() == 0 || text.value().empty()) {
    return Value::text({});
  }
  const auto unit_len = static_cast<std::uint64_t>(utf16_units_in(text.value()));
  const auto reps = static_cast<std::uint64_t>(count.value());
  if (unit_len > 0 && reps > kExcelTextCapUnits / unit_len) {
    return Value::error(ErrorCode::Value);
  }
  std::string out;
  out.reserve(text.value().size() * reps);
  for (std::uint64_t i = 0; i < reps; ++i) {
    out.append(text.value());
  }
  return Value::text(arena.intern(out));
}

// SUBSTITUTE(text, old_text, new_text, [instance_num]) - case-sensitive,
// byte-exact replace. Without `instance_num`, every occurrence is replaced.
// With `instance_num`, only the Nth (1-based) occurrence. Empty `old_text`
// and `instance_num` greater than the number of occurrences both return
// `text` unchanged. `instance_num < 1` -> `#VALUE!`.
Value Substitute(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto old_text = coerce_to_text(args[1]);
  if (!old_text) {
    return Value::error(old_text.error());
  }
  auto new_text = coerce_to_text(args[2]);
  if (!new_text) {
    return Value::error(new_text.error());
  }
  bool nth_only = false;
  int instance = 0;
  if (arity >= 4) {
    auto parsed = read_int_arg(args[3]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    if (parsed.value() < 1) {
      return Value::error(ErrorCode::Value);
    }
    nth_only = true;
    instance = parsed.value();
  }
  const std::string& haystack = text.value();
  const std::string& needle = old_text.value();
  if (needle.empty()) {
    return Value::text(arena.intern(haystack));
  }
  std::string out;
  out.reserve(haystack.size());
  std::size_t i = 0;
  int hits = 0;
  while (i < haystack.size()) {
    const std::size_t pos = haystack.find(needle, i);
    if (pos == std::string::npos) {
      out.append(haystack, i, std::string::npos);
      break;
    }
    out.append(haystack, i, pos - i);
    ++hits;
    if (!nth_only || hits == instance) {
      out.append(new_text.value());
    } else {
      out.append(needle);
    }
    i = pos + needle.size();
  }
  return Value::text(arena.intern(out));
}

// FIND(find_text, within_text, [start_num]) - case-sensitive, no wildcards.
// 1-based UTF-16-unit position of the first occurrence at or after
// `start_num` (default 1). Not found -> `#VALUE!`. Out-of-range `start_num`
// -> `#VALUE!`. Empty `find_text` returns `start_num` (Excel quirk).
Value Find(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto needle = coerce_to_text(args[0]);
  if (!needle) {
    return Value::error(needle.error());
  }
  auto haystack = coerce_to_text(args[1]);
  if (!haystack) {
    return Value::error(haystack.error());
  }
  int start = 1;
  if (arity >= 3) {
    auto parsed = read_int_arg(args[2]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    start = parsed.value();
  }
  const std::uint32_t total = utf16_units_in(haystack.value());
  if (start < 1 || static_cast<std::uint32_t>(start) > total + 1) {
    return Value::error(ErrorCode::Value);
  }
  if (needle.value().empty()) {
    return Value::number(static_cast<double>(start));
  }
  const std::size_t start_byte = utf16_to_byte_offset(haystack.value(), static_cast<std::uint32_t>(start - 1));
  const std::size_t pos = haystack.value().find(needle.value(), start_byte);
  if (pos == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  // Convert byte offset back to a 1-based UTF-16 unit position.
  const std::uint32_t units = utf16_units_in(std::string_view(haystack.value()).substr(0, pos));
  return Value::number(static_cast<double>(units + 1));
}

// SEARCH(find_text, within_text, [start_num]) - case-insensitive, no
// wildcards in this MVP. Otherwise mirrors FIND. Wildcard support (`*`,
// `?`, `~*`, `~?`) is a known limitation; with it, `SEARCH("a*c", "abc")`
// would match. Today the call returns `#VALUE!` because the literal
// substring `"a*c"` is not present.
Value Search(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto needle = coerce_to_text(args[0]);
  if (!needle) {
    return Value::error(needle.error());
  }
  auto haystack = coerce_to_text(args[1]);
  if (!haystack) {
    return Value::error(haystack.error());
  }
  int start = 1;
  if (arity >= 3) {
    auto parsed = read_int_arg(args[2]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    start = parsed.value();
  }
  const std::uint32_t total = utf16_units_in(haystack.value());
  if (start < 1 || static_cast<std::uint32_t>(start) > total + 1) {
    return Value::error(ErrorCode::Value);
  }
  if (needle.value().empty()) {
    return Value::number(static_cast<double>(start));
  }
  const std::string lowered_haystack = to_lower_ascii(haystack.value());
  const std::string lowered_needle = to_lower_ascii(needle.value());
  const std::size_t start_byte = utf16_to_byte_offset(haystack.value(), static_cast<std::uint32_t>(start - 1));
  const std::size_t pos = lowered_haystack.find(lowered_needle, start_byte);
  if (pos == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t units = utf16_units_in(std::string_view(haystack.value()).substr(0, pos));
  return Value::number(static_cast<double>(units + 1));
}

// VALUE(text) - parse text as a number. Numeric inputs round-trip; bools
// are rejected (Excel's VALUE deliberately disallows boolean inputs even
// though they otherwise coerce to 1/0 in arithmetic). Text parse failure
// -> `#VALUE!`. Errors propagate.
Value Value_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  switch (v.kind()) {
    case ValueKind::Number:
      return v;
    case ValueKind::Bool:
      return Value::error(ErrorCode::Value);
    case ValueKind::Error:
      return v;
    case ValueKind::Blank:
    case ValueKind::Text: {
      auto text = coerce_to_text(v);
      if (!text) {
        return Value::error(text.error());
      }
      const Value as_text = Value::text(text.value());
      auto num = coerce_to_number(as_text);
      if (!num) {
        return Value::error(num.error());
      }
      return Value::number(num.value());
    }
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

// EXACT(text1, text2) - byte-wise (case-sensitive) equality.
Value Exact(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto a = coerce_to_text(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  auto b = coerce_to_text(args[1]);
  if (!b) {
    return Value::error(b.error());
  }
  return Value::boolean(a.value() == b.value());
}

// --- Text manipulation, second batch ------------------------------------
//
// TEXTJOIN, UNICHAR, UNICODE, CLEAN, PROPER. The same conventions as the
// first text batch apply: argument coercion via the helpers in
// `eval/coerce.h`, error propagation through the dispatcher's left-most
// rule, results interned into the call's arena.

// TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
//
// Joins every text argument with `delimiter`. When `ignore_empty` is TRUE
// (the typical case), arguments whose text representation is the empty
// string are skipped, so two consecutive empty inputs do NOT produce a
// double delimiter. With `ignore_empty` FALSE every argument participates
// even if empty (yielding consecutive delimiters). Result length is capped
// at Excel's 32,767-unit limit; exceeding it surfaces `#VALUE!`.
Value TextJoin(const Value* args, std::uint32_t arity, Arena& arena) {
  auto delimiter = coerce_to_text(args[0]);
  if (!delimiter) {
    return Value::error(delimiter.error());
  }
  auto ignore_empty = coerce_to_bool(args[1]);
  if (!ignore_empty) {
    return Value::error(ignore_empty.error());
  }
  std::string out;
  bool first = true;
  for (std::uint32_t i = 2; i < arity; ++i) {
    auto piece = coerce_to_text(args[i]);
    if (!piece) {
      return Value::error(piece.error());
    }
    if (ignore_empty.value() && piece.value().empty()) {
      continue;
    }
    if (!first) {
      out.append(delimiter.value());
    }
    out.append(piece.value());
    first = false;
    // Early-out cap check: once the byte length exceeds the byte upper bound
    // for the cap (4 bytes per UTF-16 unit pessimistically), the UTF-16 unit
    // count must also exceed the cap. Definitive check below.
    if (utf16_units_in(out) > kExcelTextCapUnits) {
      return Value::error(ErrorCode::Value);
    }
  }
  return Value::text(arena.intern(out));
}

// UNICHAR(number) - returns the Unicode character whose codepoint is
// `number`. Truncates the input to an integer. Out-of-range and surrogate
// codepoints surface `#VALUE!`. Result is encoded as UTF-8 bytes.
Value Unichar(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto parsed = read_int_arg(args[0]);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const int n = parsed.value();
  if (n < 1 || n > 0x10FFFF) {
    return Value::error(ErrorCode::Value);
  }
  if (n >= 0xD800 && n <= 0xDFFF) {
    // UTF-16 surrogate halves do not represent characters on their own.
    return Value::error(ErrorCode::Value);
  }
  const std::string encoded = encode_utf8_codepoint(static_cast<std::uint32_t>(n));
  if (encoded.empty()) {
    // Defensive: encoder validates internally; an empty string here would
    // mean the helper rejected our codepoint despite the range checks.
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(encoded));
}

// UNICODE(text) - returns the Unicode codepoint of the first character in
// `text`. Empty text yields `#VALUE!`. The returned value is the actual
// codepoint, not a UTF-16 code unit: supplementary-plane characters return
// values above 0xFFFF (e.g. `UNICODE("😀")` = 128512, not the high surrogate).
Value Unicode_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  if (text.value().empty()) {
    return Value::error(ErrorCode::Value);
  }
  const Utf8DecodeResult decoded = decode_first_utf8_codepoint(text.value());
  if (!decoded.valid) {
    return Value::error(ErrorCode::Value);
  }
  return Value::number(static_cast<double>(decoded.codepoint));
}

// CLEAN(text) - strips ASCII control characters (0x00..0x1F) from `text`.
// Bytes >= 0x20 (including 0x7F DEL and the entire UTF-8 multi-byte range
// 0x80..0xFF) are preserved verbatim. Embedded NUL is NOT a string
// terminator here: the input is a `string_view` and we copy through every
// non-control byte.
Value Clean(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  std::string out;
  out.reserve(src.size());
  for (char c : src) {
    if (static_cast<unsigned char>(c) >= 0x20u) {
      out.push_back(c);
    }
  }
  return Value::text(arena.intern(out));
}

// PROPER(text) - title-case `text`. ASCII letters that begin a "word" are
// uppercased; ASCII letters that follow another ASCII letter are lowercased.
// A "word boundary" is any byte that is NOT an ASCII letter, including
// digits, punctuation, whitespace, and any byte >= 0x80 (so a Japanese
// character followed by an ASCII letter starts a new word). Non-ASCII bytes
// pass through unchanged - matching the existing UPPER / LOWER policy.
Value Proper(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  std::string out;
  out.reserve(src.size());
  bool start_of_word = true;
  for (char c : src) {
    const auto u = static_cast<unsigned char>(c);
    const bool is_lower = (u >= 'a' && u <= 'z');
    const bool is_upper = (u >= 'A' && u <= 'Z');
    if (is_lower || is_upper) {
      if (start_of_word) {
        out.push_back(is_upper ? c : static_cast<char>(c - 32));
      } else {
        out.push_back(is_lower ? c : static_cast<char>(c + 32));
      }
      start_of_word = false;
    } else {
      out.push_back(c);
      start_of_word = true;
    }
  }
  return Value::text(arena.intern(out));
}

// --- Type predicates ----------------------------------------------------
//
// These functions are registered with `propagate_errors = false` so the
// dispatcher hands them error-typed inputs verbatim. Each returns a Bool
// based purely on the input's `ValueKind` (and, for ISERR / ISNA, the
// specific `ErrorCode`). None of them ever surface a formula error of
// their own beyond the registry's arity check.

// ISNUMBER(value) - true iff the input is a Number cell.
Value IsNumber(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Number);
}

// ISTEXT(value) - true iff the input is a Text cell.
Value IsText(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Text);
}

// ISBLANK(value) - true iff the input is the Blank scalar. Empty text
// (`""`) is NOT blank in Excel - `ISBLANK("")` returns FALSE.
Value IsBlank(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Blank);
}

// ISLOGICAL(value) - true iff the input is a Bool. Numeric 0/1 do not
// count: only the actual TRUE/FALSE booleans qualify.
Value IsLogical(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Bool);
}

// ISERROR(value) - true iff the input is any formula error, including
// `#N/A`. Combined with the cleared `propagate_errors` flag this lets
// callers branch on errors without first wrapping them in IFERROR.
Value IsError(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Error);
}

// ISERR(value) - true iff the input is a formula error OTHER than `#N/A`.
// `ISERR(#N/A)` is FALSE; `ISERR(#DIV/0!)`, `ISERR(#REF!)`, etc. are TRUE.
Value IsErr(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() != ValueKind::Error) {
    return Value::boolean(false);
  }
  return Value::boolean(v.as_error() != ErrorCode::NA);
}

// ISNA(value) - true iff the input is exactly `#N/A`. All other errors
// (and all non-error values) yield FALSE.
Value IsNa(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() != ValueKind::Error) {
    return Value::boolean(false);
  }
  return Value::boolean(v.as_error() == ErrorCode::NA);
}

// --- Coercion-style info functions --------------------------------------
//
// `N` and `T` propagate errors via the dispatcher's default short-circuit
// (no flag override needed). They differ from `VALUE` / generic
// `coerce_to_*` in that they NEVER fail: any non-matching input maps to
// the function's neutral element (0 for `N`, "" for `T`).

// N(value) - coerce to a Number with Excel's narrow rules:
//   - Number          -> the number unchanged
//   - Bool            -> 1.0 for TRUE, 0.0 for FALSE
//   - Text            -> 0.0 ALWAYS (N intentionally does not parse text;
//                        contrast with VALUE, which does)
//   - Blank           -> 0.0
//   - Error           -> propagated by the dispatcher before this body runs
Value N(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  switch (v.kind()) {
    case ValueKind::Number:
      return v;
    case ValueKind::Bool:
      return Value::number(v.as_boolean() ? 1.0 : 0.0);
    case ValueKind::Text:
    case ValueKind::Blank:
      return Value::number(0.0);
    case ValueKind::Error:
      // Unreachable in practice: dispatcher short-circuits errors. Defensive
      // fall-through keeps the switch exhaustive.
      return v;
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

// T(value) - coerce to Text with Excel's narrow rules:
//   - Text            -> the text unchanged
//   - Number / Bool   -> empty string ""
//   - Blank           -> empty string ""
//   - Error           -> propagated by the dispatcher before this body runs
Value T(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() == ValueKind::Text) {
    return v;
  }
  if (v.kind() == ValueKind::Error) {
    // Unreachable in practice: dispatcher short-circuits errors.
    return v;
  }
  return Value::text({});
}

// --- Statistical aggregators --------------------------------------------
//
// MEDIAN, MODE / MODE.SNGL, LARGE / SMALL, PERCENTILE[.INC],
// QUARTILE[.INC], STDEV[.S], STDEV.P, VAR[.S], VAR.P.
//
// Argument-type rule (DIFFERENT from SUM / AVERAGE / MIN / MAX / PRODUCT):
// these functions silently SKIP text, boolean, and blank inputs instead of
// coercing them. Only values whose kind is `Number` participate. The
// dispatcher runs with `propagate_errors = true` so error-typed arguments
// still short-circuit before the impl executes; once inside the body we
// only need to filter the non-numeric non-error kinds. Contrast this with
// `Sum` at the top of this file, which coerces every argument through
// `coerce_to_number` and surfaces `#VALUE!` on text like `"abc"`.
//
// For `LARGE`, `SMALL`, `PERCENTILE.INC`, and `QUARTILE.INC` the dispatcher
// lays out arguments in source order: any leading range expands into
// scalar cells, and the trailing scalar (k or quart) lives at
// `args[arity - 1]`. The data slice is therefore `args[0 .. arity - 2]`.
// That trimming is done explicitly at each callsite (not in the helper)
// so the slice boundary stays visible.

// Extracts the numeric values from `args[0..count-1]`. Non-Number values
// (text / bool / blank after range expansion) are silently skipped. Errors
// never reach this helper because the dispatcher short-circuits with
// `propagate_errors = true`.
std::vector<double> collect_numerics(const Value* args, std::uint32_t count) {
  std::vector<double> out;
  out.reserve(count);
  for (std::uint32_t i = 0; i < count; ++i) {
    if (args[i].is_number()) {
      out.push_back(args[i].as_number());
    }
  }
  return out;
}

// MEDIAN(value, ...) - median of numeric values. Non-numerics are skipped;
// an empty collection yields `#NUM!`. For an even count the result is the
// average of the two central values.
Value Median(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const std::size_t n = xs.size();
  const double r = (n % 2u == 1u) ? xs[n / 2u] : 0.5 * (xs[(n / 2u) - 1u] + xs[n / 2u]);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// MODE(value, ...) / MODE.SNGL(value, ...) - most frequent numeric value.
// Ties are broken by first-occurrence order. An empty collection or a
// collection whose values are all unique (no value appears >= 2 times)
// yields `#N/A`. Implementation: stable-sort a (value, original-index)
// array, then sweep once to find the longest run; ties are broken by the
// smallest first-occurrence index of tied runs.
Value Mode(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < 2u) {
    return Value::error(ErrorCode::NA);
  }
  // Pair (value, original-index). Sorting by value lets us find runs, and
  // we remember the smallest original index within each run for tie-break.
  std::vector<std::pair<double, std::size_t>> pairs;
  pairs.reserve(xs.size());
  for (std::size_t i = 0; i < xs.size(); ++i) {
    pairs.emplace_back(xs[i], i);
  }
  std::sort(pairs.begin(), pairs.end(),
            [](const std::pair<double, std::size_t>& a, const std::pair<double, std::size_t>& b) {
              return a.first < b.first;
            });
  std::size_t best_run_len = 1;
  std::size_t best_first_idx = pairs[0].second;
  double best_value = pairs[0].first;
  bool best_valid = false;  // Needs run length >= 2 to qualify.
  std::size_t i = 0;
  while (i < pairs.size()) {
    std::size_t j = i;
    std::size_t first_idx = pairs[i].second;
    while (j < pairs.size() && pairs[j].first == pairs[i].first) {
      if (pairs[j].second < first_idx) {
        first_idx = pairs[j].second;
      }
      ++j;
    }
    const std::size_t run_len = j - i;
    if (run_len >= 2u) {
      // Strictly longer wins; on ties the earlier first-occurrence wins.
      if (!best_valid || run_len > best_run_len || (run_len == best_run_len && first_idx < best_first_idx)) {
        best_run_len = run_len;
        best_first_idx = first_idx;
        best_value = pairs[i].first;
        best_valid = true;
      }
    }
    i = j;
  }
  if (!best_valid) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(best_value);
}

// Helper: read the trailing scalar `k` argument for LARGE / SMALL /
// PERCENTILE / QUARTILE. Returns the raw coerced double so each caller
// applies its own range / truncation rules.
Expected<double, ErrorCode> read_kth_arg(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return d;
}

// LARGE(array, k) - k-th largest numeric. k is truncated toward zero; any
// fractional part is discarded. k must be in [1, numeric_count], else
// `#NUM!`. An empty numeric slice trivially fails the upper bound.
Value Large(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto k_raw = read_kth_arg(args[arity - 1u]);
  if (!k_raw) {
    return Value::error(k_raw.error());
  }
  const double k_floor = std::floor(k_raw.value());
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty() || k_floor < 1.0 || k_floor > static_cast<double>(xs.size())) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const auto k = static_cast<std::size_t>(k_floor);
  return Value::number(xs[xs.size() - k]);
}

// SMALL(array, k) - k-th smallest numeric. Same rules as LARGE.
Value Small(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto k_raw = read_kth_arg(args[arity - 1u]);
  if (!k_raw) {
    return Value::error(k_raw.error());
  }
  const double k_floor = std::floor(k_raw.value());
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty() || k_floor < 1.0 || k_floor > static_cast<double>(xs.size())) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const auto k = static_cast<std::size_t>(k_floor);
  return Value::number(xs[k - 1u]);
}

// PERCENTILE.INC(array, k) / PERCENTILE(array, k) - linear-interpolation
// percentile. k is the fractional rank in [0, 1]; out-of-range yields
// `#NUM!`. Empty numeric slice yields `#NUM!`. The interpolation point is
// `pos = k * (n - 1)` (0-based); fractional `pos` blends the two neighbours.
Value PercentileInc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto k_raw = read_kth_arg(args[arity - 1u]);
  if (!k_raw) {
    return Value::error(k_raw.error());
  }
  const double k = k_raw.value();
  if (k < 0.0 || k > 1.0) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const double pos = k * static_cast<double>(xs.size() - 1u);
  const double floor_pos = std::floor(pos);
  const auto idx = static_cast<std::size_t>(floor_pos);
  const double frac = pos - floor_pos;
  double r;
  if (frac == 0.0 || idx + 1u >= xs.size()) {
    r = xs[idx];
  } else {
    r = xs[idx] + frac * (xs[idx + 1u] - xs[idx]);
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// QUARTILE.INC(array, quart) / QUARTILE(array, quart) - quartile by
// `PERCENTILE.INC(array, quart/4)`. `quart` must be an integer in
// {0, 1, 2, 3, 4}; non-integer or out-of-range yields `#NUM!`.
Value QuartileInc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const std::uint32_t data_count = arity - 1u;
  auto q_raw = read_kth_arg(args[arity - 1u]);
  if (!q_raw) {
    return Value::error(q_raw.error());
  }
  const double q = q_raw.value();
  // Excel truncates QUARTILE's quart toward zero and rejects values
  // outside [0, 4]. A fractional quart (e.g. 1.5) is rejected.
  if (q < 0.0 || q > 4.0 || std::floor(q) != q) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> xs = collect_numerics(args, data_count);
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  std::sort(xs.begin(), xs.end());
  const double k = q / 4.0;
  const double pos = k * static_cast<double>(xs.size() - 1u);
  const double floor_pos = std::floor(pos);
  const auto idx = static_cast<std::size_t>(floor_pos);
  const double frac = pos - floor_pos;
  double r;
  if (frac == 0.0 || idx + 1u >= xs.size()) {
    r = xs[idx];
  } else {
    r = xs[idx] + frac * (xs[idx + 1u] - xs[idx]);
  }
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Helper: compute `(mean, sum_of_squared_deviations)` over a numeric slice.
// Empty input returns `{0, 0}` which the callers treat as a DIV/0! case.
struct MeanSS {
  double mean;
  double ss;  // Sum of squared deviations from the mean.
};
MeanSS compute_mean_ss(const std::vector<double>& xs) {
  if (xs.empty()) {
    return {0.0, 0.0};
  }
  double sum = 0.0;
  for (double x : xs) {
    sum += x;
  }
  const double mean = sum / static_cast<double>(xs.size());
  double ss = 0.0;
  for (double x : xs) {
    const double d = x - mean;
    ss += d * d;
  }
  return {mean, ss};
}

// VAR.S(value, ...) / VAR(value, ...) - sample variance with divisor n - 1.
// Fewer than 2 numeric inputs yields `#DIV/0!`.
Value VarS(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < 2u) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double r = ms.ss / static_cast<double>(xs.size() - 1u);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// VAR.P(value, ...) - population variance with divisor n. A single numeric
// input yields 0; no numeric inputs yields `#DIV/0!`.
Value VarP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double r = ms.ss / static_cast<double>(xs.size());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// STDEV.S(value, ...) / STDEV(value, ...) - sample standard deviation,
// `sqrt(VAR.S)`. Fewer than 2 numeric inputs yields `#DIV/0!`.
Value StdevS(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < 2u) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double var = ms.ss / static_cast<double>(xs.size() - 1u);
  const double r = std::sqrt(var);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// STDEV.P(value, ...) - population standard deviation, `sqrt(VAR.P)`.
Value StdevP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double var = ms.ss / static_cast<double>(xs.size());
  const double r = std::sqrt(var);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace

void register_builtins(FunctionRegistry& registry) {
  {
    // SUM is range-aware: `=SUM(A1:A100)` expands the rectangle into scalar
    // cell values before this impl runs.
    FunctionDef def{"SUM", 1u, kVariadic, &Sum};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  registry.register_function(FunctionDef{"CONCAT", 1u, kVariadic, &Concat});
  // CONCATENATE is the legacy spelling kept by Excel for compatibility; it
  // shares the implementation with CONCAT.
  registry.register_function(FunctionDef{"CONCATENATE", 1u, kVariadic, &Concat});
  registry.register_function(FunctionDef{"LEN", 1u, 1u, &Len});
  registry.register_function(FunctionDef{"TRUE", 0u, 0u, &True_});
  registry.register_function(FunctionDef{"FALSE", 0u, 0u, &False_});
  registry.register_function(FunctionDef{"NOT", 1u, 1u, &Not});
  registry.register_function(FunctionDef{"AND", 1u, kVariadic, &And_});
  registry.register_function(FunctionDef{"OR", 1u, kVariadic, &Or_});

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

  // Aggregates (min_arity = 1, variadic). Each is range-aware: a RangeOp
  // argument is flattened into scalar cell values by the dispatcher before
  // the impl runs.
  {
    FunctionDef def{"MIN", 1u, kVariadic, &Min};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"MAX", 1u, kVariadic, &Max};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"AVERAGE", 1u, kVariadic, &Average};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"PRODUCT", 1u, kVariadic, &Product};
    def.accepts_ranges = true;
    registry.register_function(def);
  }

  // Counting aggregators. All three are range-aware and opt out of the
  // dispatcher's left-most-error rule so the impl itself decides which
  // values to count (see the block comment above `Count` in builtins.cpp
  // for the range-vs-direct divergence).
  {
    FunctionDef def{"COUNT", 1u, kVariadic, &Count, /*propagate_errors=*/false};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"COUNTA", 1u, kVariadic, &CountA, /*propagate_errors=*/false};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"COUNTBLANK", 1u, kVariadic, &CountBlank, /*propagate_errors=*/false};
    def.accepts_ranges = true;
    registry.register_function(def);
  }

  // Statistical aggregators. Every entry below is range-aware and keeps the
  // default `propagate_errors = true`: errors short-circuit before the impl
  // runs, and the impls filter non-numeric kinds (text / bool / blank)
  // themselves -- see the block comment above `Median` in this file.
  {
    FunctionDef def{"MEDIAN", 1u, kVariadic, &Median};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"MODE", 1u, kVariadic, &Mode};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // MODE.SNGL is Excel 2010+'s canonical spelling; implementation is
    // identical to MODE. Registry already handles the dotted name.
    FunctionDef def{"MODE.SNGL", 1u, kVariadic, &Mode};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // LARGE(array, k) - trailing scalar k lives at `args[arity - 1]` after
    // the dispatcher has expanded any leading range; the impl trims the
    // slice explicitly before collecting numerics.
    FunctionDef def{"LARGE", 2u, kVariadic, &Large};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"SMALL", 2u, kVariadic, &Small};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"PERCENTILE.INC", 2u, kVariadic, &PercentileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    // PERCENTILE is the pre-2010 spelling of PERCENTILE.INC; same impl.
    FunctionDef def{"PERCENTILE", 2u, kVariadic, &PercentileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"QUARTILE.INC", 2u, kVariadic, &QuartileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"QUARTILE", 2u, kVariadic, &QuartileInc};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV.S", 1u, kVariadic, &StdevS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV", 1u, kVariadic, &StdevS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"STDEV.P", 1u, kVariadic, &StdevP};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR.S", 1u, kVariadic, &VarS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR", 1u, kVariadic, &VarS};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"VAR.P", 1u, kVariadic, &VarP};
    def.accepts_ranges = true;
    registry.register_function(def);
  }

  // Exponential / logarithmic / trigonometric.
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

  // Text manipulation.
  registry.register_function(FunctionDef{"UPPER", 1u, 1u, &Upper});
  registry.register_function(FunctionDef{"LOWER", 1u, 1u, &Lower});
  registry.register_function(FunctionDef{"TRIM", 1u, 1u, &Trim});
  registry.register_function(FunctionDef{"LEFT", 1u, 2u, &Left});
  registry.register_function(FunctionDef{"RIGHT", 1u, 2u, &Right});
  registry.register_function(FunctionDef{"MID", 3u, 3u, &Mid});
  registry.register_function(FunctionDef{"REPT", 2u, 2u, &Rept});
  registry.register_function(FunctionDef{"SUBSTITUTE", 3u, 4u, &Substitute});
  registry.register_function(FunctionDef{"FIND", 2u, 3u, &Find});
  registry.register_function(FunctionDef{"SEARCH", 2u, 3u, &Search});
  registry.register_function(FunctionDef{"VALUE", 1u, 1u, &Value_});
  registry.register_function(FunctionDef{"EXACT", 2u, 2u, &Exact});

  // Text manipulation, second batch.
  registry.register_function(FunctionDef{"TEXTJOIN", 3u, kVariadic, &TextJoin});
  registry.register_function(FunctionDef{"UNICHAR", 1u, 1u, &Unichar});
  registry.register_function(FunctionDef{"UNICODE", 1u, 1u, &Unicode_});
  registry.register_function(FunctionDef{"CLEAN", 1u, 1u, &Clean});
  registry.register_function(FunctionDef{"PROPER", 1u, 1u, &Proper});

  // Info / type queries.
  // The IS* family must inspect error-typed arguments verbatim, so each
  // entry clears `propagate_errors` to opt out of the dispatcher's default
  // left-most-error short-circuit. `N` and `T` use the default (errors
  // propagate before the body runs).
  registry.register_function(FunctionDef{"ISNUMBER", 1u, 1u, &IsNumber, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISTEXT", 1u, 1u, &IsText, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISBLANK", 1u, 1u, &IsBlank, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISLOGICAL", 1u, 1u, &IsLogical, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISERROR", 1u, 1u, &IsError, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISERR", 1u, 1u, &IsErr, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISNA", 1u, 1u, &IsNa, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"N", 1u, 1u, &N});
  registry.register_function(FunctionDef{"T", 1u, 1u, &T});
}

}  // namespace eval
}  // namespace formulon
