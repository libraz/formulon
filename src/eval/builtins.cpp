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

#include <cmath>
#include <cstdint>
#include <string>

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

}  // namespace

void register_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"SUM", 1u, kVariadic, &Sum});
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

  // Aggregates (min_arity = 1, variadic).
  registry.register_function(FunctionDef{"MIN", 1u, kVariadic, &Min});
  registry.register_function(FunctionDef{"MAX", 1u, kVariadic, &Max});
  registry.register_function(FunctionDef{"AVERAGE", 1u, kVariadic, &Average});
  registry.register_function(FunctionDef{"PRODUCT", 1u, kVariadic, &Product});

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
}

}  // namespace eval
}  // namespace formulon
