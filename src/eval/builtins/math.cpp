//
// Implementation of Formulon's arithmetic and rounding built-in functions:
// ABS, SIGN, INT, TRUNC, SQRT, MOD, POWER, ROUND, ROUNDDOWN, ROUNDUP,
// EVEN, ODD, and QUOTIENT. Each impl follows the same recipe as the rest
// of the builtin catalog: coerce arguments via `eval/coerce.h`, propagate
// the left-most coercion error, and return a `Value`.

#include "eval/builtins/math.h"

#include <cmath>
#include <cstdint>
#include <limits>

#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::snap_to_integer;
using builtins_detail::to_finite_value;

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
  // Converting a double outside `int`'s range is undefined, and the two
  // architectures disagree on what they produce: AArch64 saturates to INT_MAX
  // while x86-64 yields INT_MIN. A digit count like `1e50` therefore read as a
  // huge positive on one and a huge negative on the other, so the same formula
  // truncated to a no-op or collapsed to zero depending on the host. Clamp
  // before converting; every caller compares against thresholds (+/-308) far
  // inside this range, so saturating carries the same meaning as the real
  // magnitude.
  const double truncated = std::trunc(d);
  constexpr double kIntMax = 2147483647.0;
  constexpr double kIntMin = -2147483648.0;
  if (truncated >= kIntMax) {
    return std::numeric_limits<int>::max();
  }
  if (truncated <= kIntMin) {
    return std::numeric_limits<int>::min();
  }
  return static_cast<int>(truncated);
}

// Shared frame of TRUNC and the ROUND family: coerces `value`, reads
// `digits`, and returns `scale(value, 10^digits)` for the in-range cases.
//
// Extreme `digits` are clamped so `10^digits` cannot overflow to +-Inf:
// beyond ~308 places a double has no digits left to round (a no-op), and
// below ~-308 every finite double rounds to a multiple of 10^|digits|, i.e. 0.
template <typename Scale>
Value round_to_digits(const Value* args, std::uint32_t arity, Scale scale) {
  auto value = coerce_to_number(args[0]);
  if (!value) {
    return Value::error(value.error());
  }
  auto digits = read_digits(args, arity, 1);
  if (!digits) {
    return Value::error(digits.error());
  }
  const int d = digits.value();
  if (d >= 308) {
    return Value::number(value.value());
  }
  if (d <= -308) {
    return Value::number(0.0);
  }
  const double factor = std::pow(10.0, d);
  if (std::isnan(factor) || std::isinf(factor)) {
    return Value::error(ErrorCode::Num);
  }
  return scale(value.value(), factor);
}

// TRUNC(value, digits?) - truncate toward zero. With no second arg or
// `digits = 0`, equivalent to `std::trunc(value)`. With `digits != 0`, the
// value is scaled by `10^digits`, truncated, then rescaled. `digits` may be
// negative (e.g. `TRUNC(1234.5, -1) = 1230`). A non-finite scale factor
// (caused by very large `|digits|`) yields `#NUM!`.
Value Trunc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  // Empty-string arguments surface #VALUE! rather than coercing to 0, same
  // rule as CEILING / FLOOR / MROUND.
  auto is_empty = [](const Value& v) {
    if (v.kind() != ValueKind::Text)
      return false;
    for (char c : v.as_text()) {
      if (c != ' ' && c != '\t')
        return false;
    }
    return true;
  };
  if (is_empty(args[0]) || (arity >= 2 && is_empty(args[1]))) {
    return Value::error(ErrorCode::Value);
  }
  return round_to_digits(args, arity, [](double x, double factor) {
    // Overflow on `value * factor` would truncate a no-op operation into a
    // spurious `#NUM!`. When the product overflows, the double's mantissa
    // already has no precision left for a fractional tail at this scale, so
    // truncation is a no-op and we return the input unchanged. Verified
    // against IronCalc TRUNC C57: `TRUNC(9.99999e+307, 5)` → 9.99999e+307.
    const double product = x * factor;
    if (std::isinf(product)) {
      return Value::number(x);
    }
    // `snap_to_integer` absorbs the same IEEE-754 near-integer noise ROUND /
    // CEILING / FLOOR already compensate for (e.g. `9.99 * 100 ==
    // 998.9999999999999`), so TRUNC/ROUNDDOWN/ROUNDUP don't lose the last
    // decimal digit that ROUND/CEILING/FLOOR keep on the same input.
    return to_finite_value(std::trunc(snap_to_integer(product)) / factor);
  });
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
  return to_finite_value(r);
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
  return to_finite_value(r);
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
// All three take `(value, digits)`. `digits` may be negative. They share the
// `round_to_digits` frame, but each keeps its own inline formula: the modes
// behave differently and a visible formula keeps each impl auditable.

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
  return round_to_digits(args, 2, [](double x, double factor) {
    const double scaled = x * factor;
    const double bias = std::copysign(std::fabs(scaled) * 2.0 * std::numeric_limits<double>::epsilon(), scaled);
    return to_finite_value(std::round(scaled + bias) / factor);
  });
}

// ROUNDDOWN - always toward zero. `ROUNDDOWN(2.99, 0) = 2`,
// `ROUNDDOWN(-2.99, 0) = -2`.
Value RoundDown(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return round_to_digits(args, 2, [](double x, double factor) {
    // `snap_to_integer` absorbs binary-representation noise the same way
    // TRUNC does (see its comment); without it, e.g. `0.29 * 100 ==
    // 28.999999999999996` truncates to 28 instead of 29.
    return to_finite_value(std::trunc(snap_to_integer(x * factor)) / factor);
  });
}

// ROUNDUP - always away from zero. `ROUNDUP(2.01, 0) = 3`,
// `ROUNDUP(-2.01, 0) = -3`. Positive inputs use `std::ceil`, negative
// inputs use `std::floor`; zero round-trips through either branch.
Value RoundUp(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return round_to_digits(args, 2, [](double x, double factor) {
    // `snap_to_integer` absorbs binary-representation noise the same way
    // TRUNC / ROUNDDOWN do (see TRUNC's comment above).
    const double scaled = snap_to_integer(x * factor);
    return to_finite_value((x > 0.0) ? std::ceil(scaled) / factor : std::floor(scaled) / factor);
  });
}

// --- Significance-aware rounding (legacy CEILING / FLOOR / MROUND) ------
//
// Mac Excel 365 applies an *asymmetric* sign-mismatch rule to the legacy
// CEILING / FLOOR forms:
//
//   * number > 0 AND significance < 0  -> #NUM!
//   * number < 0 AND significance > 0  -> numeric result via math ceil/floor
//   * signs match                      -> magnitude operation on |n|, |s|
//
// The pre-365 rule was symmetric #NUM! on either mismatch direction; Mac
// 365 dropped it only for the (negative-number, positive-significance)
// direction. `MROUND` still inherits the fully-symmetric #NUM! rule. The
// modern `*.MATH` variants (below) drop this rule entirely and use
// `|significance|` unconditionally.

// CEILING / FLOOR / MROUND treat empty-string text as `#VALUE!` rather
// than coercing it to zero (which is what the general arithmetic rule in
// `coerce_to_number` does). This matters because the subsequent `s == 0`
// branch would otherwise convert `FLOOR(10, "")` to `#DIV/0!` -- Excel's
// observed behaviour is `#VALUE!`.
inline bool is_empty_text(const Value& v) {
  if (v.kind() != ValueKind::Text)
    return false;
  const std::string_view t = v.as_text();
  for (char c : t) {
    if (c != ' ' && c != '\t')
      return false;
  }
  return true;
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

// Shared frame of legacy CEILING / FLOOR. `zero_significance` is the result
// for `significance == 0`; `round` is `ceil` or `floor` on the quotient.
//
// Matching signs round the magnitude, then restore the sign; mismatched signs
// round the signed quotient. `snap_to_integer` absorbs the quotient's
// floating-point noise (`FLOOR(7.1, 0.1)` is 7.1, not 7).
template <typename RoundFn>
Value legacy_significance_round(const Value* args, const Value& zero_significance, RoundFn round) {
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
  // Mac Excel 365 asymmetric sign-mismatch rule: positive number with
  // negative significance is #NUM!; the reverse direction falls through to
  // the signed branch. Runs before the `s == 0.0` check (s < 0 implies s != 0).
  if (n > 0.0 && s < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (s == 0.0) {
    return zero_significance;
  }
  const double abs_s = std::fabs(s);
  const double r = (signum(n) == signum(s)) ? signum(n) * round(snap_to_integer(std::fabs(n) / abs_s)) * abs_s
                                            : round(snap_to_integer(n / abs_s)) * abs_s;
  return to_finite_value(r);
}

// CEILING(number, significance) - legacy: nearest multiple of
// `|significance|` in the direction determined by sign matching:
//
//   * number == 0  -> 0
//   * number > 0 AND significance < 0  -> #NUM! (Mac 365 asymmetric rule).
//   * significance == 0 -> 0 (no #DIV/0!, unlike FLOOR)
//   * sign(number) == sign(significance)  -> round AWAY from zero
//     (magnitude ceil on the positive half-line).
//   * number < 0 AND significance > 0  -> round toward +infinity
//     (math ceil on the signed value, matching CEILING.MATH defaults).
//
// Mac Excel 365 kept #NUM! for the (pos-num, neg-sig) direction only; the
// reverse (neg-num, pos-sig) produces a numeric result. See the oracle
// suite `floor_ceiling_edges.yaml` for the reference values.
Value Ceiling(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  // Legacy CEILING: significance of zero yields zero (no #DIV/0!).
  return legacy_significance_round(args, Value::number(0.0), [](double q) { return std::ceil(q); });
}

// FLOOR(number, significance) - legacy: nearest multiple of
// `|significance|` in the direction determined by sign matching. Mirror
// image of CEILING above, with two differences:
//
//   * significance == 0 -> #DIV/0! (Excel-documented quirk).
//   * Direction in the (neg-num, pos-sig) case is toward -infinity
//     (math floor).
//
// As with CEILING, Mac Excel 365 applies the asymmetric sign-mismatch
// rule: `number > 0 AND significance < 0` yields #NUM!; the reverse
// direction falls through to the math-floor branch and returns a numeric
// value. See oracle suite `floor_ceiling_edges.yaml` for reference.
Value Floor(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return legacy_significance_round(args, Value::error(ErrorCode::Div0), [](double q) { return std::floor(q); });
}

// Coerces two numeric arguments, rejecting a Bool in either slot with
// `#VALUE!` (MROUND and QUOTIENT, unlike the arithmetic operators, do not
// accept direct Bool operands).
Expected<builtins_detail::NumberPair, ErrorCode> coerce_non_bool_pair(const Value* args) {
  if (args[0].kind() == ValueKind::Bool || args[1].kind() == ValueKind::Bool) {
    return ErrorCode::Value;
  }
  auto first = coerce_to_number(args[0]);
  if (!first) {
    return first.error();
  }
  auto second = coerce_to_number(args[1]);
  if (!second) {
    return second.error();
  }
  return builtins_detail::NumberPair{first.value(), second.value()};
}

// MROUND(number, multiple) - nearest multiple of `|multiple|` to `number`,
// with ties rounded away from zero. Opposite-signed inputs yield #NUM!;
// `multiple = 0` returns 0 (Excel's documented quirk).
Value MRound(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  if (is_empty_text(args[0]) || is_empty_text(args[1])) {
    return Value::error(ErrorCode::Value);
  }
  // Excel 365 rejects direct Bool arguments to MROUND with #VALUE!
  // (mirrors the strict-Bool rejection in BIN2*/OCT2*/HEX2*). Number
  // and Text arguments still coerce normally.
  auto pair = coerce_non_bool_pair(args);
  if (!pair) {
    return Value::error(pair.error());
  }
  const double n = pair.value().first;
  const double m = pair.value().second;
  if (m == 0.0) {
    return Value::number(0.0);
  }
  if (n != 0.0 && signum(n) != signum(m)) {
    return Value::error(ErrorCode::Num);
  }
  const double abs_m = std::fabs(m);
  const double q = std::fabs(n) / abs_m;
  const double integer_quotient = std::floor(q);
  const double fraction = q - integer_quotient;
  // Mac Excel 16.112 ja-JP oracle-derived local cutoff: values below the
  // half boundary by at most C = 90 * 2^-54 = 0x1.68p-48 round up.
  constexpr double kMacExcelMRoundCutoff = 0x1.68p-48;
  const bool round_up = fraction >= 0.5 || (fraction < 0.5 && 0.5 - fraction <= kMacExcelMRoundCutoff);
  const double rounded_quotient = integer_quotient + (round_up ? 1.0 : 0.0);
  // Restore the sign only after choosing the nearest magnitude quotient.
  const double r = signum(n) * rounded_quotient * abs_m;
  return to_finite_value(r);
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

// Shared body of CEILING.MATH (`ceiling = true`) and FLOOR.MATH. Positive
// inputs always round in the function's own direction; a non-zero `mode`
// flips the direction for negative inputs.
Value math_mode_rounding(const Value* args, std::uint32_t arity, bool ceiling) {
  auto number = coerce_to_number(args[0]);
  if (!number) {
    return Value::error(number.error());
  }
  auto sig = builtins_detail::read_optional_number(args, arity, 1, 1.0, /*check_finite=*/false);
  if (!sig) {
    return Value::error(sig.error());
  }
  const double significance = sig.value();
  auto flip = builtins_detail::read_optional_number(args, arity, 2, 0.0, /*check_finite=*/false);
  if (!flip) {
    return Value::error(flip.error());
  }
  const bool flip_negative = flip.value() != 0.0;
  const double n = number.value();
  if (n == 0.0) {
    return Value::number(0.0);
  }
  if (significance == 0.0) {
    return Value::number(0.0);
  }
  const double abs_s = std::fabs(significance);
  // `snap_to_integer` absorbs the IEEE-754 noise that would otherwise make
  // e.g. `CEILING.MATH(7.1, 0.1)` return `7` instead of `7.1` (7.1 / 0.1 is
  // 70.999... rather than exactly 71). Mirrors the legacy CEILING / FLOOR
  // path so the modern variants snap exact multiples the same way.
  const double scaled = snap_to_integer(n / abs_s);
  const bool use_ceil = (n > 0.0 || !flip_negative) ? ceiling : !ceiling;
  const double rounded = use_ceil ? std::ceil(scaled) : std::floor(scaled);
  const double r = rounded * abs_s;
  return to_finite_value(r);
}

Value CeilingMath(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return math_mode_rounding(args, arity, /*ceiling=*/true);
}

Value FloorMath(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return math_mode_rounding(args, arity, /*ceiling=*/false);
}

// --- Parity-aware rounding ----------------------------------------------
//
// EVEN and ODD both round AWAY from zero to the nearest integer of the
// required parity. ODD has a single documented special case: `ODD(0) = 1`
// (because 0 has no odd neighbour in the "nearest" direction, Excel
// promotes it to +1). All other zero / non-integer inputs behave as a
// normal away-from-zero rounding followed by a parity fix-up.

// Shared body of EVEN / ODD: rounds away from zero to the nearest integer,
// then steps one further away when that integer has the wrong parity.
Value round_away_to_parity(const Value* args, bool want_odd) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  if (x == 0.0) {
    return Value::number(want_odd ? 1.0 : 0.0);
  }
  // `std::ceil / std::floor` on the signed value is the away-from-zero step.
  const double away = (x > 0.0) ? std::ceil(x) : std::floor(x);
  const bool is_odd = std::fmod(std::fabs(away), 2.0) != 0.0;
  const double r = (is_odd == want_odd) ? away : away + ((x > 0.0) ? 1.0 : -1.0);
  return to_finite_value(r);
}

// EVEN(x) - nearest even integer, rounded AWAY from zero.
//   `EVEN(1.5) = 2`, `EVEN(3) = 4`, `EVEN(-1.5) = -2`, `EVEN(-2.1) = -4`,
//   `EVEN(0) = 0`.
Value Even(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return round_away_to_parity(args, /*want_odd=*/false);
}

// ODD(x) - nearest odd integer, rounded AWAY from zero. `ODD(0) = 1` is the
// documented quirk; otherwise symmetric to EVEN.
//   `ODD(1.5) = 3`, `ODD(2) = 3`, `ODD(-1.5) = -3`, `ODD(-2) = -3`.
Value Odd(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return round_away_to_parity(args, /*want_odd=*/true);
}

// QUOTIENT(numerator, denominator) - integer division, truncated TOWARD
// zero (not floor-division). Matches Excel: `QUOTIENT(-10, 3) = -3`.
// `denominator == 0` -> `#DIV/0!`. Very large quotients fall through the
// finite-check and surface `#NUM!`.
Value Quotient(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  // Excel's QUOTIENT rejects boolean operands (#VALUE!) even though MOD and
  // arithmetic operators accept them. Verified against Mac Excel 365 / IronCalc.
  auto pair = coerce_non_bool_pair(args);
  if (!pair) {
    return Value::error(pair.error());
  }
  if (pair.value().second == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = std::trunc(pair.value().first / pair.value().second);
  return to_finite_value(r);
}

}  // namespace

void register_math_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"ABS", 1u, 1u, &Abs},
      {"SIGN", 1u, 1u, &Sign},
      {"INT", 1u, 1u, &Int_},
      {"TRUNC", 1u, 2u, &Trunc},
      {"SQRT", 1u, 1u, &Sqrt},
      {"MOD", 2u, 2u, &Mod},
      {"POWER", 2u, 2u, &Power},
      {"QUOTIENT", 2u, 2u, &Quotient},
      {"ROUND", 2u, 2u, &Round},
      {"ROUNDDOWN", 2u, 2u, &RoundDown},
      {"ROUNDUP", 2u, 2u, &RoundUp},
      {"EVEN", 1u, 1u, &Even},
      {"ODD", 1u, 1u, &Odd},
      {"CEILING", 2u, 2u, &Ceiling},
      {"FLOOR", 2u, 2u, &Floor},
      {"MROUND", 2u, 2u, &MRound, true, false, false, false, false, FunctionDef::BlankScalarPolicy::RejectLiteralEmpty,
       ErrorCode::NA},
      {"CEILING.MATH", 1u, 3u, &CeilingMath},
      {"FLOOR.MATH", 1u, 3u, &FloorMath},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
