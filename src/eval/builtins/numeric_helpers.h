//
// Shared scalar-numeric helpers used across the built-in catalog. These
// were previously duplicated (4x for `kPi`, 5x for the `non-finite ->
// #NUM!` wrap, and 3x for `read_(required|optional)_number` /
// `NumberPair` / `NumberTriple`); centralising them keeps the math /
// stats / financial / distributions / complex / dynamic-array TUs in
// lock-step on coercion semantics and lets the compiler share one
// definition across translation units.
//
// All helpers are inline so the header is self-contained; no
// numeric_helpers.cpp is needed.

#ifndef FORMULON_EVAL_BUILTINS_NUMERIC_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_NUMERIC_HELPERS_H_

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <limits>
#include <string_view>

#include "eval/coerce.h"
#include "utils/double_format.h"
#include "utils/double_parse.h"
#include "utils/expected.h"
#include "utils/number_text.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace builtins_detail {

/// Mathematical constant pi to ~15 significant digits. Matches the value
/// `std::acos(-1.0)` produces on any IEEE-754 system, which keeps
/// `RADIANS(180) == kPi` exact and the normal-PDF normalisation
/// (`1 / sqrt(2 * kPi)`) byte-for-byte identical across math, stats,
/// and distributions TUs.
inline constexpr double kPi = 3.14159265358979323846;

/// Wraps a scalar result in a `Value::number`, surfacing `#NUM!` for
/// non-finite (`NaN` / `Inf`) values. This is the universal "finalise a
/// computed double" convention used by EXP, the trig family, the
/// stats descriptive functions, the probability distributions, and the
/// time-value-of-money financials. Centralising it ensures every
/// numeric impl reports overflow / domain-violation via the same code
/// path.
inline Value to_finite_value(double r) noexcept {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

/// The 15-significant-digit decimal view of a quotient, applied before a
/// CEILING / FLOOR style rounding step so that an exact multiple (7.1 / 0.1)
/// and the noise just below or above it round the way Excel's displayed value does.
inline double significance_quotient(double quotient) noexcept {
  return excel_decimal_view(quotient, 15);
}

/// (a, b) pair returned by `read_number_pair`. Used by distribution
/// builtins (chi-squared / F-dist / fisher / phi) and the financial
/// TVM family for compact two-argument argument extraction.
struct NumberPair {
  double first;
  double second;
};

/// (a, b, c) triple returned by `read_number_triple`. Used by the
/// three-argument distribution builtins (normal / beta / gamma /
/// weibull / lognormal) and several financial / stats helpers.
struct NumberTriple {
  double first;
  double second;
  double third;
};

/// Decimal digits of |x| rendered to 15 significant digits: value = 0.d1d2...dn * 10^exponent with
/// trailing zeros dropped (`count == 0` for zero).
struct DecimalView {
  char digits[16];
  int count;
  int exponent;
  bool negative;
};

inline DecimalView decimal_view15(double x) noexcept {
  DecimalView view{};
  view.negative = x < 0.0;
  const double magnitude = std::fabs(x);
  if (magnitude == 0.0 || !std::isfinite(magnitude)) {
    return view;
  }
  char text[40];
  format_exponential(text, sizeof(text), excel_decimal_view(magnitude, 15), 14);
  // text is "d.dddddddddddddde[+-]XX".
  const char* cursor = text;
  int n = 0;
  for (; *cursor != 'e'; ++cursor) {
    if (*cursor != '.') {
      view.digits[n++] = *cursor;
    }
  }
  view.exponent = std::atoi(cursor + 1) + 1;
  while (n > 1 && view.digits[n - 1] == '0') {
    --n;
  }
  view.count = n;
  return view;
}

/// The double nearest 0.`digits` * 10^`exponent` (a correctly rounded decimal conversion), signed.
inline double decimal_to_double(const char* digits, int count, int exponent, bool negative) noexcept {
  if (count == 0) {
    return 0.0;
  }
  char literal[48];
  std::memcpy(literal, digits, static_cast<std::size_t>(count));
  literal[count] = 'e';
  const std::size_t offset = static_cast<std::size_t>(count) + 1;
  const int length =
      format_signed(literal + offset, sizeof(literal) - offset, static_cast<long long>(exponent) - count);
  double value = 0.0;
  parse_double_exact(std::string_view(literal, offset + static_cast<std::size_t>(length)), &value);
  return negative ? -value : value;
}

/// 10^n as a correctly rounded double.
inline double power_of_ten_exact(int n) noexcept {
  char literal[16] = {'1', 'e'};
  const int length = format_signed(literal + 2, sizeof(literal) - 2, n);
  double value = 1.0;
  parse_double_exact(std::string_view(literal, static_cast<std::size_t>(2 + length)), &value);
  return value;
}

/// ROUNDDOWN / TRUNC to `digits` decimals: keeps the first (digits + exponent) digits of the 15-digit
/// view and rebuilds the double.
inline double round_down_decimals(double x, int digits) noexcept {
  const DecimalView view = decimal_view15(x);
  const int keep = digits + view.exponent;
  if (view.count == 0 || keep <= 0) {
    return 0.0;
  }
  return decimal_to_double(view.digits, keep < view.count ? keep : view.count, view.exponent, view.negative);
}

/// ROUND to `digits` decimals: the digit string of the 15-digit view rounded half away from zero.
inline double round_half_away_decimals(double x, int digits) noexcept {
  DecimalView view = decimal_view15(x);
  const int keep = digits + view.exponent;
  if (view.count == 0 || keep < 0) {
    return 0.0;
  }
  if (keep >= view.count) {
    return decimal_to_double(view.digits, view.count, view.exponent, view.negative);
  }
  const bool round_up = view.digits[keep] >= '5';
  if (!round_up) {
    return keep == 0 ? 0.0 : decimal_to_double(view.digits, keep, view.exponent, view.negative);
  }
  if (keep == 0) {
    view.digits[0] = '1';
    return decimal_to_double(view.digits, 1, view.exponent + 1, view.negative);
  }
  int i = keep - 1;
  while (i >= 0 && view.digits[i] == '9') {
    view.digits[i--] = '0';
  }
  int count = keep;
  int exponent = view.exponent;
  if (i >= 0) {
    ++view.digits[i];
  } else {
    // A carry out of the leading digit adds one: 99.7 -> 100.
    view.digits[0] = '1';
    count = 1;
    ++exponent;
  }
  return decimal_to_double(view.digits, count, exponent, view.negative);
}

/// ROUNDUP: ROUNDDOWN, plus one unit in the last kept place (away from zero) when digits were dropped.
inline double round_up_decimals(double x, int digits) noexcept {
  const double down = round_down_decimals(x, digits);
  if (down == x) {
    return down;
  }
  const double step = power_of_ten_exact(-digits);
  return x > 0.0 ? down + step : down - step;
}

/// 1-based ascending position of the k-th largest (`want_large`) or k-th
/// smallest of `n` values: `#NUM!` when `k < 1` or the (near-integer-snapped
/// for LARGE) k exceeds `n`. A fractional k picks LARGE's `trunc(n + 1 - k)`
/// and SMALL's `trunc(k)`. Shared by LARGE / SMALL and AGGREGATE 14 / 15.
inline Expected<std::size_t, ErrorCode> kth_ascending_position(double k_raw, std::size_t n, bool want_large) {
  const double k = want_large ? snap_near_integer(k_raw) : k_raw;
  const auto count = static_cast<double>(n);
  if (n == 0U || k_raw < 1.0 || k > count) {
    return ErrorCode::Num;
  }
  return static_cast<std::size_t>(want_large ? std::trunc(count + 1.0 - k) : std::trunc(k));
}

/// Reads a required numeric argument from `args[index]`; a non-finite value
/// surfaces as `#NUM!` from `coerce_to_number`.
inline Expected<double, ErrorCode> read_required_number(const Value* args, std::uint32_t index) {
  return coerce_to_number(args[index]);
}

/// Reads an optional trailing numeric argument at position `index`,
/// returning `default_value` when `arity <= index`. Otherwise behaves
/// like `read_required_number`. The default is returned by value so
/// callers never have to worry about the lifetime of a sentinel
/// reference.
inline Expected<double, ErrorCode> read_optional_number(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                        double default_value) {
  if (arity <= index) {
    return default_value;
  }
  return read_required_number(args, index);
}

/// Reads two required numeric arguments and returns them as a
/// `NumberPair`. Propagates the left-most coercion / non-finite error.
inline Expected<NumberPair, ErrorCode> read_number_pair(const Value* args, std::uint32_t first_index,
                                                        std::uint32_t second_index) {
  auto first = read_required_number(args, first_index);
  if (!first) {
    return std::move(first.error());
  }
  auto second = read_required_number(args, second_index);
  if (!second) {
    return std::move(second.error());
  }
  return NumberPair{first.value(), second.value()};
}

/// Reads three required numeric arguments and returns them as a
/// `NumberTriple`. Propagates the left-most coercion / non-finite
/// error.
inline Expected<NumberTriple, ErrorCode> read_number_triple(const Value* args, std::uint32_t first_index,
                                                            std::uint32_t second_index, std::uint32_t third_index) {
  auto first_second = read_number_pair(args, first_index, second_index);
  if (!first_second) {
    return std::move(first_second.error());
  }
  auto third = read_required_number(args, third_index);
  if (!third) {
    return std::move(third.error());
  }
  return NumberTriple{first_second.value().first, first_second.value().second, third.value()};
}

}  // namespace builtins_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_NUMERIC_HELPERS_H_
