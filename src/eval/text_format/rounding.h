//
// Internal helpers for Excel-style decimal display rounding.

#ifndef FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_
#define FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_

#include <cmath>
#include <cstddef>
#include <cstring>
#include <string_view>

#include "utils/double_format.h"
#include "utils/double_parse.h"
#include "utils/elementary_math.h"
#include "utils/number_text.h"

namespace formulon {
namespace text_format {

// Excel stores and displays numeric worksheet values with at most this many
// significant decimal digits, and decides display ties on that decimal view.
constexpr int kExcelSignificantDigits = 15;

// Snaps `value` to its 15-significant-digit decimal. Keeps binary residue
// such as 0.1 + 0.2 out of a rendered digit string.
inline double round_to_15_significant_digits(double value) noexcept {
  if (value == 0.0 || !std::isfinite(value)) {
    return value;
  }
  const double exponent = std::floor(elementary::log10(std::fabs(value)));
  const double quantum = elementary::pow(10.0, exponent - 14.0);
  if (quantum == 0.0 || !std::isfinite(quantum)) {
    return value;
  }
  return std::round(value / quantum) * quantum;
}

// The 15-significant-digit exponential view of `v` (`-d.ddddddddddddddde+XX`)
// split into its mantissa text and decimal exponent. `has_exponent` is false
// for a non-finite `v`, which has no exponent.
struct Exponential15 {
  char text[48];
  std::size_t mantissa_length = 0;
  int exp10 = 0;
  bool has_exponent = false;

  std::string_view mantissa() const { return std::string_view(text, mantissa_length); }
};

inline Exponential15 exponential15(double v) {
  Exponential15 out;
  const int n = format_exponential(out.text, sizeof(out.text), excel_decimal_view(v, kExcelSignificantDigits),
                                   kExcelSignificantDigits - 1);
  const std::string_view text(out.text, n > 0 ? static_cast<std::size_t>(n) : 0U);
  const std::size_t e = text.find('e');
  if (e != std::string_view::npos) {
    out.mantissa_length = e;
    out.exp10 = static_cast<int>(parse_double_prefix(text.substr(e + 1U)).value);
    out.has_exponent = true;
  }
  return out;
}

// Rounds ties away from zero at `decimals` places. Negative `decimals`
// rounds to the left of the decimal point.
//
// The tie is decided on the 15-significant-digit decimal of `value`, never on
// the nearest double of `value * 10^decimals`. A decimal literal such as
// 1.005 is stored one ULP below its decimal value, so `1.005 * 100` yields
// 100.49999999999999 and a bare `std::round` would round it down while Excel
// (and therefore `ROUND`) rounds it up. Reading the digits back out of a
// correctly rounded 15-digit conversion makes the decision on the decimal
// Excel itself displays, which also drops binary residue (0.1 + 0.2).
inline double round_display_decimal(double value, int decimals) noexcept {
  if (value == 0.0 || !std::isfinite(value)) {
    return value;
  }
  const Exponential15 view = exponential15(std::fabs(value));
  if (!view.has_exponent) {
    return value;
  }
  // `d.dddddddddddddd`: collect the significant digits without assuming where
  // the decimal point sits.
  char digits[kExcelSignificantDigits + 1];
  int digit_count = 0;
  for (const char c : view.mantissa()) {
    if (c >= '0' && c <= '9' && digit_count < kExcelSignificantDigits) {
      digits[digit_count++] = c;
    }
  }
  if (digit_count < kExcelSignificantDigits) {
    return value;
  }
  int exponent = view.exp10;

  // `keep` digits survive the cut; the first dropped digit is `digits[keep]`.
  // `keep >= 15` means the decimal already ends at or above the requested
  // place, `keep < 0` that the whole value sits below the rounding unit.
  int keep = exponent + decimals + 1;
  if (keep > kExcelSignificantDigits) {
    keep = kExcelSignificantDigits;
  }
  if (keep < 0) {
    return std::copysign(0.0, value);
  }
  const bool round_away = keep < kExcelSignificantDigits && digits[keep] >= '5';
  if (keep == 0) {
    if (!round_away) {
      return std::copysign(0.0, value);
    }
    // Every digit was dropped and the first of them rounds away from zero:
    // the result is one unit at the target place.
    digits[0] = '1';
    keep = 1;
    exponent = -decimals;
  } else if (round_away) {
    int carry = keep - 1;
    for (; carry >= 0; --carry) {
      if (digits[carry] != '9') {
        ++digits[carry];
        break;
      }
      digits[carry] = '0';
    }
    if (carry < 0) {
      // All-nines carried out into one more leading digit (9.99 -> 10.0).
      digits[0] = '1';
      for (int i = 1; i < keep; ++i) {
        digits[i] = '0';
      }
      ++exponent;
    }
  }

  // Rebuild through a point-free literal: the mantissa is an integer and the
  // exponent carries the scale, so the conversion back is exact for every
  // representable magnitude and never touches a locale decimal point.
  char rebuilt[kExcelSignificantDigits + 8];
  const std::size_t mantissa_length = static_cast<std::size_t>(keep);
  std::memcpy(rebuilt, digits, mantissa_length);
  rebuilt[mantissa_length] = 'e';
  const std::size_t exponent_offset = mantissa_length + 1;
  const int exponent_length =
      format_signed(rebuilt + exponent_offset, sizeof(rebuilt) - exponent_offset, exponent - keep + 1);
  if (exponent_length <= 0 || static_cast<std::size_t>(exponent_length) >= sizeof(rebuilt) - exponent_offset) {
    return value;
  }
  const std::string_view literal(rebuilt, exponent_offset + static_cast<std::size_t>(exponent_length));
  return std::copysign(parse_double_prefix(literal).value, value);
}

}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_
