//
// Internal helpers for Excel-style decimal display rounding.

#ifndef FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_
#define FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_

#include <cmath>
#include <cstddef>
#include <cstdio>
#include <cstdlib>

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
  const double exponent = std::floor(std::log10(std::fabs(value)));
  const double quantum = std::pow(10.0, exponent - 14.0);
  if (quantum == 0.0 || !std::isfinite(quantum)) {
    return value;
  }
  return std::round(value / quantum) * quantum;
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
  char text[32];
  const int written = std::snprintf(text, sizeof(text), "%.*e", kExcelSignificantDigits - 1, std::fabs(value));
  if (written <= 0 || static_cast<std::size_t>(written) >= sizeof(text)) {
    return value;
  }
  // `%.14e` yields `d.dddddddddddddde[+-]XX`. Collect the significant digits
  // without assuming where the decimal point sits, then read the exponent of
  // the leading digit.
  char digits[kExcelSignificantDigits + 1];
  int digit_count = 0;
  int cursor = 0;
  for (; cursor < written && text[cursor] != 'e' && text[cursor] != 'E'; ++cursor) {
    if (text[cursor] >= '0' && text[cursor] <= '9' && digit_count < kExcelSignificantDigits) {
      digits[digit_count++] = text[cursor];
    }
  }
  if (digit_count < kExcelSignificantDigits || cursor >= written) {
    return value;
  }
  ++cursor;  // Step over the `e`.
  bool exponent_negative = false;
  if (cursor < written && (text[cursor] == '+' || text[cursor] == '-')) {
    exponent_negative = text[cursor] == '-';
    ++cursor;
  }
  int exponent = 0;
  for (; cursor < written && text[cursor] >= '0' && text[cursor] <= '9'; ++cursor) {
    exponent = exponent * 10 + (text[cursor] - '0');
  }
  if (exponent_negative) {
    exponent = -exponent;
  }

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
  digits[keep] = '\0';

  // Rebuild through a point-free literal: the mantissa is an integer and the
  // exponent carries the scale, so the conversion back is exact for every
  // representable magnitude and never touches a locale decimal point.
  char rebuilt[kExcelSignificantDigits + 8];
  const int rebuilt_length = std::snprintf(rebuilt, sizeof(rebuilt), "%se%d", digits, exponent - keep + 1);
  if (rebuilt_length <= 0 || static_cast<std::size_t>(rebuilt_length) >= sizeof(rebuilt)) {
    return value;
  }
  return std::copysign(std::strtod(rebuilt, nullptr), value);
}

}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_
