#ifndef FORMULON_TESTS_ORACLE_SHOWN_PRECISION_H_
#define FORMULON_TESTS_ORACLE_SHOWN_PRECISION_H_

#include <cctype>
#include <cmath>
#include <cstddef>
#include <string_view>

namespace formulon::tests::oracle {

/// Half a unit in the last digit a number's text shows: the precision a value
/// Excel displayed as that text is known to. A `numeric_text` golden, a cell
/// Excel read back as its displayed text, is compared to the engine value with
/// at least this absolute tolerance. `"-471.7808"` gives 5e-5, `"2100"` gives
/// 0.5 and `"1.22514845490862E-16"` gives 5e-31. Either `.` or `,` may be the
/// decimal mark. Returns 0 when `text` is not a plain decimal or E-notation
/// number, so such text keeps the case's own tolerance. Complex-number text is
/// not read this way: Excel prints it to 15 significant digits and drops
/// trailing zeros, so its last shown digit understates its precision.
inline double shown_half_unit(std::string_view text) {
  std::size_t i = 0;
  const std::size_t n = text.size();
  while (i < n && std::isspace(static_cast<unsigned char>(text[i])) != 0) {
    ++i;
  }
  std::size_t end = n;
  while (end > i && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
    --end;
  }
  if (i < end && (text[i] == '+' || text[i] == '-')) {
    ++i;
  }
  std::size_t digits = 0;
  while (i < end && std::isdigit(static_cast<unsigned char>(text[i])) != 0) {
    ++i;
    ++digits;
  }
  int decimals = 0;
  if (i < end && (text[i] == '.' || text[i] == ',')) {
    ++i;
    while (i < end && std::isdigit(static_cast<unsigned char>(text[i])) != 0) {
      ++i;
      ++decimals;
    }
  }
  if (digits == 0 && decimals == 0) {
    return 0.0;
  }
  int exponent = 0;
  if (i < end && (text[i] == 'e' || text[i] == 'E')) {
    ++i;
    bool negative = false;
    if (i < end && (text[i] == '+' || text[i] == '-')) {
      negative = text[i] == '-';
      ++i;
    }
    const std::size_t exponent_start = i;
    while (i < end && std::isdigit(static_cast<unsigned char>(text[i])) != 0 && exponent < 10000) {
      exponent = exponent * 10 + (text[i] - '0');
      ++i;
    }
    if (i == exponent_start) {
      return 0.0;
    }
    exponent = negative ? -exponent : exponent;
  }
  if (i != end) {
    return 0.0;
  }
  return 0.5 * std::pow(10.0, exponent - decimals);
}

}  // namespace formulon::tests::oracle

#endif  // FORMULON_TESTS_ORACLE_SHOWN_PRECISION_H_
