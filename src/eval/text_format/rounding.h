//
// Internal helpers for Excel-style decimal display rounding.

#ifndef FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_
#define FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_

#include <cmath>
#include <limits>

namespace formulon {
namespace text_format {

// Excel stores and displays numeric worksheet values with at most 15
// significant decimal digits. Normalising before a fixed-decimal render keeps
// binary residue such as 0.1 + 0.2 out of TEXT/FIXED/DOLLAR output.
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
// rounds to the left of the decimal point. A scale that cannot safely affect
// the binary value is deliberately left unchanged rather than overflowing.
inline double round_display_decimal(double value, int decimals) noexcept {
  value = round_to_15_significant_digits(value);
  const double scale = std::pow(10.0, std::fabs(static_cast<double>(decimals)));
  if (scale == 0.0 || !std::isfinite(scale)) {
    return value;
  }
  if (decimals < 0) {
    return std::round(value / scale) * scale;
  }
  if (std::fabs(value) > std::numeric_limits<double>::max() / scale) {
    return value;
  }
  return std::round(value * scale) / scale;
}

}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_ROUNDING_H_
