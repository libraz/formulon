//
// Locale-independent decimal-to-double parser shared by the formula
// tokenizer, text-to-number coercion, the number-format scanner and the
// OOXML readers.
//
// Only the decimal grammar is accepted:
//
//   [ASCII whitespace] [+|-] (digits [. digits] | . digits) [(e|E) [+|-] digits]
//
// Hexadecimal floats and the `inf` / `infinity` / `nan` spellings that
// `strtod` also takes are not numbers anywhere in Excel, so they are not
// parsed here. Conversion is correctly rounded (round-half-even) through
// double-conversion's `Strtod`, the same result `strtod` gives for every
// input both accept, and it never consults the process locale.

#ifndef FORMULON_UTILS_DOUBLE_PARSE_H_
#define FORMULON_UTILS_DOUBLE_PARSE_H_

#include <cstddef>
#include <string_view>

namespace formulon {

/// Outcome of parsing the longest decimal-number prefix of a string.
struct ParsedDouble {
  /// Bytes consumed, including leading whitespace; 0 when the text does not
  /// start with a number (then `value` is 0).
  std::size_t consumed = 0;
  /// Correctly rounded value; ±inf on overflow.
  double value = 0.0;
  /// The decimal value lies outside the normal double range: it overflowed
  /// to ±inf, or nonzero digits rounded to zero or to a subnormal.
  bool out_of_range = false;
};

/// Parses the longest prefix of `text` that matches the decimal grammar in
/// the header comment. An exponent marker with no digits after it is not
/// consumed (`"1e"` consumes one byte).
ParsedDouble parse_double_prefix(std::string_view text) noexcept;

/// Parses `text` as one decimal number with nothing after it. Returns false
/// for empty text or trailing bytes; the value is written even when it is
/// out of range, so callers that need a finite result check it themselves.
bool parse_double_exact(std::string_view text, double* out) noexcept;

}  // namespace formulon

#endif  // FORMULON_UTILS_DOUBLE_PARSE_H_
