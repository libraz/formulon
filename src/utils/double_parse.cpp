//
// Implementation of the shared decimal parser. See `double_parse.h` for the
// contract.
//
// The scan collects significant digits and a decimal exponent, then hands
// them to double-conversion's `Strtod`, which rounds correctly. `Strtod`
// looks at no more than 780 significant digits; when more nonzero digits
// follow, a trailing sticky `1` keeps the rounding direction, the same
// device `Strtod` uses when it cuts a longer buffer itself.

#include "utils/double_parse.h"

#include <cfloat>
#include <cmath>
#include <cstring>

#include "double-conversion/strtod.h"

namespace formulon {
namespace {

constexpr int kMaxDigits = 780;
// Exponent digits beyond this magnitude cannot change the result: every
// value is already 0 or ±inf, whatever the mantissa length.
constexpr int kExponentCap = 100000;

bool is_c_space(char c) noexcept {
  return c == ' ' || (c >= '\t' && c <= '\r');
}

bool is_digit(char c) noexcept {
  return c >= '0' && c <= '9';
}

}  // namespace

ParsedDouble parse_double_prefix(std::string_view text) noexcept {
  ParsedDouble out;
  const std::size_t n = text.size();
  std::size_t i = 0;
  while (i < n && is_c_space(text[i])) {
    ++i;
  }
  bool negative = false;
  if (i < n && (text[i] == '+' || text[i] == '-')) {
    negative = text[i] == '-';
    ++i;
  }

  char digits[kMaxDigits + 1];
  int count = 0;
  int exponent = 0;
  bool mantissa_digit = false;
  bool dropped_nonzero = false;
  for (; i < n && is_digit(text[i]); ++i) {
    mantissa_digit = true;
    if (count == 0 && text[i] == '0') {
      continue;
    }
    if (count < kMaxDigits) {
      digits[count++] = text[i];
    } else {
      ++exponent;
      dropped_nonzero = dropped_nonzero || text[i] != '0';
    }
  }
  if (i < n && text[i] == '.' && (mantissa_digit || (i + 1 < n && is_digit(text[i + 1])))) {
    for (++i; i < n && is_digit(text[i]); ++i) {
      mantissa_digit = true;
      if (count == 0 && text[i] == '0') {
        --exponent;
      } else if (count < kMaxDigits) {
        digits[count++] = text[i];
        --exponent;
      } else {
        dropped_nonzero = dropped_nonzero || text[i] != '0';
      }
    }
  }
  if (!mantissa_digit) {
    return out;
  }

  if (i < n && (text[i] == 'e' || text[i] == 'E')) {
    std::size_t j = i + 1;
    bool exponent_negative = false;
    if (j < n && (text[j] == '+' || text[j] == '-')) {
      exponent_negative = text[j] == '-';
      ++j;
    }
    if (j < n && is_digit(text[j])) {
      int e = 0;
      for (; j < n && is_digit(text[j]); ++j) {
        if (e < kExponentCap) {
          e = e * 10 + (text[j] - '0');
        }
      }
      exponent += exponent_negative ? -e : e;
      i = j;
    }
  }
  out.consumed = i;

  if (count == 0) {
    out.value = negative ? -0.0 : 0.0;
    return out;
  }
  if (dropped_nonzero) {
    digits[count++] = '1';
    --exponent;
  }
  const double magnitude = double_conversion::Strtod(double_conversion::Vector<const char>(digits, count), exponent);
  out.value = negative ? -magnitude : magnitude;
  out.out_of_range = std::isinf(magnitude) || magnitude < DBL_MIN;
  return out;
}

bool parse_double_exact(std::string_view text, double* out) noexcept {
  const ParsedDouble parsed = parse_double_prefix(text);
  if (parsed.consumed == 0 || parsed.consumed != text.size()) {
    return false;
  }
  *out = parsed.value;
  return true;
}

}  // namespace formulon

#if defined(FORMULON_WASM)
// Target of the WASM link's `--wrap=strtod` (cmake/FormulonWasm.cmake): the
// only remaining strtod callers are inside pugixml, and routing them here
// keeps libc's scanner out of the module.
extern "C" double __wrap_strtod(const char* str, char** endptr) {  // NOLINT(bugprone-reserved-identifier)
  const formulon::ParsedDouble parsed = formulon::parse_double_prefix(std::string_view(str, std::strlen(str)));
  if (endptr != nullptr) {
    *endptr = const_cast<char*>(str + parsed.consumed);  // NOLINT(cppcoreguidelines-pro-type-const-cast)
  }
  return parsed.value;
}
#endif
