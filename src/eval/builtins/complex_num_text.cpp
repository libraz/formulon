//
// Parsing and formatting of complex-number text ("a+bi" / "a+bj") for the
// complex-number built-ins.

#include "eval/builtins/complex_num_text.h"

#include <cmath>
#include <cstddef>
#include <cstdlib>
#include <limits>
#include <optional>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/locale_text.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/number_text.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using complex_detail::Complex;

// Text of one part to a number: the ordinary locale text-to-number conversion.
// A boolean name is not a number.
Expected<double, ErrorCode> part_to_number(std::string_view text) {
  if (text == "TRUE" || text == "FALSE") {
    return ErrorCode::Num;
  }
  auto n = coerce_text_to_number(text);
  if (!n) {
    return ErrorCode::Num;
  }
  return n.value();
}

// A part must be zero or lie in [1e-309, 1e308).
bool part_in_range(double v) {
  if (std::isnan(v) || std::isinf(v)) {
    return false;
  }
  return v == 0.0 || (std::fabs(v) >= 1e-309 && std::fabs(v) < 1e308);
}

// Parses a complex-number text view. Returns `std::nullopt` on any
// malformed input (caller should surface #NUM!). Spaces are dropped anywhere
// in the text. Accepted shapes:
//
//   Pure real:       <number>
//   Pure imaginary:  [+-]?(<number>)?[ij]   -- empty coefficient means 1
//   Full:            <number>[+-]<number>?[ij]
//
// Each part goes through the locale text-to-number conversion. The parts split
// at the first `+`/`-` after position 0 that does not follow `e`/`E`; the
// imaginary coefficient may carry a sign of its own (`3--4i` is 3+4i).
// `i`/`j` are lowercase only.
std::optional<Complex> parse_complex_text(std::string_view raw) {
  std::string src;
  for (char c : raw) {
    if (c != ' ') {
      src.push_back(c);
    }
  }
  if (src.empty()) {
    return std::nullopt;
  }
  const bool has_i = src.find('i') != std::string::npos;
  const bool has_j = src.find('j') != std::string::npos;
  if (has_i && has_j) {
    return std::nullopt;
  }
  if (!has_i && !has_j) {
    auto re = part_to_number(src);
    if (!re || !part_in_range(re.value())) {
      return std::nullopt;
    }
    return Complex{re.value(), 0.0, 'i'};
  }
  const char suffix = has_j ? 'j' : 'i';
  std::size_t suffix_count = 0;
  for (char c : src) {
    suffix_count += c == suffix ? 1U : 0U;
  }
  if (src.back() != suffix || suffix_count != 1U) {
    return std::nullopt;
  }
  const std::string_view body = std::string_view(src).substr(0, src.size() - 1);
  if (body.empty() || body == "+") {
    return Complex{0.0, 1.0, suffix};
  }
  if (body == "-") {
    return Complex{0.0, -1.0, suffix};
  }
  std::size_t split = std::string_view::npos;
  for (std::size_t idx = 1; idx < body.size(); ++idx) {
    if ((body[idx] == '+' || body[idx] == '-') && body[idx - 1] != 'e' && body[idx - 1] != 'E') {
      split = idx;
      break;
    }
  }
  if (split == std::string_view::npos) {
    const std::string_view coef = body.front() == '+' ? body.substr(1) : body;
    auto im = part_to_number(coef);
    if (!im || !part_in_range(im.value())) {
      return std::nullopt;
    }
    return Complex{0.0, im.value(), suffix};
  }
  std::string_view real_part = body.substr(0, split);
  if (real_part.front() == '+') {
    real_part = real_part.substr(1);
  }
  const std::string_view imag_part = body.substr(split);
  const char sign = imag_part.front();
  const std::string_view coef = imag_part.substr(1);
  auto re = part_to_number(real_part);
  if (!re) {
    return std::nullopt;
  }
  double magnitude = 1.0;
  if (!coef.empty()) {
    auto m = part_to_number(coef);
    if (!m) {
      return std::nullopt;
    }
    magnitude = m.value();
  }
  const double im = sign == '-' ? -magnitude : magnitude;
  if (!part_in_range(re.value()) || !part_in_range(im)) {
    return std::nullopt;
  }
  return Complex{re.value(), im, suffix};
}

}  // namespace

namespace complex_detail {

// Front-door parser: accepts Text / Number / Bool / Blank / Error. Returns
// an `Expected` so we can propagate the Excel error code verbatim.
Expected<Complex, ErrorCode> parse_complex_value(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Number: {
      const double d = v.as_number();
      if (std::isnan(d) || std::isinf(d)) {
        return ErrorCode::Num;
      }
      return Complex{d, 0.0, 'i'};
    }
    case ValueKind::Bool:
      return Complex{v.as_boolean() ? 1.0 : 0.0, 0.0, 'i'};
    case ValueKind::Blank:
      return Complex{0.0, 0.0, 'i'};
    case ValueKind::Text: {
      auto parsed = parse_complex_text(v.as_text());
      if (!parsed) {
        return ErrorCode::Num;
      }
      return *parsed;
    }
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

}  // namespace complex_detail

namespace {

// Text of one complex part: the value rounded to 15 significant digits, in
// fixed notation when that string (sign excluded) is at most 21 characters,
// otherwise mantissa (trailing zeros stripped) + E+dd / E-dd.
std::string complex_part_text(double v) {
  constexpr std::size_t kMaxFixedLength = 21;
  if (v == 0.0) {
    return "0";
  }
  // A subnormal part reads as a signed zero.
  if (flush_subnormal(v) == 0.0) {
    return v < 0.0 ? "-0" : "0";
  }
  char buf[40];
  format_exponential(buf, sizeof(buf), std::fabs(v), 14);
  // buf is "d.dddddddddddddde[+-]XX".
  std::string digits;
  const char* cursor = buf;
  for (; *cursor != 'e'; ++cursor) {
    if (*cursor != '.') {
      digits.push_back(*cursor);
    }
  }
  const int exponent = std::atoi(cursor + 1);
  while (digits.size() > 1 && digits.back() == '0') {
    digits.pop_back();
  }
  const std::size_t point = exponent >= 0 ? static_cast<std::size_t>(exponent) + 1 : 0;
  std::string fixed;
  if (exponent >= 0) {
    fixed = digits.substr(0, point);
    fixed.append(point > fixed.size() ? point - fixed.size() : 0, '0');
    if (digits.size() > point) {
      fixed += '.' + digits.substr(point);
    }
  } else {
    fixed = "0." + std::string(static_cast<std::size_t>(-exponent) - 1, '0') + digits;
  }
  std::string out = v < 0.0 ? "-" : "";
  if (fixed.size() <= kMaxFixedLength) {
    out += fixed;
  } else {
    out += digits.substr(0, 1);
    if (digits.size() > 1) {
      out += '.' + digits.substr(1);
    }
    out += exponent >= 0 ? "E+" : "E-";
    char exp_text[16];
    format_signed(exp_text, sizeof(exp_text), exponent >= 0 ? exponent : -exponent, 2);
    out += exp_text;
  }
  const char separator = locale_decimal_separator();
  if (separator != '.') {
    for (char& c : out) {
      if (c == '.') {
        c = separator;
      }
    }
  }
  return out;
}

}  // namespace

namespace complex_detail {

// Formats a complex triple into Excel's canonical string form: every IM*
// text result goes through here. An imaginary part of +-1 prints as the bare
// suffix and a zero real part is omitted.
std::string format_complex(double re, double im, char suffix) {
  if (im == 0.0) {
    return complex_part_text(re);
  }
  std::string out;
  if (re != 0.0) {
    out += complex_part_text(re);
    if (im > 0.0) {
      out.push_back('+');
    }
  }
  if (im == -1.0) {
    out.push_back('-');
  } else if (im != 1.0) {
    out += complex_part_text(im);
  }
  out.push_back(suffix);
  return out;
}

// Interns and returns a Text value for the complex triple. A part that is not
// finite is #NUM!, and a subnormal part flushes to zero before formatting (so
// it is omitted like any zero part).
Value text_complex(Complex z, Arena& arena) {
  if (!std::isfinite(z.re) || !std::isfinite(z.im)) {
    return Value::error(ErrorCode::Num);
  }
  constexpr double kMin = std::numeric_limits<double>::min();
  const double re = std::fabs(z.re) < kMin ? 0.0 : z.re;
  const double im = std::fabs(z.im) < kMin ? 0.0 : z.im;
  return Value::text(arena.intern(format_complex(re, im, z.suffix)));
}

}  // namespace complex_detail

}  // namespace eval
}  // namespace formulon
