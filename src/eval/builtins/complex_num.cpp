//
// Implementation of Excel's complex-number built-ins (COMPLEX + 24 IM*).
//
// Complex numbers travel through the Formulon calc engine as `Text` values
// of the form "a+bi" (or "a+bj"). The helpers below:
//
//   * `parse_complex`  parses such text into (real, imag, suffix); accepts
//     Number / Bool / Blank directly for Excel quirk compatibility.
//   * `format_complex` renders (real, imag, suffix) back to a canonical
//     Excel-compatible string, each part rounded to 15 significant digits.
//   * Individual IM* impls reuse those two helpers and a small `CplxOp`
//     set (multiplication, reciprocal, exp, ln) for arithmetic.
//
// Suffix propagation: a pure-real operand takes the other operand's suffix;
// mixing `i` and `j` is `#VALUE!` in IMSUM / IMPRODUCT and `#NUM!` in IMSUB /
// IMDIV.

#include "eval/builtins/complex_num.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "eval/locale_text.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/number_text.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::kPi;

// ---------------------------------------------------------------------------
// Core data types
// ---------------------------------------------------------------------------

/// A parsed complex number: (real, imag, suffix). `suffix` is always 'i'
/// or 'j'; `'i'` is the default when the input had no imaginary part.
struct Complex {
  double re;
  double im;
  char suffix;  // 'i' or 'j'
};

// ---------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Formatting
// ---------------------------------------------------------------------------

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

Value im_unary_number(double (*op)(Complex), const Value* args) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return builtins_detail::to_finite_value(op(z.value()));
}

// ---------------------------------------------------------------------------
// Primitives
// ---------------------------------------------------------------------------

// A product with a subnormal result flushed to zero.
double flushed_square(double x) noexcept {
  return flush_subnormal(x * x);
}

// a^2 + b^2 with flushed squares; NaN when the sum overflows.
double square_sum(double a, double b) noexcept {
  const double s = flushed_square(a) + flushed_square(b);
  return std::isinf(s) ? std::numeric_limits<double>::quiet_NaN() : s;
}

// |z|: a pure part is its own magnitude, otherwise sqrt of the square sum
// (NaN on overflow).
double naive_abs(Complex z) noexcept {
  if (z.re == 0.0) {
    return std::fabs(z.im);
  }
  if (z.im == 0.0) {
    return std::fabs(z.re);
  }
  return std::sqrt(square_sum(z.re, z.im));
}

// Argument of a non-zero z: atan(b / a), moved by +-pi (the sign of b) when a < 0.
double complex_argument(double a, double b) noexcept {
  if (a > 0.0) {
    return std::atan(b / a);
  }
  if (a < 0.0) {
    return std::atan(b / a) + (b >= 0.0 ? kPi : -kPi);
  }
  return b < 0.0 ? -(kPi / 2.0) : kPi / 2.0;
}

// IMABS: scaled by the larger part, `|a| * sqrt(1 + (b/a)^2)`.
double cplx_abs(Complex z) {
  const double a = std::fabs(z.re);
  const double b = std::fabs(z.im);
  if (a == 0.0) {
    return b;
  }
  if (b == 0.0) {
    return a;
  }
  if (a >= b) {
    const double r = b / a;
    return a * std::sqrt(1.0 + r * r);
  }
  const double r = a / b;
  return b * std::sqrt(1.0 + r * r);
}

double cplx_real(Complex z) {
  return z.re;
}

double cplx_imaginary(Complex z) {
  return z.im;
}

// The suffix of a binary result: a pure-real operand takes the other
// operand's suffix; two imaginary operands must agree (false otherwise).
bool reconcile_suffix(Complex a, Complex b, char* out) {
  if (a.im != 0.0 && b.im != 0.0 && a.suffix != b.suffix) {
    return false;
  }
  *out = a.im != 0.0 ? a.suffix : b.suffix;
  return true;
}

Value Complex_fn(const Value* args, std::uint32_t arity, Arena& arena) {
  auto real_imaginary = builtins_detail::read_number_pair(args, 0, 1);
  if (!real_imaginary) {
    return Value::error(real_imaginary.error());
  }
  const auto& pair = real_imaginary.value();
  char suffix = 'i';
  if (arity >= 3) {
    // Suffix argument: must be Text "i" or "j"; Blank defaults to "i";
    // anything else is #VALUE!.
    const Value& s = args[2];
    if (s.is_blank()) {
      suffix = 'i';
    } else if (s.is_text()) {
      const std::string_view sv = s.as_text();
      if (sv.empty() || sv == "i") {
        // Excel treats "" (including the result of `=""`) the same as
        // omitted / blank — both default the suffix to "i".
        suffix = 'i';
      } else if (sv == "j") {
        suffix = 'j';
      } else {
        return Value::error(ErrorCode::Value);
      }
    } else {
      return Value::error(ErrorCode::Value);
    }
  }
  return Value::text(arena.intern(format_complex(pair.first, pair.second, suffix)));
}

Value ImAbs(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return im_unary_number(&cplx_abs, args);
}

Value ImReal(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return im_unary_number(&cplx_real, args);
}

Value ImAginary(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return im_unary_number(&cplx_imaginary, args);
}

Value ImConjugate(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  Complex r = z.value();
  r.im = -r.im;
  return text_complex(r, arena);
}

Value ImArgument(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.re == 0.0 && c.im == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  return Value::number(complex_argument(c.re, c.im));
}

// ---------------------------------------------------------------------------
// Arithmetic
// ---------------------------------------------------------------------------

// Left fold over every argument, range and array cells flattened in order;
// operands with different imaginary suffixes are #VALUE!. The result's suffix
// is reconciled across the operands. `product` selects IMPRODUCT (an infinite
// component is #VALUE!) over IMSUM.
Value fold_complex(const Value* args, std::uint32_t arity, Arena& arena, bool product) {
  // A reference with no populated cells folds like a single blank cell.
  if (arity == 0U) {
    const Value blank = Value::blank();
    return fold_complex(&blank, 1U, arena, product);
  }
  auto first = parse_complex_value(args[0]);
  if (!first) {
    return Value::error(first.error());
  }
  Complex acc = first.value();
  for (std::uint32_t i = 1; i < arity; ++i) {
    auto nxt = parse_complex_value(args[i]);
    if (!nxt) {
      return Value::error(nxt.error());
    }
    char suffix = acc.suffix;
    if (!reconcile_suffix(acc, nxt.value(), &suffix)) {
      return Value::error(ErrorCode::Value);
    }
    const Complex n = nxt.value();
    if (product) {
      acc = Complex{acc.re * n.re - acc.im * n.im, acc.re * n.im + acc.im * n.re, suffix};
      if (std::isinf(acc.re) || std::isinf(acc.im) || std::isnan(acc.re) || std::isnan(acc.im)) {
        return Value::error(ErrorCode::Value);
      }
    } else {
      acc = Complex{acc.re + n.re, acc.im + n.im, suffix};
    }
  }
  return text_complex(acc, arena);
}

Value ImSum(const Value* args, std::uint32_t arity, Arena& arena) {
  return fold_complex(args, arity, arena, /*product=*/false);
}

// Parses the two operands of a binary complex function and reconciles their
// suffix into `*suffix`; a mismatch is #NUM!.
Expected<std::pair<Complex, Complex>, ErrorCode> parse_complex_pair(const Value* args, char* suffix) {
  auto a = parse_complex_value(args[0]);
  if (!a) {
    return std::move(a.error());
  }
  auto b = parse_complex_value(args[1]);
  if (!b) {
    return std::move(b.error());
  }
  if (!reconcile_suffix(a.value(), b.value(), suffix)) {
    return ErrorCode::Num;
  }
  return std::pair<Complex, Complex>{a.value(), b.value()};
}

Value ImSub(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  char suffix;
  auto ab = parse_complex_pair(args, &suffix);
  if (!ab) {
    return Value::error(ab.error());
  }
  const Complex& a = ab.value().first;
  const Complex& b = ab.value().second;
  return text_complex(Complex{a.re - b.re, a.im - b.im, suffix}, arena);
}

Value ImProduct(const Value* args, std::uint32_t arity, Arena& arena) {
  return fold_complex(args, arity, arena, /*product=*/true);
}

// IMDIV: Smith's scaled division, scaled by the larger part of the divisor.
Value ImDiv(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  char suffix;
  auto ab = parse_complex_pair(args, &suffix);
  if (!ab) {
    return Value::error(ab.error());
  }
  const Complex a = ab.value().first;
  const double c = ab.value().second.re;
  const double d = ab.value().second.im;
  if (c == 0.0 && d == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  double re;
  double im;
  if (std::fabs(c) >= std::fabs(d)) {
    const double r = d / c;
    const double den = c + d * r;
    re = (a.re + a.im * r) / den;
    im = (a.im - a.re * r) / den;
  } else {
    const double r = c / d;
    const double den = c * r + d;
    re = (a.re * r + a.im) / den;
    im = (a.im * r - a.re) / den;
  }
  return text_complex(Complex{re, im, suffix}, arena);
}

// IMPOWER: polar form with the magnitude through POWER. A zero magnitude
// (including an underflowed square sum) is 0 for a positive exponent and
// #NUM! otherwise.
Value ImPower(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  auto n = coerce_to_number(args[1]);
  if (!n) {
    return Value::error(n.error());
  }
  const Complex c = z.value();
  const double exponent = n.value();
  if (c.re == 0.0 && c.im == 0.0) {
    if (exponent <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
    return text_complex(Complex{0.0, 0.0, c.suffix}, arena);
  }
  const double r = naive_abs(c);
  if (std::isnan(r)) {
    return Value::error(ErrorCode::Num);
  }
  const double theta = complex_argument(c.re, c.im);
  if (r == 0.0) {
    if (exponent <= 0.0) {
      return Value::error(ErrorCode::Num);
    }
    return text_complex(Complex{0.0, 0.0, c.suffix}, arena);
  }
  auto power = apply_pow(r, exponent);
  if (!power || std::isnan(power.value()) || std::isinf(power.value())) {
    return Value::error(ErrorCode::Num);
  }
  const double mag = power.value();
  const double ang = exponent * theta;
  return text_complex(Complex{mag * std::cos(ang), mag * std::sin(ang), c.suffix}, arena);
}

// ---------------------------------------------------------------------------
// Exponentials / logarithms / roots
// ---------------------------------------------------------------------------

// IMEXP: a pure-real argument that overflows is #NUM!; with an imaginary part
// the exponential clamps at the largest double. sin / cos are capped.
Value ImExp(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.im == 0.0) {
    return text_complex(Complex{std::exp(c.re), 0.0, c.suffix}, arena);
  }
  if (!builtins_detail::trig_argument_in_range(c.im)) {
    return Value::error(ErrorCode::Num);
  }
  double mag = std::exp(c.re);
  if (std::isinf(mag)) {
    mag = std::numeric_limits<double>::max();
  }
  return text_complex(Complex{mag * std::cos(c.im), mag * std::sin(c.im), c.suffix}, arena);
}

// Shared kernel for IMLN / IMLOG10 / IMLOG2: ln|z| scaled by `inv_ln_base`
// (IMLN passes 1) and the argument likewise. ln|z| is `log(a*a + b*b) / 2`
// with flushed squares (#NUM! when the sum overflows or underflows to zero);
// a pure part uses `log` of that part directly. z == 0 is #NUM!.
[[gnu::noinline]] Value im_log_base(double inv_ln_base, const Value* args, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.re == 0.0 && c.im == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  double ln_r;
  if (c.re == 0.0 || c.im == 0.0) {
    ln_r = std::log(c.re == 0.0 ? std::fabs(c.im) : std::fabs(c.re));
  } else {
    const double sum = square_sum(c.re, c.im);
    if (std::isnan(sum) || sum == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    ln_r = std::log(sum) / 2.0;
  }
  const double theta = complex_argument(c.re, c.im);
  return text_complex(Complex{ln_r * inv_ln_base, theta * inv_ln_base, c.suffix}, arena);
}

Value ImLn(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_log_base(1.0, args, arena);
}

Value ImLog10(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_log_base(1.0 / std::log(10.0), args, arena);
}

Value ImLog2(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_log_base(1.0 / std::log(2.0), args, arena);
}

// IMSQRT: the naive magnitude sqrt(a*a + b*b) (no pure-part shortcut), then the polar form.
Value ImSqrt(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.re == 0.0 && c.im == 0.0) {
    return text_complex(Complex{0.0, 0.0, c.suffix}, arena);
  }
  const double sum = square_sum(c.re, c.im);
  if (std::isnan(sum)) {
    return Value::error(ErrorCode::Num);
  }
  const double sqrt_r = std::sqrt(std::sqrt(sum));
  const double theta = complex_argument(c.re, c.im);
  return text_complex(Complex{sqrt_r * std::cos(theta / 2.0), sqrt_r * std::sin(theta / 2.0), c.suffix}, arena);
}

// ---------------------------------------------------------------------------
// Trigonometric
// ---------------------------------------------------------------------------

// 2 * cosh(y) and -2 * sinh(y), each with its own large-|y| and small-|y| form.
double two_cosh(double y) noexcept {
  const double ay = std::fabs(y);
  if (ay > 20.0) {
    return std::exp(ay);
  }
  if (ay > 0.0001) {
    const double e = std::exp(y);
    return (1.0 / e) + e;
  }
  const double y2 = y * y;
  return ((y2 * y2) / 12.0) + (y2 + 2.0);
}

double neg_two_sinh(double y) noexcept {
  const double ay = std::fabs(y);
  if (ay > 20.0) {
    const double e = std::exp(ay);
    return y > 0.0 ? -e : e;
  }
  if (ay > 0.01) {
    const double e = std::exp(y);
    return (1.0 / e) - e;
  }
  const double ny = -y;
  const double y2 = ny * ny;
  double sum = ny * 2.0;
  double term = (y2 * ny) / 3.0;
  sum = term + sum;
  term = (term * y2) / 20.0;
  return term + sum;
}

// IMCSC (sin / cos) and IMSEC (cos / sin, with the imaginary sign flipped):
// 1 / sin(z) and 1 / cos(z) from the real and imaginary parts directly.
//
//   D    = (2 f1(x))^2 + (2cosh(2y) - 2)
//   real = (2cosh(y) / D) * (2 f1(x))
//   imag = (-2sinh(y) / D) * (s * 2 f2(x))
//
// with f1, f2, s = sin, cos, +1 for IMCSC and cos, sin, -1 for IMSEC. Beyond
// |y| ~ 354 the squares overflow, so the quotient form
// (2 f1(x)) / 2cosh(y), (s 2 f2(x)) / -2sinh(y) is used, and past |y| ~ 709
// the result is 0.
Value im_csc_sec(const Value* args, Arena& arena, bool sec) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  // Far from the real axis the result is 0 before any sin / cos is taken.
  if (c.im > 709.0 || c.im < -708.0) {
    return text_complex(Complex{0.0, 0.0, c.suffix}, arena);
  }
  if (!builtins_detail::trig_argument_in_range(c.re)) {
    return Value::error(ErrorCode::Num);
  }
  const double f1 = sec ? std::cos(c.re) : std::sin(c.re);
  const double f2 = sec ? std::sin(c.re) : std::cos(c.re);
  const double sign2 = sec ? -2.0 : 2.0;
  if (c.im == 0.0) {
    if (f1 == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    return text_complex(Complex{1.0 / f1, 0.0, c.suffix}, arena);
  }
  if (c.im > 354.5 || c.im < -354.0) {
    const double a = f1 * 2.0;
    const double b = f2 * sign2;
    return text_complex(Complex{a / two_cosh(c.im), b / neg_two_sinh(c.im), c.suffix}, arena);
  }
  const double sinh_term = neg_two_sinh(c.im);
  const double cosh_term = two_cosh(c.im);
  const double a = f1 * 2.0;
  const double d = a * a + (two_cosh(c.im * 2.0) - 2.0);
  if (d == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double real = (cosh_term / d) * a;
  const double imag = (sinh_term / d) * (f2 * sign2);
  return text_complex(Complex{real, imag, c.suffix}, arena);
}

Value ImSec(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_csc_sec(args, arena, /*sec=*/true);
}

Value ImCsc(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_csc_sec(args, arena, /*sec=*/false);
}

// ---------------------------------------------------------------------------
// Trigonometric and hyperbolic kernels
// ---------------------------------------------------------------------------
//
// Each is evaluated from the real and imaginary parts with the 2cosh / -2sinh
// helpers above, so the large-argument branches (|y| or |x| beyond ~354, ~709)
// and the order of the products follow one scheme.

// exp(y) bounded to the range whose reciprocal is also finite; NaN outside it.
double bounded_exp(double y) noexcept {
  if (y > 709.782712893384 || y < -708.3964185322641) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  return std::exp(y);
}

// IMSIN (sin x cosh y + i cos x sinh y) and IMCOS (cos x cosh y - i sin x sinh y).
Value im_sin_cos(const Value* args, Arena& arena, bool cos) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (!builtins_detail::trig_argument_in_range(c.re)) {
    return Value::error(ErrorCode::Num);
  }
  const double e = bounded_exp(c.im);
  if (!(e > 0.0) || !(e < std::numeric_limits<double>::max())) {
    return Value::error(ErrorCode::Num);
  }
  const double inv = 1.0 / e;
  const double real_factor = cos ? std::cos(c.re) : std::sin(c.re);
  const double imag_factor = cos ? -std::sin(c.re) : std::cos(c.re);
  const double imag = ((e - inv) * imag_factor) * 0.5;
  const double real = ((e + inv) * real_factor) * 0.5;
  return text_complex(Complex{real, imag, c.suffix}, arena);
}

// IMTAN (cot == false) and IMCOT: sin / cos and cos / sin of z.
//
//   D = (2cosh(2y) + 2) - 4 t1^2,  t1 = sin x (tan) | cos x (cot)
//   real = (t2 * (4 t1)) / D,  t2 = cos x | sin x
//   imag = (2sinh(2y) | -2sinh(2y)) / D
//
// A real argument is tan x or 1 / tan x; |y| past ~354 saturates to +-i.
Value im_tan_cot(const Value* args, Arena& arena, bool cot) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.im == 0.0) {
    if (!builtins_detail::trig_argument_in_range(c.re)) {
      return Value::error(ErrorCode::Num);
    }
    const double t = std::tan(c.re);
    return text_complex(Complex{cot ? 1.0 / t : t, 0.0, c.suffix}, arena);
  }
  if (c.im > 354.5 || c.im < -354.0) {
    const double direction = (c.im > 354.5 ? 1.0 : -1.0) * (cot ? -1.0 : 1.0);
    return text_complex(Complex{0.0, direction, c.suffix}, arena);
  }
  if (!builtins_detail::trig_argument_in_range(c.re)) {
    return Value::error(ErrorCode::Num);
  }
  const double y2 = c.im * 2.0;
  const double sinh_term = neg_two_sinh(y2);
  const double cosh_term = two_cosh(y2);
  const double t1 = cot ? std::cos(c.re) : std::sin(c.re);
  const double a = t1 * 4.0;
  const double d = (cosh_term + 2.0) - t1 * a;
  const double t2 = cot ? std::sin(c.re) : std::cos(c.re);
  const double real = t2 * a;
  const double imag = cot ? sinh_term : -sinh_term;
  return text_complex(Complex{real / d, imag / d, c.suffix}, arena);
}

// IMSINH (cosh == false) and IMCOSH: sinh x cos y + i cosh x sin y and
// cosh x cos y + i sinh x sin y. Past |x| ~ 709 the products are formed
// through logarithms; a negative factor there is #NUM!.
Value im_sinh_cosh(const Value* args, Arena& arena, bool cosh) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (!builtins_detail::trig_argument_in_range(c.im)) {
    return Value::error(ErrorCode::Num);
  }
  const double neg_re = -c.re;
  if (c.im == 0.0) {
    const double v = cosh ? two_cosh(neg_re) : neg_two_sinh(neg_re);
    return text_complex(Complex{v / 2.0, 0.0, c.suffix}, arena);
  }
  if (c.re > 709.0 || c.re < -708.0) {
    const double magnitude = std::fabs(c.re);
    const double a = std::cos(c.im) / 2.0;
    const double b = std::sin(c.im) / 2.0;
    if (!(a > 0.0) || !(b > 0.0)) {
      return Value::error(ErrorCode::Num);
    }
    double real = std::exp(std::log(a) + magnitude);
    double imag = std::exp(std::log(b) + magnitude);
    if (c.re < 0.0) {
      if (cosh) {
        imag = -imag;
      } else {
        real = -real;
      }
    }
    return text_complex(Complex{real, imag, c.suffix}, arena);
  }
  const double sinh_term = neg_two_sinh(neg_re);
  const double cosh_term = two_cosh(neg_re);
  const double real = (std::cos(c.im) / 2.0) * (cosh ? cosh_term : sinh_term);
  const double imag = (std::sin(c.im) / 2.0) * (cosh ? sinh_term : cosh_term);
  return text_complex(Complex{real, imag, c.suffix}, arena);
}

// IMCSCH (sech == false) and IMSECH: 1 / sinh z and 1 / cosh z.
//
//   D = (2sinh x)^2 + (-2 sin y)^2 | (2cos y)^2
//   real = (2cos y) * (2sinh x | 2cosh x) / D
//   imag = (-2 sin y) * (2cosh x | 2sinh x) / D
//
// For 354.5 < |x| <= 709 the quotient form is used, with the sinh helper
// applied to x un-negated (so IMCSCH there carries the opposite sign).
Value im_csch_sech(const Value* args, Arena& arena, bool sech) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.im != 0.0 && (c.re > 709.0 || c.re < -708.0)) {
    return text_complex(Complex{0.0, 0.0, c.suffix}, arena);
  }
  if (!builtins_detail::trig_argument_in_range(c.im)) {
    return Value::error(ErrorCode::Num);
  }
  const double neg_re = -c.re;
  if (c.im == 0.0) {
    const double v = sech ? two_cosh(neg_re) : neg_two_sinh(neg_re);
    return text_complex(Complex{2.0 / v, 0.0, c.suffix}, arena);
  }
  if (c.re > 354.5 || c.re < -354.0) {
    const double sinh_term = neg_two_sinh(c.re);
    const double cosh_term = two_cosh(c.re);
    const double real = (std::cos(c.im) * 2.0) / (sech ? cosh_term : sinh_term);
    const double imag = (std::sin(c.im) * 2.0) / (sech ? sinh_term : cosh_term);
    return text_complex(Complex{real, imag, c.suffix}, arena);
  }
  const double sinh_term = neg_two_sinh(neg_re);
  const double cosh_term = two_cosh(neg_re);
  const double a = std::cos(c.im) * 2.0;
  const double b = std::sin(c.im) * -2.0;
  const double real = a * (sech ? cosh_term : sinh_term);
  const double imag = b * (sech ? sinh_term : cosh_term);
  const double d = sinh_term * sinh_term + (sech ? a * a : b * b);
  return text_complex(Complex{real / d, imag / d, c.suffix}, arena);
}

Value ImSin(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_sin_cos(args, arena, /*cos=*/false);
}

Value ImCos(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_sin_cos(args, arena, /*cos=*/true);
}

Value ImTan(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_tan_cot(args, arena, /*cot=*/false);
}

Value ImCot(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_tan_cot(args, arena, /*cot=*/true);
}

Value ImSinh(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_sinh_cosh(args, arena, /*cosh=*/false);
}

Value ImCosh(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_sinh_cosh(args, arena, /*cosh=*/true);
}

Value ImSech(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_csch_sech(args, arena, /*sech=*/true);
}

Value ImCsch(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_csch_sech(args, arena, /*sech=*/false);
}

}  // namespace

void register_complex_num_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      builtins_detail::analysis_toolpak({"COMPLEX", 2u, 3u, &Complex_fn}),
      builtins_detail::analysis_toolpak({"IMABS", 1u, 1u, &ImAbs}),
      builtins_detail::analysis_toolpak({"IMAGINARY", 1u, 1u, &ImAginary}),
      builtins_detail::analysis_toolpak({"IMREAL", 1u, 1u, &ImReal}),
      builtins_detail::analysis_toolpak({"IMCONJUGATE", 1u, 1u, &ImConjugate}),
      builtins_detail::analysis_toolpak({"IMARGUMENT", 1u, 1u, &ImArgument}),
      builtins_detail::analysis_toolpak({"IMSUM", 1u, kVariadic, &ImSum, true, true}),
      builtins_detail::analysis_toolpak({"IMSUB", 2u, 2u, &ImSub}),
      builtins_detail::analysis_toolpak({"IMPRODUCT", 1u, kVariadic, &ImProduct, true, true}),
      builtins_detail::analysis_toolpak({"IMDIV", 2u, 2u, &ImDiv}),
      builtins_detail::analysis_toolpak({"IMPOWER", 2u, 2u, &ImPower}),
      builtins_detail::analysis_toolpak({"IMEXP", 1u, 1u, &ImExp}),
      builtins_detail::analysis_toolpak({"IMLN", 1u, 1u, &ImLn}),
      builtins_detail::analysis_toolpak({"IMLOG10", 1u, 1u, &ImLog10}),
      builtins_detail::analysis_toolpak({"IMLOG2", 1u, 1u, &ImLog2}),
      builtins_detail::analysis_toolpak({"IMSQRT", 1u, 1u, &ImSqrt}),
      builtins_detail::analysis_toolpak({"IMSIN", 1u, 1u, &ImSin}),
      builtins_detail::analysis_toolpak({"IMCOS", 1u, 1u, &ImCos}),
      builtins_detail::analysis_toolpak({"IMTAN", 1u, 1u, &ImTan}),
      builtins_detail::analysis_toolpak({"IMSEC", 1u, 1u, &ImSec}),
      builtins_detail::analysis_toolpak({"IMCSC", 1u, 1u, &ImCsc}),
      builtins_detail::analysis_toolpak({"IMCOT", 1u, 1u, &ImCot}),
      builtins_detail::analysis_toolpak({"IMSINH", 1u, 1u, &ImSinh}),
      builtins_detail::analysis_toolpak({"IMCOSH", 1u, 1u, &ImCosh}),
      builtins_detail::analysis_toolpak({"IMSECH", 1u, 1u, &ImSech}),
      builtins_detail::analysis_toolpak({"IMCSCH", 1u, 1u, &ImCsch}),
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
