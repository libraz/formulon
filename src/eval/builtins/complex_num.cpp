// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Excel's complex-number built-ins (COMPLEX + 24 IM*).
//
// Complex numbers travel through the Formulon calc engine as `Text` values
// of the form "a+bi" (or "a+bj"). The helpers below:
//
//   * `parse_complex`  parses such text into (real, imag, suffix); accepts
//     Number / Bool / Blank directly for Excel quirk compatibility.
//   * `format_complex` renders (real, imag, suffix) back to a canonical
//     Excel-compatible string. Uses `format_double` (the same shortest-
//     round-trip formatter as SUM's text coerce) so output is stable across
//     the engine.
//   * Individual IM* impls reuse those two helpers and a small `CplxOp`
//     set (multiplication, reciprocal, exp, ln) for arithmetic.
//
// Suffix propagation: binary / variadic ops take the suffix from the first
// argument; mixing `i` and `j` is rejected with `#VALUE!`, matching Excel.

#include "eval/builtins/complex_num.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <optional>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/double_format.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

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

// Parses a decimal double from `s`. Accepts scientific notation; requires
// that the *entire* view be consumed (no trailing characters).
bool parse_double(std::string_view s, double* out) {
  if (s.empty()) {
    return false;
  }
  char stack_buf[64];
  char* heap_buf = nullptr;
  char* buf = stack_buf;
  const std::size_t n = s.size();
  if (n + 1 > sizeof(stack_buf)) {
    heap_buf = static_cast<char*>(std::malloc(n + 1));
    if (heap_buf == nullptr) {
      return false;
    }
    buf = heap_buf;
  }
  std::memcpy(buf, s.data(), n);
  buf[n] = '\0';
  char* end_ptr = nullptr;
  const double v = std::strtod(buf, &end_ptr);
  const bool ok = end_ptr == buf + n;
  if (heap_buf != nullptr) {
    std::free(heap_buf);
  }
  if (!ok) {
    return false;
  }
  if (std::isnan(v) || std::isinf(v)) {
    return false;
  }
  *out = v;
  return true;
}

// Parses a complex-number text view. Returns `std::nullopt` on any
// malformed input (caller should surface #NUM!). The accepted grammar:
//
//   Pure real:       [+-]?<number>
//   Pure imaginary:  [+-]?(<number>)?[ij]   -- empty coef means 1
//   Full:            [+-]?<number>[+-]<number>?[ij]
//
// `i`/`j` are lowercase only; an uppercase `I` or `J` is rejected.
std::optional<Complex> parse_complex_text(std::string_view src) {
  if (src.empty()) {
    return std::nullopt;
  }
  // Reject any forbidden character up-front so we never fall into surprise
  // parse behaviours (e.g. hex floats via strtod's `0x` prefix).
  for (char c : src) {
    const bool ok =
        (c >= '0' && c <= '9') || c == '+' || c == '-' || c == '.' || c == 'e' || c == 'E' || c == 'i' || c == 'j';
    if (!ok) {
      return std::nullopt;
    }
  }
  // Find the suffix (last `i` or `j`, if any). Mixed i/j is rejected.
  bool has_i = false;
  bool has_j = false;
  std::size_t suffix_pos = src.size();
  for (std::size_t i = 0; i < src.size(); ++i) {
    if (src[i] == 'i') {
      has_i = true;
      suffix_pos = i;
    } else if (src[i] == 'j') {
      has_j = true;
      suffix_pos = i;
    }
  }
  if (has_i && has_j) {
    return std::nullopt;
  }
  // The suffix (if any) must be the final character: e.g. "3i" not "3i4".
  if (suffix_pos != src.size() && suffix_pos != src.size() - 1) {
    return std::nullopt;
  }

  // Pure-real path: no suffix.
  if (!has_i && !has_j) {
    double re = 0.0;
    if (!parse_double(src, &re)) {
      return std::nullopt;
    }
    return Complex{re, 0.0, 'i'};
  }

  const char suffix = has_j ? 'j' : 'i';
  // Strip the trailing suffix character.
  std::string_view body = src.substr(0, suffix_pos);

  // Bare-suffix cases: "i" -> +1i, "+i" -> +1i, "-i" -> -1i.
  if (body.empty()) {
    return Complex{0.0, 1.0, suffix};
  }
  if (body == "+") {
    return Complex{0.0, 1.0, suffix};
  }
  if (body == "-") {
    return Complex{0.0, -1.0, suffix};
  }

  // Split `body` into (optional real part) + imaginary coefficient. The
  // imaginary coefficient is the final signed number in `body`. Scientific
  // notation means the split sign must not be part of an exponent, so we
  // scan right-to-left looking for a `+`/`-` that is NOT immediately
  // preceded by `e`/`E`.
  std::size_t split = std::string_view::npos;
  for (std::size_t idx = body.size(); idx-- > 0;) {
    const char c = body[idx];
    if (c == '+' || c == '-') {
      if (idx == 0) {
        // Leading sign belongs to the imag coefficient; no split.
        break;
      }
      const char prev = body[idx - 1];
      if (prev == 'e' || prev == 'E') {
        continue;
      }
      split = idx;
      break;
    }
  }

  if (split == std::string_view::npos) {
    // No split -> entire body is the imaginary coefficient (real part is 0).
    std::string_view coef = body;
    // Bare leading sign cases already handled above; strip a redundant `+`.
    if (coef.size() >= 1 && coef.front() == '+') {
      coef = coef.substr(1);
    }
    double im = 0.0;
    if (!parse_double(coef, &im)) {
      return std::nullopt;
    }
    return Complex{0.0, im, suffix};
  }

  // Full form. `real_part = body[0..split)`, `imag_part = body[split..end)`.
  std::string_view real_part = body.substr(0, split);
  std::string_view imag_part = body.substr(split);

  // Leading `+` on the real part is allowed; strip it.
  if (!real_part.empty() && real_part.front() == '+') {
    real_part = real_part.substr(1);
  }
  double re = 0.0;
  if (!parse_double(real_part, &re)) {
    return std::nullopt;
  }

  // Imaginary part: `[+-]` then either empty (meaning 1) or a number.
  const char sign = imag_part.front();
  std::string_view coef = imag_part.substr(1);
  double mag = 1.0;
  if (!coef.empty()) {
    // Reject a double-sign prefix like "+-4".
    if (coef.front() == '+' || coef.front() == '-') {
      return std::nullopt;
    }
    if (!parse_double(coef, &mag)) {
      return std::nullopt;
    }
  }
  const double im = sign == '-' ? -mag : mag;
  return Complex{re, im, suffix};
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

// Normalises IEEE negative zero to plain zero so that format_double renders
// "0" rather than "-0". This is the same collapse used internally by
// `format_double` but we apply it up-front so that comparisons against 0
// further down remain intuitive.
double normalize_zero(double d) {
  return d == 0.0 ? 0.0 : d;
}

// Formats a complex triple into Excel's canonical string form.
std::string format_complex(double re, double im, char suffix) {
  re = normalize_zero(re);
  im = normalize_zero(im);

  // Pure real.
  if (im == 0.0) {
    std::string out;
    format_double(out, re);
    return out;
  }

  // Pure imaginary.
  if (re == 0.0) {
    std::string out;
    if (im == 1.0) {
      // Bare "+i" is rendered as "i".
      out.push_back(suffix);
      return out;
    }
    if (im == -1.0) {
      out.push_back('-');
      out.push_back(suffix);
      return out;
    }
    format_double(out, im);
    out.push_back(suffix);
    return out;
  }

  // Full form.
  std::string out;
  format_double(out, re);
  if (im > 0.0) {
    out.push_back('+');
  }
  // im < 0 -> format_double already renders the leading '-'.
  if (im == 1.0) {
    // Drop the coefficient but keep the '+' we just appended.
    out.push_back(suffix);
    return out;
  }
  if (im == -1.0) {
    out.push_back('-');
    out.push_back(suffix);
    return out;
  }
  format_double(out, im);
  out.push_back(suffix);
  return out;
}

// Snaps results that are essentially zero to exact zero before formatting.
// Many IM* impls arrive at (tiny_epsilon, something) or (something,
// tiny_epsilon) via polar-form round-trips; without snapping, IMDIV("1+i",
// "1-i") would render as "6.1E-17+i" instead of the expected "i". The
// threshold is keyed off the magnitude of the other component so that we
// never snap a legitimately-small-but-nonzero input.
void snap_zeros(double* re, double* im) {
  // Scale threshold by the magnitude of the non-snapped component plus 1
  // (to handle both very large and very small complex magnitudes). 1e-14
  // is tight enough to leave real numerical differences visible while
  // collapsing accumulated floating-point dust.
  const double scale = std::max({std::fabs(*re), std::fabs(*im), 1.0});
  const double eps = 1e-14 * scale;
  if (std::fabs(*re) < eps) {
    *re = 0.0;
  }
  if (std::fabs(*im) < eps) {
    *im = 0.0;
  }
}

// Interns and returns a Text value for a (possibly-snapped) triple.
Value text_complex(Complex z, Arena& arena) {
  snap_zeros(&z.re, &z.im);
  return Value::text(arena.intern(format_complex(z.re, z.im, z.suffix)));
}

// ---------------------------------------------------------------------------
// Complex arithmetic primitives
// ---------------------------------------------------------------------------

Complex cplx_mul(Complex a, Complex b) {
  return Complex{a.re * b.re - a.im * b.im, a.re * b.im + a.im * b.re, a.suffix};
}

// Returns `1 / z`. Caller must ensure `|z| != 0`.
Complex cplx_recip(Complex z) {
  const double denom = z.re * z.re + z.im * z.im;
  return Complex{z.re / denom, -z.im / denom, z.suffix};
}

// e^z = e^a * (cos b + i sin b).
Complex cplx_exp(Complex z) {
  const double mag = std::exp(z.re);
  return Complex{mag * std::cos(z.im), mag * std::sin(z.im), z.suffix};
}

// ln(z) = ln(r) + i*theta; caller must ensure z != 0.
Complex cplx_ln(Complex z) {
  const double r = std::hypot(z.re, z.im);
  const double theta = std::atan2(z.im, z.re);
  return Complex{std::log(r), theta, z.suffix};
}

// Reconciles the suffix on a binary op: both must agree when both inputs
// contained an explicit imaginary part; a pure-real argument inherits the
// other's suffix. Returns false if the two disagree (mixed i/j).
bool reconcile_suffix(Complex a, Complex b, char* out) {
  // A Complex parsed from a pure-real input defaults to 'i'; there is no
  // way to tell that apart from an explicit 'i'. Excel's observable rule
  // is simpler than "explicit vs implicit": if both are non-zero-imag, the
  // suffixes must match; otherwise the first operand's suffix wins.
  if (a.im != 0.0 && b.im != 0.0 && a.suffix != b.suffix) {
    return false;
  }
  *out = a.suffix;
  return true;
}

// ---------------------------------------------------------------------------
// Constructor / inspectors
// ---------------------------------------------------------------------------

Value Complex_fn(const Value* args, std::uint32_t arity, Arena& arena) {
  auto re = coerce_to_number(args[0]);
  if (!re) {
    return Value::error(re.error());
  }
  auto im = coerce_to_number(args[1]);
  if (!im) {
    return Value::error(im.error());
  }
  char suffix = 'i';
  if (arity >= 3) {
    // Suffix argument: must be Text "i" or "j"; Blank defaults to "i";
    // anything else is #VALUE!.
    const Value& s = args[2];
    if (s.is_blank()) {
      suffix = 'i';
    } else if (s.is_text()) {
      const std::string_view sv = s.as_text();
      if (sv == "i") {
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
  return Value::text(arena.intern(format_complex(re.value(), im.value(), suffix)));
}

Value ImAbs(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return Value::number(std::hypot(z.value().re, z.value().im));
}

Value ImReal(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return Value::number(z.value().re);
}

Value ImAginary(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return Value::number(z.value().im);
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
  return Value::number(std::atan2(c.im, c.re));
}

// ---------------------------------------------------------------------------
// Arithmetic
// ---------------------------------------------------------------------------

Value ImSum(const Value* args, std::uint32_t arity, Arena& arena) {
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
    acc.re += nxt.value().re;
    acc.im += nxt.value().im;
    acc.suffix = suffix;
  }
  return text_complex(acc, arena);
}

Value ImSub(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto a = parse_complex_value(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  auto b = parse_complex_value(args[1]);
  if (!b) {
    return Value::error(b.error());
  }
  char suffix = 'i';
  if (!reconcile_suffix(a.value(), b.value(), &suffix)) {
    return Value::error(ErrorCode::Value);
  }
  return text_complex(Complex{a.value().re - b.value().re, a.value().im - b.value().im, suffix}, arena);
}

Value ImProduct(const Value* args, std::uint32_t arity, Arena& arena) {
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
    acc = cplx_mul(acc, nxt.value());
    acc.suffix = suffix;
  }
  return text_complex(acc, arena);
}

Value ImDiv(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto a = parse_complex_value(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  auto b = parse_complex_value(args[1]);
  if (!b) {
    return Value::error(b.error());
  }
  char suffix = 'i';
  if (!reconcile_suffix(a.value(), b.value(), &suffix)) {
    return Value::error(ErrorCode::Value);
  }
  const Complex z2 = b.value();
  const double denom = z2.re * z2.re + z2.im * z2.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const Complex z1 = a.value();
  const double re = (z1.re * z2.re + z1.im * z2.im) / denom;
  const double im = (z1.im * z2.re - z1.re * z2.im) / denom;
  return text_complex(Complex{re, im, suffix}, arena);
}

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
  const double r = std::hypot(c.re, c.im);
  const double theta = std::atan2(c.im, c.re);
  const double mag = std::pow(r, exponent);
  const double ang = exponent * theta;
  if (std::isnan(mag) || std::isinf(mag)) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(Complex{mag * std::cos(ang), mag * std::sin(ang), c.suffix}, arena);
}

// ---------------------------------------------------------------------------
// Exponentials / logarithms / roots
// ---------------------------------------------------------------------------

Value ImExp(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return text_complex(cplx_exp(z.value()), arena);
}

Value ImLn(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.re == 0.0 && c.im == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(cplx_ln(c), arena);
}

Value ImLog10(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.re == 0.0 && c.im == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const Complex ln_z = cplx_ln(c);
  const double inv = 1.0 / std::log(10.0);
  return text_complex(Complex{ln_z.re * inv, ln_z.im * inv, c.suffix}, arena);
}

Value ImLog2(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.re == 0.0 && c.im == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const Complex ln_z = cplx_ln(c);
  const double inv = 1.0 / std::log(2.0);
  return text_complex(Complex{ln_z.re * inv, ln_z.im * inv, c.suffix}, arena);
}

Value ImSqrt(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.re == 0.0 && c.im == 0.0) {
    return text_complex(Complex{0.0, 0.0, c.suffix}, arena);
  }
  const double r = std::hypot(c.re, c.im);
  const double theta = std::atan2(c.im, c.re);
  const double sqrt_r = std::sqrt(r);
  return text_complex(Complex{sqrt_r * std::cos(theta / 2.0), sqrt_r * std::sin(theta / 2.0), c.suffix}, arena);
}

// ---------------------------------------------------------------------------
// Trigonometric
// ---------------------------------------------------------------------------

Complex cplx_sin(Complex z) {
  return Complex{std::sin(z.re) * std::cosh(z.im), std::cos(z.re) * std::sinh(z.im), z.suffix};
}

Complex cplx_cos(Complex z) {
  return Complex{std::cos(z.re) * std::cosh(z.im), -std::sin(z.re) * std::sinh(z.im), z.suffix};
}

Value ImSin(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return text_complex(cplx_sin(z.value()), arena);
}

Value ImCos(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return text_complex(cplx_cos(z.value()), arena);
}

Value ImTan(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex s = cplx_sin(z.value());
  const Complex c = cplx_cos(z.value());
  const double denom = c.re * c.re + c.im * c.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(cplx_mul(s, cplx_recip(c)), arena);
}

Value ImSec(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = cplx_cos(z.value());
  const double denom = c.re * c.re + c.im * c.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(cplx_recip(c), arena);
}

Value ImCsc(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex s = cplx_sin(z.value());
  const double denom = s.re * s.re + s.im * s.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(cplx_recip(s), arena);
}

Value ImCot(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex s = cplx_sin(z.value());
  const Complex c = cplx_cos(z.value());
  const double denom = s.re * s.re + s.im * s.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(cplx_mul(c, cplx_recip(s)), arena);
}

// ---------------------------------------------------------------------------
// Hyperbolic
// ---------------------------------------------------------------------------

Complex cplx_sinh(Complex z) {
  return Complex{std::sinh(z.re) * std::cos(z.im), std::cosh(z.re) * std::sin(z.im), z.suffix};
}

Complex cplx_cosh(Complex z) {
  return Complex{std::cosh(z.re) * std::cos(z.im), std::sinh(z.re) * std::sin(z.im), z.suffix};
}

Value ImSinh(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return text_complex(cplx_sinh(z.value()), arena);
}

Value ImCosh(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return text_complex(cplx_cosh(z.value()), arena);
}

Value ImSech(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = cplx_cosh(z.value());
  const double denom = c.re * c.re + c.im * c.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(cplx_recip(c), arena);
}

Value ImCsch(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex s = cplx_sinh(z.value());
  const double denom = s.re * s.re + s.im * s.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(cplx_recip(s), arena);
}

}  // namespace

void register_complex_num_builtins(FunctionRegistry& registry) {
  // Constructor / inspectors.
  registry.register_function(FunctionDef{"COMPLEX", 2u, 3u, &Complex_fn});
  registry.register_function(FunctionDef{"IMABS", 1u, 1u, &ImAbs});
  registry.register_function(FunctionDef{"IMAGINARY", 1u, 1u, &ImAginary});
  registry.register_function(FunctionDef{"IMREAL", 1u, 1u, &ImReal});
  registry.register_function(FunctionDef{"IMCONJUGATE", 1u, 1u, &ImConjugate});
  registry.register_function(FunctionDef{"IMARGUMENT", 1u, 1u, &ImArgument});

  // Arithmetic.
  registry.register_function(FunctionDef{"IMSUM", 1u, kVariadic, &ImSum});
  registry.register_function(FunctionDef{"IMSUB", 2u, 2u, &ImSub});
  registry.register_function(FunctionDef{"IMPRODUCT", 1u, kVariadic, &ImProduct});
  registry.register_function(FunctionDef{"IMDIV", 2u, 2u, &ImDiv});
  registry.register_function(FunctionDef{"IMPOWER", 2u, 2u, &ImPower});

  // Exponentials / logs / roots.
  registry.register_function(FunctionDef{"IMEXP", 1u, 1u, &ImExp});
  registry.register_function(FunctionDef{"IMLN", 1u, 1u, &ImLn});
  registry.register_function(FunctionDef{"IMLOG10", 1u, 1u, &ImLog10});
  registry.register_function(FunctionDef{"IMLOG2", 1u, 1u, &ImLog2});
  registry.register_function(FunctionDef{"IMSQRT", 1u, 1u, &ImSqrt});

  // Trigonometric.
  registry.register_function(FunctionDef{"IMSIN", 1u, 1u, &ImSin});
  registry.register_function(FunctionDef{"IMCOS", 1u, 1u, &ImCos});
  registry.register_function(FunctionDef{"IMTAN", 1u, 1u, &ImTan});
  registry.register_function(FunctionDef{"IMSEC", 1u, 1u, &ImSec});
  registry.register_function(FunctionDef{"IMCSC", 1u, 1u, &ImCsc});
  registry.register_function(FunctionDef{"IMCOT", 1u, 1u, &ImCot});

  // Hyperbolic.
  registry.register_function(FunctionDef{"IMSINH", 1u, 1u, &ImSinh});
  registry.register_function(FunctionDef{"IMCOSH", 1u, 1u, &ImCosh});
  registry.register_function(FunctionDef{"IMSECH", 1u, 1u, &ImSech});
  registry.register_function(FunctionDef{"IMCSCH", 1u, 1u, &ImCsch});
}

}  // namespace eval
}  // namespace formulon
