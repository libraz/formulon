//
// Implementation of Excel's complex-number built-ins (COMPLEX + 24 IM*).
//
// Complex numbers travel through the Formulon calc engine as `Text` values
// of the form "a+bi" (or "a+bj"). The helpers below:
//
//   * `parse_complex`  parses such text into (real, imag, suffix); accepts
//     Number / Bool / Blank directly for Excel quirk compatibility.
//   * `format_complex` renders (real, imag, suffix) back to a canonical
//     Excel-compatible string. Uses `locale_number_text` (the same shortest-
//     round-trip formatter as the text coerce) so output is stable across
//     the engine.
//   * Individual IM* impls reuse those two helpers and a small `CplxOp`
//     set (multiplication, reciprocal, exp, ln) for arithmetic.
//
// Suffix propagation: binary / variadic ops take the suffix from the first
// argument; mixing `i` and `j` is rejected with `#VALUE!`, matching Excel.

#include "eval/builtins/complex_num.h"

#include <cmath>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "eval/locale_text.h"
#include "excel_locale.h"
#include "utils/arena.h"
#include "utils/double_parse.h"
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
// that the entire view be consumed (no trailing characters) and the
// parsed value to be finite (NaN / Inf are rejected so the IM* impls
// can treat success as "usable as a number"). Delegates the raw parse
// to the shared decimal parser in `utils/double_parse.h`.
inline bool parse_double(std::string_view s, double* out) {
  double v = 0.0;
  if (!parse_double_exact(s, &v)) {
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
std::optional<Complex> parse_complex_text(std::string_view raw) {
  if (raw.empty()) {
    return std::nullopt;
  }
  // The text carries the locale decimal separator; the numeric scan below
  // works on the invariant `.` form.
  const char decimal = locale_facts(current_eval_profile()).decimal_separator;
  std::string invariant(raw);
  // Reject any forbidden character up-front so we never fall into surprise
  // parse behaviours (e.g. the leading whitespace the decimal parser skips).
  for (char& c : invariant) {
    const bool ok =
        (c >= '0' && c <= '9') || c == '+' || c == '-' || c == decimal || c == 'e' || c == 'E' || c == 'i' || c == 'j';
    if (!ok) {
      return std::nullopt;
    }
    if (c == decimal) {
      c = '.';
    }
  }
  const std::string_view src = invariant;
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
    out += locale_number_text(re);
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
    out += locale_number_text(im);
    out.push_back(suffix);
    return out;
  }

  // Full form.
  std::string out;
  out += locale_number_text(re);
  if (im > 0.0) {
    out.push_back('+');
  }
  // im < 0 -> the number text already carries the leading '-'.
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
  out += locale_number_text(im);
  out.push_back(suffix);
  return out;
}

// Interns and returns a Text value for the complex triple. We deliberately
// do not snap tiny-but-nonzero components to zero: Mac Excel 365 surfaces
// the IEEE residue from polar-form round-trips (e.g. IMSQRT("-1") renders
// as "6.12E-17+i" rather than "i"), and 1-bit parity requires us to do the
// same. Algebraic IM* impls that operate on direct real/imag arithmetic
// (IMDIV, IMPRODUCT, IMSUM, ...) already produce exact zeros without help.
Value text_complex(Complex z, Arena& arena) {
  if (!std::isfinite(z.re) || !std::isfinite(z.im)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::text(arena.intern(format_complex(z.re, z.im, z.suffix)));
}

Value im_unary_text(Complex (*op)(Complex), const Value* args, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return text_complex(op(z.value()), arena);
}

Value im_unary_number(double (*op)(Complex), const Value* args) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  return Value::number(op(z.value()));
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

double cplx_abs(Complex z) {
  return std::hypot(z.re, z.im);
}

double cplx_real(Complex z) {
  return z.re;
}

double cplx_imaginary(Complex z) {
  return z.im;
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
  return Value::text(arena.intern(format_complex(re.value(), im.value(), suffix)));
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
  return Value::number(std::atan2(c.im, c.re));
}

// ---------------------------------------------------------------------------
// Arithmetic
// ---------------------------------------------------------------------------

// Left fold of `combine` over every argument; the result's suffix is the
// one reconciled across all operands.
Value fold_complex(const Value* args, std::uint32_t arity, Arena& arena, Complex (*combine)(Complex, Complex)) {
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
    acc = combine(acc, nxt.value());
    acc.suffix = suffix;
  }
  return text_complex(acc, arena);
}

Complex add_complex(Complex a, Complex b) {
  return Complex{a.re + b.re, a.im + b.im, a.suffix};
}

Value ImSum(const Value* args, std::uint32_t arity, Arena& arena) {
  return fold_complex(args, arity, arena, add_complex);
}

// Parses the two operands of a binary complex function and reconciles their
// suffix into `*suffix`.
Expected<std::pair<Complex, Complex>, ErrorCode> parse_complex_pair(const Value* args, char* suffix) {
  auto a = parse_complex_value(args[0]);
  if (!a) {
    return std::move(a.error());
  }
  auto b = parse_complex_value(args[1]);
  if (!b) {
    return std::move(b.error());
  }
  *suffix = 'i';
  if (!reconcile_suffix(a.value(), b.value(), suffix)) {
    return ErrorCode::Value;
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
  return fold_complex(args, arity, arena, cplx_mul);
}

Value ImDiv(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  char suffix;
  auto ab = parse_complex_pair(args, &suffix);
  if (!ab) {
    return Value::error(ab.error());
  }
  const Complex z2 = ab.value().second;
  const double denom = z2.re * z2.re + z2.im * z2.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const Complex z1 = ab.value().first;
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
  return im_unary_text(&cplx_exp, args, arena);
}

// Shared kernel for IMLN / IMLOG10 / IMLOG2: computes
// log_base(z) = ln(z) * inv_ln_base (IMLN passes 1, which scales exactly). Returns #NUM! when z == 0; otherwise
// scales the real and imaginary components of ln(z) by `inv_ln_base`. The
// caller passes the precomputed `1 / ln(base)` constant so we keep one body.
// Marked `noinline` so the body is emitted exactly once — the wrappers are
// short enough that O3+LTO would otherwise inline a separate copy per call
// site, defeating the size-sharing this kernel exists to provide.
[[gnu::noinline]] Value im_log_base(double inv_ln_base, const Value* args, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex c = z.value();
  if (c.re == 0.0 && c.im == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const Complex ln_z = cplx_ln(c);
  return text_complex(Complex{ln_z.re * inv_ln_base, ln_z.im * inv_ln_base, c.suffix}, arena);
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
  return im_unary_text(&cplx_sin, args, arena);
}

Value ImCos(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_unary_text(&cplx_cos, args, arena);
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

// Shared kernel for IMSEC / IMCSC / IMSECH / IMCSCH: computes 1 / op(z),
// where `op` is one of cplx_cos / cplx_sin / cplx_cosh / cplx_sinh. Returns
// #NUM! when |op(z)| == 0 (avoids divide-by-zero). Marked `noinline` so the
// body is emitted once and shared across all four wrappers — without it,
// O3+LTO would inline a separate specialization per call site and undo the
// deduplication this kernel exists to provide.
[[gnu::noinline]] Value im_unary_recip(Complex (*op)(Complex), const Value* args, Arena& arena) {
  auto z = parse_complex_value(args[0]);
  if (!z) {
    return Value::error(z.error());
  }
  const Complex w = op(z.value());
  const double denom = w.re * w.re + w.im * w.im;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return text_complex(cplx_recip(w), arena);
}

Value ImSec(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_unary_recip(&cplx_cos, args, arena);
}

Value ImCsc(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_unary_recip(&cplx_sin, args, arena);
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
  return im_unary_text(&cplx_sinh, args, arena);
}

Value ImCosh(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_unary_text(&cplx_cosh, args, arena);
}

Value ImSech(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_unary_recip(&cplx_cosh, args, arena);
}

Value ImCsch(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  return im_unary_recip(&cplx_sinh, args, arena);
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
      builtins_detail::analysis_toolpak({"IMSUM", 1u, kVariadic, &ImSum}),
      builtins_detail::analysis_toolpak({"IMSUB", 2u, 2u, &ImSub}),
      builtins_detail::analysis_toolpak({"IMPRODUCT", 1u, kVariadic, &ImProduct}),
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
