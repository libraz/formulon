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

#include "eval/builtins/complex_num_text.h"
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
using complex_detail::Complex;
using complex_detail::format_complex;
using complex_detail::parse_complex_value;
using complex_detail::text_complex;

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
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
