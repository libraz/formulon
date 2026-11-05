//
// Implementation of Excel's complex-number trigonometric and hyperbolic
// built-ins (IMSIN ... IMCSCH).

#include "eval/builtins/complex_num_trig.h"

#include <cmath>
#include <cstdint>
#include <limits>

#include "eval/builtins/complex_num_text.h"
#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using complex_detail::Complex;
using complex_detail::parse_complex_value;
using complex_detail::text_complex;

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

// Past these parts the cosh / sinh squares overflow and the quotient forms apply (asymmetric: Excel's cut-offs).
constexpr bool beyond_square_overflow(double v) noexcept {
  return v > 354.5 || v < -354.0;
}

// Past these parts exp overflows and the hyperbolic-reciprocal results are 0.
constexpr bool beyond_exp_overflow(double v) noexcept {
  return v > 709.0 || v < -708.0;
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
  if (beyond_exp_overflow(c.im)) {
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
  if (beyond_square_overflow(c.im)) {
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
  if (beyond_square_overflow(c.im)) {
    const double direction = (c.im > 0.0 ? 1.0 : -1.0) * (cot ? -1.0 : 1.0);
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
  if (beyond_exp_overflow(c.re)) {
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
  if (c.im != 0.0 && beyond_exp_overflow(c.re)) {
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
  if (beyond_square_overflow(c.re)) {
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

void register_complex_num_trig_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
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
