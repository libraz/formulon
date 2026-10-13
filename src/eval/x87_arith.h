#ifndef FORMULON_EVAL_X87_ARITH_H_
#define FORMULON_EVAL_X87_ARITH_H_

/// @file
/// Emulation of Windows Excel's x87 evaluation: operations run with a 64-bit
/// significand and the result is rounded to double on store, so each result
/// is round53(round64(exact)). Only the double-rounding tie cases differ from
/// plain IEEE double arithmetic.

namespace formulon::eval {

/// `a + b` rounded to a 64-bit significand and then to double.
double x87_add(double a, double b);

/// `a - b` rounded to a 64-bit significand and then to double.
double x87_sub(double a, double b);

/// `a * b` rounded to a 64-bit significand and then to double.
double x87_mul(double a, double b);

/// `a / b` rounded to a 64-bit significand and then to double. The caller
/// guarantees `b != 0`.
double x87_div(double a, double b);

/// `e^x` evaluated as Windows Excel does: a 64-bit `log2(e)` argument
/// reduction, `2^f - 1`, `+ 1`, then scaling by a power of two.
/// Arguments with `|x| > 709` return `std::exp(x)`.
double x87_exp(double x);

/// Natural logarithm rounded to a 64-bit significand and then to double.
/// Arguments that are not finite and positive return `std::log(x)`.
double x87_ln(double x);

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_X87_ARITH_H_
