//
// Elementary functions computed with IEEE double and 64-bit integer arithmetic only, so a result is the same bits
// on every supported target and compiler. Call these instead of the <cmath> functions of the same name wherever a
// result can reach a cell; tools/dev/check_source_seams.py enforces that.
//
//   sin(x), cos(x)   odd and even polynomial kernels on x reduced to [-pi/4, pi/4].
//   tan(x)           their quotient on the same reduction.
//   atan(x)          polynomials on intervals of |x| up to 2, a series in 1/|x| above.
//   exp(x)           a 2^(j/128) table and a polynomial in the remainder.
//   log(x), log10(x) a reciprocal and logarithm table and a polynomial in m/c - 1.
//   pow(x, y)        2^(y log2 x) with y log2 x formed as a double-double.

#ifndef FORMULON_UTILS_ELEMENTARY_MATH_H_
#define FORMULON_UTILS_ELEMENTARY_MATH_H_

namespace formulon {
namespace elementary {

/// Sine, cosine and tangent of x radians. NaN and infinities give NaN.
double sin(double x) noexcept;
double cos(double x) noexcept;
double tan(double x) noexcept;

/// Arctangent of x in [-pi/2, pi/2]. NaN gives NaN.
double atan(double x) noexcept;

/// e^x, ln x, log10 x and x^y with the IEEE special cases (ln 0 = -inf, ln of a negative = NaN, ...).
double exp(double x) noexcept;
double log(double x) noexcept;
double log10(double x) noexcept;
double pow(double x, double y) noexcept;

}  // namespace elementary
}  // namespace formulon

#endif  // FORMULON_UTILS_ELEMENTARY_MATH_H_
