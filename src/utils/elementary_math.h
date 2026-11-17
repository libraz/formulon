//
// Elementary functions computed with IEEE double and 64-bit integer arithmetic only, so a result is the same bits
// on every supported target and compiler. Call these instead of the <cmath> functions of the same name wherever a
// result can reach a cell; tools/dev/check_source_seams.py enforces that.
//
//   sin(x), cos(x)   odd and even polynomial kernels on x reduced to [-pi/4, pi/4].
//   tan(x)           their quotient on the same reduction.

#ifndef FORMULON_UTILS_ELEMENTARY_MATH_H_
#define FORMULON_UTILS_ELEMENTARY_MATH_H_

namespace formulon {
namespace elementary {

/// Sine, cosine and tangent of x radians. NaN and infinities give NaN.
double sin(double x) noexcept;
double cos(double x) noexcept;
double tan(double x) noexcept;

}  // namespace elementary
}  // namespace formulon

#endif  // FORMULON_UTILS_ELEMENTARY_MATH_H_
