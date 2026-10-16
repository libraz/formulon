//
// Elementary-function kernels evaluated in the order Excel evaluates them, for
// callers that must agree with Excel to the last bit.
//
//   excel_log1p(x)  atanh-series log1p: log(x + 1) once |x| >= 0.375, a two-term
//                   expansion below 2^-51, 2 * atanh(x / (x + 2)) in between.
//   excel_expm1(x)  exp(x) - 1 for |x| > 1, x itself below 2^-52, otherwise the
//                   error-compensated ((u - 1) + err) * x / log(u) with u = exp(x).

#ifndef FORMULON_EVAL_EXCEL_MATH_H_
#define FORMULON_EVAL_EXCEL_MATH_H_

#include <cmath>
#include <cstdint>
#include <cstring>

namespace formulon {
namespace eval {

namespace excel_math_detail {

inline std::uint64_t bits_of(double x) noexcept {
  std::uint64_t b;
  std::memcpy(&b, &x, sizeof(b));
  return b;
}

// |u| + sum u^(2j+1) / (2j+1) with the small terms accumulated first, signed like `u`.
inline double atanh_series(double u) noexcept {
  const double sign = u < 0.0 ? -1.0 : 1.0;
  const double a = std::fabs(u);
  const double u2 = u * u;
  constexpr double kEps = 2.220446049250313e-16;
  double k = 1.0;
  double term = a;
  double sum = 0.0;
  while (true) {
    term = ((k / (k + 2.0)) * u2) * term;
    sum = sum + term;
    if (kEps * sum > term) {
      break;
    }
    k = k + 2.0;
  }
  sum = sum + a;
  return sum * sign;
}

}  // namespace excel_math_detail

inline double excel_log1p(double x) noexcept {
  if (x == 0.0) {
    return x;
  }
  if (((excel_math_detail::bits_of(x) >> 51) & 0xfffU) >= 0x7fbU) {
    return std::log(x + 1.0);
  }
  if (std::fabs(x) <= 4.440892098500626e-16) {
    return (x * -0.5 + 1.0) * x;
  }
  return excel_math_detail::atanh_series(x / (x + 2.0)) * 2.0;
}

inline double excel_expm1(double x) noexcept {
  const double ax = std::fabs(x);
  if (ax > 1.0) {
    return std::exp(x) - 1.0;
  }
  if (ax <= 2.220446049250313e-16) {
    return x;
  }
  const double u = std::exp(x);
  const double v = u + -1.0;
  const double e1 = v - u;
  const double w2 = v - e1;
  const double w3 = -1.0 - e1;
  const double c = (u - w2) + w3;
  const double y = v + c;
  return (y * x) / std::log(u);
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_EXCEL_MATH_H_
