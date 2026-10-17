//
// Scalar operations shared by the special-function kernels, each one double
// operation with Excel's subnormal flush on multiply and divide.

#ifndef FORMULON_EVAL_STATS_FLUSHED_OPS_H_
#define FORMULON_EVAL_STATS_FLUSHED_OPS_H_

#include <cmath>
#include <cstdint>
#include <cstring>
#include <limits>

#include "utils/subnormal_flush.h"

namespace formulon {
namespace eval {
namespace stats {
namespace flushed {

inline double mul(double a, double b) noexcept {
  return flush_subnormal(a * b);
}

inline double div(double a, double b) noexcept {
  return flush_subnormal(a / b);
}

/// Quotient that is 0 when the divisor or the quotient is outside the normal range.
inline double guarded_div(double a, double b) noexcept {
  if (std::fabs(b) < std::numeric_limits<double>::min()) {
    return 0.0;
  }
  const double r = a / b;
  const double magnitude = std::fabs(r);
  if (magnitude > std::numeric_limits<double>::max() || magnitude < std::numeric_limits<double>::min()) {
    return 0.0;
  }
  return r;
}

inline std::uint64_t bits_of(double v) noexcept {
  std::uint64_t b;
  std::memcpy(&b, &v, sizeof b);
  return b;
}

/// True for +0 and tiny positive values (top 16 bits clear).
inline bool top16_clear(double v) noexcept {
  return (bits_of(v) >> 48) == 0;
}

/// Negation that leaves +0 and tiny positive values unchanged.
inline double neg(double v) noexcept {
  return top16_clear(v) ? v : -v;
}

/// `x^y` as `exp(log(x) * y)`; a negative base takes the sign of an odd integral
/// exponent and is NaN for any other exponent, and -0 is NaN.
inline double power(double x, double y) noexcept {
  if (x > 0.0) {
    return std::exp(mul(std::log(x), y));
  }
  if (x == 0.0) {
    return std::signbit(x) ? std::numeric_limits<double>::quiet_NaN() : 0.0;
  }
  if (!(x < 0.0) || y != std::trunc(y)) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  const double r = std::exp(mul(std::log(-x), y));
  const bool odd = std::fabs(std::fmod(y, 2.0)) == 1.0;
  return odd ? -r : r;
}

inline void two_sum(double a, double b, double* s, double* e) noexcept {
  *s = a + b;
  const double bb = *s - a;
  *e = (a - (*s - bb)) + (b - bb);
}

/// Sum of `n >= 2` terms with the rounding errors of each addition accumulated in `err`.
inline void compensated_sum(const double* v, int n, double* sum, double* err) noexcept {
  two_sum(v[0], v[1], sum, err);
  for (int i = 2; i < n; ++i) {
    double e;
    two_sum(*sum, v[i], sum, &e);
    *err += e;
  }
}

}  // namespace flushed
}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_FLUSHED_OPS_H_
