//
// Scalar operations shared by the special-function kernels, each one double
// operation with Excel's subnormal flush on multiply and divide.

#ifndef FORMULON_EVAL_STATS_FLUSHED_OPS_H_
#define FORMULON_EVAL_STATS_FLUSHED_OPS_H_

#include <cmath>
#include <cstdint>
#include <cstring>

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

/// `x^y` as `exp(log(x) * y)`.
inline double power(double x, double y) noexcept {
  return std::exp(mul(std::log(x), y));
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
