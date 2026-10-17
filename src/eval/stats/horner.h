//
// Horner evaluation shared by the special-function kernels: one multiply and
// one add per coefficient, never fused, with a subnormal product flushed to
// zero as Excel's multiply does.

#ifndef FORMULON_EVAL_STATS_HORNER_H_
#define FORMULON_EVAL_STATS_HORNER_H_

#include <cstddef>

#include "utils/subnormal_flush.h"

namespace formulon {
namespace eval {
namespace stats {

/// Polynomial with `c[0]` the highest-degree coefficient, evaluated at `y`.
template <std::size_t N>
double horner(const double (&c)[N], double y) noexcept {
  double acc = c[0];
  for (std::size_t i = 1; i < N; ++i) {
    acc = flush_subnormal(acc * y);
    acc += c[i];
  }
  return acc;
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_HORNER_H_
