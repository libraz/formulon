//
// Excel's near-zero cancellation snap: an addition or subtraction whose
// result is within 2^-50 of its larger operand's magnitude reports exactly 0
// (`=0.5-0.4-0.1` is 0, not -2.78E-17). Excel applies it to the root `+` / `-`
// of a cell formula (element-wise on an array result) and to each step of the
// SUM / AVERAGE / AVERAGEA accumulation, SUBTOTAL and AGGREGATE included; the
// *IF(S) family, D-functions and SUMPRODUCT keep the residue (Mac Excel 365).

#ifndef FORMULON_UTILS_CANCELLATION_SNAP_H_
#define FORMULON_UTILS_CANCELLATION_SNAP_H_

#include <cmath>

namespace formulon {

/// Returns 0 when `sum`, the IEEE result of `a + b` (or `a - b`), satisfies
/// |sum| <= 2^-50 * max(|a|, |b|); otherwise returns `sum` unchanged.
inline double snap_cancellation(double a, double b, double sum) noexcept {
  const double scale = std::fmax(std::fabs(a), std::fabs(b));
  return std::isfinite(sum) && std::fabs(sum) <= std::ldexp(scale, -50) ? 0.0 : sum;
}

/// `a + b` with the cancellation snap applied.
inline double snapped_add(double a, double b) noexcept {
  return snap_cancellation(a, b, a + b);
}

}  // namespace formulon

#endif  // FORMULON_UTILS_CANCELLATION_SNAP_H_
