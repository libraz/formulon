//
// Excel's near-zero cancellation snap: an addition or subtraction whose
// result is below 8 ulps of its LEFT operand's binade reports exactly 0
// (`=0.5-0.4-0.1` is 0, not -2.78E-17). Excel applies it to the root `+` / `-`
// of a cell formula (element-wise on an array result), to the final addition
// of SUM / AVERAGE / AVERAGEA, SUBTOTAL and AGGREGATE (scaled by the partial
// sum before it), and to every step of a pivot SUM with the addend on the
// left. The *IF(S) family, D-functions and SUMPRODUCT keep the residue
// (Mac Excel 365).

#ifndef FORMULON_UTILS_CANCELLATION_SNAP_H_
#define FORMULON_UTILS_CANCELLATION_SNAP_H_

#include <cmath>

#include "utils/subnormal_flush.h"

namespace formulon {

/// Returns 0 when `result`, the IEEE value of `left + x` (or `left - x`),
/// satisfies |result| < 2^(ilogb(left) - 49); otherwise returns it unchanged.
/// A subnormal threshold is itself flushed, so nothing snaps below
/// |left| = 2^-973.
inline double snap_cancellation(double left, double result) noexcept {
  if (left == 0.0 || !std::isfinite(left) || !std::isfinite(result)) {
    return result;
  }
  return std::fabs(result) < flush_subnormal(std::ldexp(1.0, std::ilogb(left) - 49)) ? 0.0 : result;
}

/// `left + right` with the cancellation snap scaled by `left`.
inline double snapped_add(double left, double right) noexcept {
  return snap_cancellation(left, left + right);
}

/// Left-to-right IEEE sum that snaps only its final addition, scaled by the
/// partial sum before it (SUM / AVERAGE / SUBTOTAL / AGGREGATE).
class FinalStepSnapSum {
 public:
  void add(double x) noexcept {
    before_last_ = total_;
    total_ += x;
  }

  double result() const noexcept { return snap_cancellation(before_last_, total_); }

 private:
  double total_ = 0.0;
  double before_last_ = 0.0;
};

}  // namespace formulon

#endif  // FORMULON_UTILS_CANCELLATION_SNAP_H_
