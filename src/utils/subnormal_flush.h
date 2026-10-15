//
// Excel's multiply and divide report a subnormal result (magnitude below
// DBL_MIN) as exactly 0: `=1E-300*1E-10` is 0, and so is `=EXP(-720)*1`.
// Addition, subtraction, SUM and EXP keep a subnormal (Mac Excel 365).

#ifndef FORMULON_UTILS_SUBNORMAL_FLUSH_H_
#define FORMULON_UTILS_SUBNORMAL_FLUSH_H_

#include <cfloat>
#include <cmath>

namespace formulon {

/// Nonnegative values below this (DBL_MIN / 16) are zero to sign tests,
/// truthiness, unary minus and number-format section selection.
constexpr double kZeroTestThreshold = 0x1p-1026;

/// Excel's zero test is one-sided: `0 <= x < 2^-1026` is zero, while every
/// negative value, however small, is nonzero (`SIGN(0-EXP(-720))` is -1).
inline bool tests_as_zero(double x) noexcept {
  return x >= 0.0 && x < kZeroTestThreshold;
}

/// Returns 0 for a nonzero `r` below DBL_MIN in magnitude; otherwise `r`.
inline double flush_subnormal(double r) noexcept {
  return r != 0.0 && std::fabs(r) < DBL_MIN ? 0.0 : r;
}

}  // namespace formulon

#endif  // FORMULON_UTILS_SUBNORMAL_FLUSH_H_
