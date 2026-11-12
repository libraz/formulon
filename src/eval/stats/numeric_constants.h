// Literal constants shared by the special-function kernels. Each value is the
// exact double the kernels were calibrated with; do not recompute them.

#ifndef FORMULON_EVAL_STATS_NUMERIC_CONSTANTS_H_
#define FORMULON_EVAL_STATS_NUMERIC_CONSTANTS_H_

#include <limits>

namespace formulon {
namespace eval {
namespace stats {

inline constexpr double kEps = 2.220446049250313e-16;
inline constexpr double kLnSqrt2Pi = 0.9189385332046728;
inline constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();
inline constexpr double kInvSqrtPi = 0x1.20dd750429b6dp-1;
inline constexpr double kTinyShape = 0x1p-104;
inline constexpr double kPi = 3.141592653589793;
inline constexpr double kTwoPi = 0x1.921fb54442d18p+2;
inline constexpr double kInvSqrt2 = 0.7071067811865476;
inline constexpr double kInvSqrt2Pi = 0.3989422804014327;

/// Iteration cap of the series and continued-fraction kernels.
// @limit: internal
inline constexpr int kMaxTerms = 100000;
/// Replacement for a zero denominator in a modified-Lentz continued fraction.
inline constexpr double kLentzTiny = 1e-30;
/// Largest degrees-of-freedom value the chi-squared, t and F families accept.
// @limit: internal
inline constexpr double kMaxDegreesOfFreedom = 1.0e10;

/// True when `df` lies outside `[1, kMaxDegreesOfFreedom]`; NaN is not out of range.
inline constexpr bool df_out_of_range(double df) noexcept {
  return df < 1.0 || df > kMaxDegreesOfFreedom;
}

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_NUMERIC_CONSTANTS_H_
