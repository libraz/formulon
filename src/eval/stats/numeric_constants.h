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

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_NUMERIC_CONSTANTS_H_
