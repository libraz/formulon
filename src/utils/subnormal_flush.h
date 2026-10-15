//
// Excel's multiply and divide report a subnormal result (magnitude below
// DBL_MIN) as exactly 0: `=1E-300*1E-10` is 0, and so is `=EXP(-720)*1`.
// Addition, subtraction, SUM and EXP keep a subnormal (Mac Excel 365).

#ifndef FORMULON_UTILS_SUBNORMAL_FLUSH_H_
#define FORMULON_UTILS_SUBNORMAL_FLUSH_H_

#include <cfloat>
#include <cmath>

namespace formulon {

/// Returns 0 for a nonzero `r` below DBL_MIN in magnitude; otherwise `r`.
inline double flush_subnormal(double r) noexcept {
  return r != 0.0 && std::fabs(r) < DBL_MIN ? 0.0 : r;
}

}  // namespace formulon

#endif  // FORMULON_UTILS_SUBNORMAL_FLUSH_H_
