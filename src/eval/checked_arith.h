//
// Checked double arithmetic of the discounted-sum walks behind IRR and NPV.
// An overflow or NaN yields 0 with fault code 1, a divisor below the normal
// range yields 0 with fault code 4, and a subnormal product or quotient flushes
// to 0. A walk continues after a fault and the last code stays.

#ifndef FORMULON_EVAL_CHECKED_ARITH_H_
#define FORMULON_EVAL_CHECKED_ARITH_H_

#include <cmath>
#include <limits>

#include "utils/subnormal_flush.h"

namespace formulon {
namespace eval {

class CheckedArith {
 public:
  static constexpr int kFaultOverflow = 1;
  static constexpr int kFaultDivide = 4;

  double add(double a, double b) noexcept { return settle(a + b); }

  double mul(double a, double b) noexcept { return flush_subnormal(settle(a * b)); }

  double div(double a, double b) noexcept {
    if (std::fabs(b) < std::numeric_limits<double>::min()) {
      fault_ = kFaultDivide;
      return 0.0;
    }
    return flush_subnormal(settle(a / b));
  }

  /// Last fault code left so far, or 0.
  int fault() const noexcept { return fault_; }

  bool failed() const noexcept { return fault_ != 0; }

 private:
  double settle(double r) noexcept {
    if (std::isnan(r) || std::fabs(r) > std::numeric_limits<double>::max()) {
      fault_ = kFaultOverflow;
      return 0.0;
    }
    return r;
  }

  int fault_ = 0;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_CHECKED_ARITH_H_
