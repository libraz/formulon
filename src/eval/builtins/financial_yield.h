//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Forward declaration for the regular-period bond yield-to-maturity
// built-in. Implementation lives in `financial_yield.cpp` and is
// registered from `financial.cpp` via `register_financial_builtins`.
//
// Functions declared here:
//   * YIELD -- yield-to-maturity (decimal) for a security paying periodic
//              interest. Inverts the closed-form PRICE function. Two
//              numerical regimes:
//                - n == 1 (last coupon period remaining): closed-form
//                  analytic solution from the simple-interest discount
//                  branch of PRICE.
//                - n >  1: Newton-Raphson over yld, calling
//                  `compute_clean_price` from `financial_clean_price.h`
//                  on each iterate. Initial guess uses Microsoft's
//                  documented approximate-yield formula. Converges to
//                  |f| < 1e-12*(|pr|+1) or |delta| < 1e-15 within ~100
//                  iterations on every well-formed bond; non-convergence
//                  surfaces as `#NUM!`.
//
// The implementation reuses PRICE's clean-price kernel rather than
// re-deriving the bond cash-flow sum, so any future change to PRICE's
// numerical behaviour automatically flows into YIELD. See the .cpp for
// the full closed-form derivation of the n==1 branch.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_YIELD_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_YIELD_H_

#include <cmath>
#include <cstdint>
#include <utility>

#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value Yield(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

// Settlement-to-maturity span in years: whole anniversary years of settlement
// plus the remaining days over the basis' year length (basis 1: the length of
// the anniversary year holding the remainder).
double years_to_maturity(double settlement, double maturity, int basis, bool date1904);

// Excel's start for the yield Newton iteration, from the coupon rate, the
// market price and the settlement-to-maturity span `ny` in years.
inline double yield_initial_guess(double rate, double pr, double ny) noexcept {
  const double d = pr - 100.0;
  double denom = ny * 100.0;
  denom = (ny * d) / 2.0 + denom;
  denom = d / 4.0 + denom;
  return (((rate * ny) * 100.0) - d) / denom;
}

// Newton shell shared by YIELD and ODDFYIELD. `step_at(yld)` returns the
// Newton step `(price - pr) / slope` as an `Expected<double, ErrorCode>`.
// A step below `1e-10` ends the iteration and is NOT applied; a non-finite
// step or 100 iterations without convergence is `#NUM!`.
template <typename StepFn>
Expected<double, ErrorCode> newton_yield_by_steps(double yld, StepFn step_at) {
  constexpr int kMaxNewtonSteps = 100;
  constexpr double kStepTolerance = 1.0e-10;
  for (int iter = 0; iter < kMaxNewtonSteps; ++iter) {
    Expected<double, ErrorCode> step = step_at(yld);
    if (!step) {
      return std::move(step.error());
    }
    if (std::isnan(step.value()) || std::isinf(step.value())) {
      return ErrorCode::Num;
    }
    if (std::fabs(step.value()) < kStepTolerance) {
      return yld;
    }
    yld -= step.value();
  }
  return ErrorCode::Num;
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_YIELD_H_
