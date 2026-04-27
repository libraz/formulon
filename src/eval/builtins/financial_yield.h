// Copyright 2026 libraz. Licensed under the MIT License.
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

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value Yield(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_YIELD_H_
