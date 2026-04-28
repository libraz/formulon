// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Forward declaration for the irregular-first-period bond yield-to-
// maturity built-in. Implementation lives in `financial_oddfyield.cpp`
// and is registered from `financial.cpp` via
// `register_financial_builtins`.
//
// Functions declared here:
//   * ODDFYIELD -- yield-to-maturity (decimal) for a security whose
//                  first coupon period is irregular. Numerically
//                  inverts the ODDFPRICE kernel via Newton-Raphson on
//                  `compute_oddf_clean_price(yld) - pr` with a
//                  central-difference derivative. Unlike ODDLYIELD
//                  (which has a closed-form analytic inverse because
//                  the ODDLPRICE residual uses simple-interest
//                  discounting), ODDFYIELD must iterate: ODDFPRICE is
//                  a polynomial in `v = 1/(1 + yld/freq)` of degree
//                  `nc + n_regular`, with no algebraic inverse.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFYIELD_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFYIELD_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value OddfYield(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFYIELD_H_
