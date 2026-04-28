// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Forward declaration for the irregular-last-period bond yield-to-
// maturity built-in. Implementation lives in `financial_oddlyield.cpp`
// and is registered from `financial.cpp` via
// `register_financial_builtins`.
//
// Functions declared here:
//   * ODDLYIELD -- yield-to-maturity (decimal) for a security whose
//                  final coupon period is irregular. The analytic
//                  inverse of ODDLPRICE: solving
//
//                    price = (redemption + cf) / (1 + DSC*yld/freq/E) - ai
//
//                  for `yld` is closed-form, no Newton-Raphson needed:
//
//                    yld = (freq * E / DSC) *
//                          ((redemption + cf - price - ai) / (price + ai))
//
//                  See `financial_oddl_helpers.h` for the schedule
//                  walker shared with ODDLPRICE.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_ODDLYIELD_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_ODDLYIELD_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value OddlYield(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDLYIELD_H_
