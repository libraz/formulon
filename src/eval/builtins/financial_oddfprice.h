// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Forward declaration for the irregular-first-period bond-pricing
// built-in. Implementation lives in `financial_oddfprice.cpp` and is
// registered from `financial.cpp` via `register_financial_builtins`.
//
// Functions declared here:
//   * ODDFPRICE -- clean price per 100 face for a security whose first
//                  coupon period is irregular (the bond is issued at
//                  `issue`, pays its first coupon at `first_coupon`,
//                  and pays subsequent coupons on the regular grid up
//                  to `maturity`). Generalises Microsoft's documented
//                  short-first-period and long-first-period formulas
//                  via a per-quasi-period schedule walk anchored on
//                  `first_coupon`. See `financial_oddf_helpers.h` for
//                  the kernel shared with ODDFYIELD.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFPRICE_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFPRICE_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value OddfPrice(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFPRICE_H_
