// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Internal header -- do not include outside `src/eval/builtins/financial*`.
//
// Forward declaration for the irregular-last-period bond-pricing
// built-in. Implementation lives in `financial_oddlprice.cpp` and is
// registered from `financial.cpp` via `register_financial_builtins`.
//
// Functions declared here:
//   * ODDLPRICE -- clean price per 100 face for a security whose final
//                  coupon period is irregular (the bond pays its
//                  periodic coupons up to `last_interest` and then a
//                  single irregular coupon + redemption at `maturity`).
//                  Uses Microsoft's documented closed form:
//
//                    cf   = 100 * rate / freq * (DC_total / E)
//                    ai   = 100 * rate / freq * (A_total  / E)
//                    disc = 1 + DSC * yld / freq / E
//                    ODDLPRICE = (redemption + cf) / disc - ai
//
//                  See `financial_oddl_helpers.h` for the schedule
//                  walker that supplies DC_total / A_total / DSC / E.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_ODDLPRICE_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_ODDLPRICE_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value OddlPrice(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDLPRICE_H_
