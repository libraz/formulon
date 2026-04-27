// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal header — do not include outside `src/eval/builtins/financial*`.
//
// Forward declaration for the regular-period bond-pricing built-in.
// Implementation lives in `financial_price.cpp` and is registered from
// `financial.cpp` via `register_financial_builtins`.
//
// Functions declared here:
//   * PRICE -- clean price per 100 face for a security paying periodic
//              interest. Two analytic branches: when only one coupon
//              remains the formula uses simple-interest (linear)
//              discounting on the final cash flow; otherwise each cash
//              flow is discounted by `v^(t1+i)` and the accrued interest
//              is subtracted to convert the dirty price to a clean
//              price. See the .cpp for the full derivation.
//
// YIELD will live in this same translation unit once it lands; both
// share the coupon-schedule engine in `eval/coupon_schedule.h` and the
// clean-price computation factored out in this file's anonymous
// namespace.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_PRICE_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_PRICE_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value Price(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_PRICE_H_
