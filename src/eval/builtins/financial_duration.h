// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal header — do not include outside `src/eval/builtins/financial*`.
//
// Forward declarations for the Macaulay-duration / modified-duration bond
// built-ins. Implementations live in `financial_duration.cpp` and are
// registered from `financial.cpp` via `register_financial_builtins`.
//
// Functions declared here:
//   * DURATION  -- Macaulay duration in years.
//   * MDURATION -- Modified duration = DURATION / (1 + yld/frequency).
//
// Both share the closed-form formula in the .cpp; both delegate
// coupon-schedule mechanics to the shared engine in
// `eval/coupon_schedule.h`.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_DURATION_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_DURATION_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value Duration(const Value* args, std::uint32_t arity, Arena& arena);
Value MDuration(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_DURATION_H_
