// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers the coupon-date financial built-ins (COUPPCD, COUPNCD,
// COUPNUM, COUPDAYBS, COUPDAYSNC, COUPDAYS) into a FunctionRegistry.
// All six share the shared coupon-schedule engine in
// `eval/coupon_schedule.h`; the same engine will be reused by the
// bond-pricing family (PRICE / YIELD / DURATION / MDURATION /
// ACCRINT) when those land.
//
// Kept in its own translation unit so the coupon + future bond
// families can evolve independently of the core TVM functions in
// `financial.cpp` and the security-rate family in `financial_rates.cpp`.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_COUPON_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_COUPON_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the six coupon-date built-ins (COUPPCD, COUPNCD,
/// COUPNUM, COUPDAYBS, COUPDAYSNC, COUPDAYS) into `registry`.
void register_financial_coupon_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_COUPON_H_
