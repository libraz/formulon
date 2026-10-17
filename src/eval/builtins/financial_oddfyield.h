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
//                  inverts ODDFPRICE with YIELD's Newton iteration.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFYIELD_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFYIELD_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value OddfYield(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_ODDFYIELD_H_
