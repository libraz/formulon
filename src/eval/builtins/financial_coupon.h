//
// Internal header — do not include outside `src/eval/builtins/financial*`.
//
// Forward declarations for the six coupon-date financial built-ins
// (COUPPCD, COUPNCD, COUPNUM, COUPDAYBS, COUPDAYSNC, COUPDAYS). All six
// share the shared coupon-schedule engine in `eval/coupon_schedule.h`;
// the same engine is reused by the bond-pricing family (PRICE / YIELD /
// DURATION / MDURATION / ACCRINT).
//
// Each impl decomposes a date serial through `compute_coupon_dates`,
// which needs the calling workbook's date1904 flag, so all six are
// registered through the lazy financial-date dispatch table
// (`eval/financial_lazy.h`) rather than the eager `FunctionRegistry`.
//
// Kept in its own translation unit so the coupon + bond families can
// evolve independently of the core TVM functions in `financial.cpp` and
// the security-rate family in `financial_rates.cpp`.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_COUPON_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_COUPON_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value CoupPcd(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value CoupNcd(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value CoupNum(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value CoupDayBs(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value CoupDaysNc(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value CoupDays(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_COUPON_H_
