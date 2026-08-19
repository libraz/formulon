//
// Implementation of the irregular-last-period bond yield-to-maturity
// built-in:
//
//   * ODDLYIELD(settlement, maturity, last_interest, rate, pr,
//               redemption, frequency, [basis=0])
//
// Returns the annual yield-to-maturity (decimal) implied by the given
// clean market price `pr`. ODDLYIELD is the analytic inverse of
// ODDLPRICE -- and unlike YIELD (which needs Newton-Raphson when more
// than one coupon remains), ODDLYIELD has a closed-form solution
// because ODDLPRICE only ever uses simple-interest discounting on a
// single residual period:
//
//   pr = (redemption + cf) / (1 + DSC * yld / freq / E) - ai
//
// solving for `yld`:
//
//   1 + DSC * yld / freq / E = (redemption + cf) / (pr + ai)
//   DSC * yld / freq / E     = (redemption + cf) / (pr + ai) - 1
//                             = (redemption + cf - pr - ai) / (pr + ai)
//   yld                       = (freq * E / DSC) *
//                              ((redemption + cf - pr - ai) / (pr + ai))
//
// See `financial_oddl_helpers.h` for the schedule walker shared with
// ODDLPRICE.

#include "eval/builtins/financial_oddlyield.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddl_helpers.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Expected<double, ErrorCode> compute_oddl_yield(const Value* args, std::uint32_t arity) {
  auto in = read_odd_last_inputs(args, arity, /*slot4_must_be_positive=*/true);
  if (!in) {
    return in.error();
  }
  const double pr_v = in.value().slot4;
  const double denom = pr_v + in.value().ai;
  if (denom == 0.0 || in.value().dsc == 0.0) {
    return ErrorCode::Num;
  }
  const double yld = (in.value().freq_d * in.value().e / in.value().dsc) *
                     ((in.value().redemption + in.value().cf - pr_v - in.value().ai) / denom);
  if (std::isnan(yld) || std::isinf(yld)) {
    return ErrorCode::Num;
  }
  return yld;
}

// --- ODDLYIELD(settlement, maturity, last_interest, rate, pr,
//              redemption, frequency, [basis=0]) ----------------------------
//
// Annual yield-to-maturity (decimal) for a security whose final coupon
// period is irregular. The analytic closed-form inverse of ODDLPRICE.
Value OddlYield(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto y = compute_oddl_yield(args, arity);
  if (!y) {
    return Value::error(y.error());
  }
  return finalize(y.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
