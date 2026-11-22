//
// Implementation of the irregular-first-period bond yield-to-maturity
// built-in:
//
//   * ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr,
//               redemption, frequency, [basis=0])
//
// Returns the annual yield-to-maturity (decimal) implied by the given clean
// market price `pr`: YIELD's Newton iteration (same start from the year count,
// same derivative over the `N` coupons with DSC = days(settlement,
// first_coupon), same stopping rule) with ODDFPRICE as the price.

#include "eval/builtins/financial_oddfyield.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddf_helpers.h"
#include "eval/builtins/financial_yield.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Expected<double, ErrorCode> compute_oddf_yield(const Value* args, std::uint32_t arity, bool date1904) {
  ASSIGN_OR_RETURN(auto bond, read_odd_first_bond(args, arity, /*slot5_must_be_positive=*/true, date1904));
  const OddFirstBond& b = bond;
  const double years = years_to_maturity(b.settlement, b.maturity, b.basis, date1904);
  const double dsc_over_e = b.dsc_first / b.e;
  return newton_yield_by_steps(
      yield_initial_guess(b.rate, b.slot5, years), [&](double y) -> Expected<double, ErrorCode> {
        ASSIGN_OR_RETURN(auto price, oddf_price_at(b, y));
        return (price - b.slot5) / bond_price_slope(y, b.rate, b.redemption, b.freq_d, dsc_over_e, b.n);
      });
}

// --- ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr,
//              redemption, frequency, [basis=0]) ----------------------------
//
// Annual yield-to-maturity (decimal) for a security whose first
// coupon period is irregular.
Value OddfYield(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  return finalize_expected(compute_oddf_yield(args, arity, date1904));
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
