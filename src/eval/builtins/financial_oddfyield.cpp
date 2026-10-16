//
// Implementation of the irregular-first-period bond yield-to-maturity
// built-in:
//
//   * ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr,
//               redemption, frequency, [basis=0])
//
// Returns the annual yield-to-maturity (decimal) implied by the given
// clean market price `pr`. Excel runs YIELD's Newton iteration (same
// start, same step shell) on ODDFPRICE, but divides by the derivative of
// the REGULAR bond price, i.e. with the odd first coupon treated as a full
// coupon:
//
//   D = ( -(N - 1 + t1) * red / b^(N + t1)
//         - sum_{k=1..N} (k - 1 + t1) * cf / b^(k + t1) ) / freq
//
// with N = n_regular + 1 coupons from settlement, t1 = nq + dsc / E and
// b = 1 + yld / freq. That derivative runs a few 1e-4 too large, so the last
// steps converge linearly and the answer sits ~1e-12 off the exact root.
//
// An issue-to-first_coupon span of exactly one regular period is `#NUM!`.

#include "eval/builtins/financial_oddfyield.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddf_helpers.h"
#include "eval/builtins/financial_yield.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Derivative of the regular PRICE formula at yield `yld`, in Excel's term order.
double regular_price_derivative(const OddFirstSchedule& sched, const OddFirstArgs& in, double yld) {
  const double cf = 100.0 * in.rate / in.freq_d;
  const double b = 1.0 + yld / in.freq_d;
  const double t1 = static_cast<double>(sched.nq) + sched.dsc / sched.e;
  const double n = static_cast<double>(sched.n_regular) + 1.0;
  double d = -(n - 1.0 + t1) * (in.redemption / excel_power(b, n + t1));
  for (std::int32_t k = 1; k <= sched.n_regular + 1; ++k) {
    d += -(static_cast<double>(k - 1) + t1) * (cf / excel_power(b, static_cast<double>(k) + t1));
  }
  return d / in.freq_d;
}

}  // namespace

Expected<double, ErrorCode> compute_oddf_yield(const Value* args, std::uint32_t arity, bool date1904) {
  OddFirstSchedule sched{};
  auto in = read_odd_first_inputs(args, arity, /*slot5_must_be_positive=*/true, date1904, sched);
  if (!in) {
    return std::move(in.error());
  }
  const double pr = in.value().slot5;
  const double issue = std::trunc(read_financial_date(args, 2).value());
  if (sched.nc == 1 && sched.qp[0].start == issue) {
    return ErrorCode::Num;
  }
  const double ny = years_to_maturity(read_financial_date(args, 0).value(), read_financial_date(args, 1).value(),
                                      read_day_count_basis(args, arity, 8).value(), date1904);
  return newton_yield_by_steps(yield_initial_guess(in.value().rate, pr, ny),
                               [&](double y) -> Expected<double, ErrorCode> {
                                 auto price = oddf_price_at(sched, in.value(), y);
                                 if (!price) {
                                   return std::move(price.error());
                                 }
                                 return (price.value() - pr) / regular_price_derivative(sched, in.value(), y);
                               });
}

// --- ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr,
//              redemption, frequency, [basis=0]) ----------------------------
//
// Annual yield-to-maturity (decimal) for a security whose first
// coupon period is irregular. Numerically inverts ODDFPRICE.
Value OddfYield(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto y = compute_oddf_yield(args, arity, date1904);
  if (!y) {
    return Value::error(y.error());
  }
  return finalize(y.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
