//
// Implementation of the regular-period bond yield-to-maturity built-in:
//
//   * YIELD(settlement, maturity, rate, pr, redemption, frequency,
//           [basis=0])
//
// Returns the annual yield (decimal) implied by the given clean market
// price `pr`. YIELD is the analytic inverse of PRICE: feeding YIELD's
// output back into PRICE (with the same settlement/maturity/rate/...
// arguments) recovers `pr` to ~1e-9.
//
// With A = accrued days, E = coupon period days, DSM = days settlement to maturity:
//
//   * n == 1 (settlement inside the final coupon period): closed form
//
//       YIELD = (((red/100 + rate/freq) - (pr/100 + (rate/freq) * (A/E))) / (pr/100 + (rate/freq) * (A/E)))
//               * ((freq * E) / DSM)
//
//     with E the actual days from the previous coupon date to maturity on
//     bases 2 and 3. A zero denominator is `#NUM!`.
//
//   * n > 1: Newton's method on the multi-period PRICE with its analytic
//     derivative, from y0 = (((rate * Y) * 100) - d) / (d/4 + ((Y * d)/2 + Y * 100)),
//     d = pr - 100, Y the settlement-to-maturity span in years. A step below
//     1e-10 ends the iteration and is not applied. Negative yields are
//     allowed; no convergence in 100 steps is `#NUM!`.
//
// The clean-price kernel itself is shared with PRICE via
// `financial_clean_price.h` -- both functions agree by construction.

#include "eval/builtins/financial_yield.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_clean_price.h"
#include "eval/builtins/financial_helpers.h"
#include "eval/coupon_schedule.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
double years_to_maturity(double settlement, double maturity, int basis, bool date1904) {
  const double s = std::trunc(settlement);
  const double m = std::trunc(maturity);
  const date_time::YMD sd = date_time::ymd_from_serial(s, date1904);
  const date_time::YMD md = date_time::ymd_from_serial(m, date1904);
  const bool after_anniversary = md.m > sd.m || (md.m == sd.m && md.d >= sd.d);
  double whole = static_cast<double>(md.y) - static_cast<double>(sd.y);
  int anniversary_year = md.y;
  if (!after_anniversary) {
    whole = whole - 1.0;
    anniversary_year = md.y - 1;
  } else if (md.m == sd.m && md.d == sd.d) {
    return whole;
  }
  // The anniversary of settlement in the year before maturity, rolled over
  // (a February 29 settlement lands on March 1 in a common year).
  const double anniversary =
      date_time::serial_from_ymd(anniversary_year, sd.m, 1u, date1904) + static_cast<double>(sd.d) - 1.0;
  const double days = date_time::basis_days_between(anniversary, m, basis, date1904);
  double year = 360.0;
  if (basis == 3) {
    year = 365.0;
  } else if (basis == 1) {
    year = date_time::days_in_month(md.y, 2u) == 29u ? 366.0 : 365.0;
  }
  return days / year + whole;
}

double bond_price_slope(double yld, double rate, double redemption, double freq, double dsc_over_e, std::int32_t n) {
  const double base = (yld / freq) + 1.0;
  const double weight = (rate / freq) * 100.0;
  double slope = 0.0;
  for (std::int32_t k = 1; k <= n; ++k) {
    const double power = excel_power(base, static_cast<double>(k) + dsc_over_e);
    const double t = dsc_over_e + (static_cast<double>(k) - 1.0);
    slope = ((-weight * t) / (freq * power)) + slope;
  }
  const double last = (dsc_over_e + (static_cast<double>(n) - 1.0)) * redemption;
  return slope - last / (freq * excel_power(base, dsc_over_e + static_cast<double>(n)));
}

namespace {

// Computes the YIELD result. Returns the yield-to-maturity (decimal) on
// success or `ErrorCode::Num` on any validation / numerical failure.
Expected<double, ErrorCode> compute_yield(const Value* args, std::uint32_t arity, bool date1904) {
  auto in = read_coupon_bond(args, arity, date1904);
  if (!in) {
    return std::move(in.error());
  }
  const auto& [tail, cd, freq_d, cf] = in.value();
  const double rate = tail.rate;
  const double pr_v = tail.amount;
  const double red = tail.redemption;
  // YIELD rejects a non-positive market price (zero would mean "infinite yield"); PRICE accepts a zero yld.
  if (pr_v <= 0.0) {
    return ErrorCode::Num;
  }
  const std::int32_t n = cd.coupons_remaining;
  auto settlement = read_financial_date(args, 0);
  auto maturity = read_financial_date(args, 1);
  if (!settlement || !maturity) {
    return ErrorCode::Num;
  }
  const double accrued_days = cd.days_bs;
  double period_days = cd.period_days;

  // --- n == 1: closed-form inversion of the simple-interest discount branch.
  if (n == 1) {
    if (tail.basis == 2 || tail.basis == 3) {
      period_days = std::trunc(maturity.value()) - cd.pcd;
    }
    const double dsm = date_time::basis_days_between(std::trunc(settlement.value()), std::trunc(maturity.value()),
                                                     tail.basis, date1904);
    const double coupon = rate / freq_d;
    const double face = red / 100.0 + coupon;
    const double paid = pr_v / 100.0 + coupon * (accrued_days / period_days);
    if (paid == 0.0 || dsm <= 0.0) {
      return ErrorCode::Num;
    }
    const double yld = ((face - paid) / paid) * ((freq_d * period_days) / dsm);
    if (std::isnan(yld) || std::isinf(yld)) {
      return ErrorCode::Num;
    }
    return yld;
  }

  // --- n > 1: Newton's method on the price kernel with its analytic derivative.
  const double ny = years_to_maturity(settlement.value(), maturity.value(), tail.basis, date1904);
  const double yld = yield_initial_guess(rate, pr_v, ny);
  // Plain locals: a lambda cannot capture structured bindings in C++17.
  const double freq = freq_d;
  const double e = period_days;
  const double f = (period_days - accrued_days) / period_days;
  const double dsc_over_e = f;
  return newton_yield_by_steps(yld, [&](double y) -> Expected<double, ErrorCode> {
    const double price = bond_price_multi(y, rate, red, freq, e, accrued_days, n);
    const double slope = bond_price_slope(y, rate, red, freq, dsc_over_e, n);
    return (price - pr_v) / slope;
  });
}

}  // namespace

// --- YIELD(settlement, maturity, rate, pr, redemption, frequency,
//           [basis=0]) -----------------------------------------------------
//
// Annual yield-to-maturity (decimal) for a security paying periodic
// interest. The analytic inverse of PRICE.
Value Yield(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto y = compute_yield(args, arity, date1904);
  if (!y) {
    return Value::error(y.error());
  }
  return finalize(y.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
