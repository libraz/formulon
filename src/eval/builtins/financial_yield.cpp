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
// The numerical strategy mirrors Microsoft's documented YIELD:
//
//   * n == 1 (only one coupon left -- settlement falls inside the final
//     coupon period). PRICE in this branch uses simple-interest
//     discounting on the single remaining cash flow, which inverts in
//     closed form. Solving
//
//       pr = (redemption + cf) / (1 + t1*(yld/freq)) - AI
//
//     for yld gives
//
//       yld = ((redemption + cf - pr - AI) / (pr + AI)) * (E / DSC) * freq
//
//     where cf = 100*rate/freq, AI = 100*rate*A/(E*freq),
//     t1 = DSC/E. (Notation: A = days_bs, E = period_days, DSC = days_nc.)
//     A zero `pr + AI` denominator surfaces as `#NUM!`.
//
//   * n > 1. The PRICE branch has no analytic inverse; Newton-Raphson on
//     the analytic derivative of PRICE starts from
//
//       y0 = (100*(rate*Ny) + (100-pr)) / (Ny*(100+pr)/2 + (pr-100)/4)
//
//     with Ny the settlement-to-maturity span in years, and stops when a
//     step is smaller than 1e-10 -- that last step is not applied. Negative
//     yields are allowed. Non-convergence in 100 steps is `#NUM!`.
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
  const std::int64_t s_days = date_time::days_from_civil(sd.y, sd.m, sd.d);
  const auto anniversary = [&](int years) { return date_time::days_from_civil(sd.y + years, sd.m, sd.d); };
  int years = md.y - sd.y;
  if (anniversary(years) > date_time::days_from_civil(md.y, md.m, md.d)) {
    --years;
  }
  const double anniv = s + static_cast<double>(anniversary(years) - s_days);
  const double remainder_actual = m - anniv;
  double remainder = remainder_actual;
  double year_len = 365.0;
  switch (basis) {
    case 0:
    case 4:
      remainder = date_time::basis_days_between(anniv, m, basis, date1904);
      year_len = 360.0;
      break;
    case 2:
      year_len = 360.0;
      break;
    case 1: {
      const std::int64_t lo = anniversary(years);
      const std::int64_t hi = anniversary(years + 1);
      for (int y = sd.y + years; y <= sd.y + years + 1; ++y) {
        const bool leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
        const std::int64_t feb29 = leap ? date_time::days_from_civil(y, 2, 29) : -1;
        if (leap && feb29 >= lo && feb29 < hi) {
          year_len = 366.0;
        }
      }
      break;
    }
    default:
      break;
  }
  return static_cast<double>(years) + remainder / year_len;
}

namespace {

// Computes the YIELD result. Returns the yield-to-maturity (decimal) on
// success or `ErrorCode::Num` on any validation / numerical failure.
Expected<double, ErrorCode> compute_yield(const Value* args, std::uint32_t arity, bool date1904) {
  auto in = read_coupon_bond(args, arity, date1904);
  if (!in) {
    return std::move(in.error());
  }
  const auto& [tail, cd, freq_d, cf, ai] = in.value();
  const double rate = tail.rate;
  const double pr_v = tail.amount;
  const double red = tail.redemption;
  // YIELD rejects a non-positive market price (zero would mean "infinite yield"); PRICE accepts a zero yld.
  if (pr_v <= 0.0) {
    return ErrorCode::Num;
  }
  const std::int32_t n = cd.coupons_remaining;

  // --- n == 1: closed-form analytic inversion of the simple-interest
  // discount branch. See the file-level comment for the derivation.
  // We use `bond_dsc` (not the raw `days_nc`) because for bases 2 / 3
  // the bond formula needs `days_bs + DSC == period_days`; see
  // `coupon_schedule.h` for the rationale.
  if (n == 1) {
    const double denom = pr_v + ai;
    if (denom == 0.0) {
      return ErrorCode::Num;
    }
    if (cd.bond_dsc <= 0.0) {
      return ErrorCode::Num;
    }
    const double yld = ((red + cf - pr_v - ai) / denom) * (cd.period_days / cd.bond_dsc) * freq_d;
    if (std::isnan(yld) || std::isinf(yld)) {
      return ErrorCode::Num;
    }
    return yld;
  }

  // --- n > 1: Newton-Raphson on the price kernel, see the file header.
  auto settlement = read_financial_date(args, 0);
  auto maturity = read_financial_date(args, 1);
  if (!settlement || !maturity) {
    return ErrorCode::Num;
  }
  const double ny = years_to_maturity(settlement.value(), maturity.value(), tail.basis, date1904);
  const double yld = yield_initial_guess(rate, pr_v, ny);
  // Plain locals: a lambda cannot capture structured bindings in C++17.
  const double freq = freq_d;
  const double coupon = cf;
  const double days_bs = cd.days_bs;
  const double period_days = cd.period_days;
  const double t1 = cd.bond_dsc / period_days;
  return newton_yield_by_steps(yld, [&](double y) -> Expected<double, ErrorCode> {
    const double b = 1.0 + y / freq;
    const double t_n = static_cast<double>(n - 1) + t1;
    double price = red / excel_power(b, t_n);
    double coupons = 0.0;
    for (std::int32_t k = 1; k <= n; ++k) {
      coupons += coupon / excel_power(b, static_cast<double>(k - 1) + t1);
    }
    price = price + coupons - coupon * days_bs / period_days;
    double d = -t_n * red / excel_power(b, static_cast<double>(n) + t1);
    for (std::int32_t k = 1; k <= n; ++k) {
      d += -(static_cast<double>(k - 1) + t1) * coupon / excel_power(b, static_cast<double>(k) + t1);
    }
    return (price - pr_v) / (d / freq);
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
