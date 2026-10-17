//
// Implementation of the regular-period bond-pricing built-in:
//
//   * PRICE(settlement, maturity, rate, yld, redemption, frequency,
//           [basis=0])
//
// Returns the clean price per 100 face. The closed form has two branches
// keyed on the number of coupons remaining after settlement:
//
//   E = coupon period days, A = accrued days, DSC = E - A, cf = 100 * rate / frequency
//
//   if n == 1:
//     PRICE = (cf + redemption) / ((yld / frequency) * (DSC / E) + 1) - cf * (A / E)
//   else:
//     b     = yld / frequency + 1,  f = DSC / E
//     PRICE = (redemption / b^((n - 1) + f) + sum_{k=1..n} cf / b^((k - 1) + f))
//             - ((rate / frequency) * (A / E)) * 100
//
// The branches do NOT collapse: the n == 1 branch discounts the single
// remaining cash flow with simple interest, so both forms are kept.
//
// Coupon-schedule mechanics are delegated to the shared engine in
// `eval/coupon_schedule.h` (the same engine used by COUP*, DURATION,
// MDURATION, and ACCRINT).

#include "eval/builtins/financial_price.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_clean_price.h"
#include "eval/builtins/financial_helpers.h"
#include "eval/coupon_schedule.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
Expected<CouponBondInputs, ErrorCode> read_coupon_bond(const Value* args, std::uint32_t arity, bool date1904) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return std::move(settlement.error());
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return std::move(maturity.error());
  }
  auto tail = read_coupon_bond_tail(args, arity, 2);
  if (!tail) {
    return std::move(tail.error());
  }
  // Every domain check past the reads is `#NUM!`, so the tail reader's rate / redemption checks may run
  // ahead of the date ordering.
  if (settlement.value() >= maturity.value()) {
    return ErrorCode::Num;
  }
  CouponBondInputs out{tail.value(), CouponDates{}, 0.0, 0.0};
  if (!compute_coupon_dates(settlement.value(), maturity.value(), out.tail.frequency, out.tail.basis, date1904,
                            &out.cd)) {
    return ErrorCode::Num;
  }
  if (out.cd.coupons_remaining <= 0 || out.cd.period_days <= 0.0) {
    return ErrorCode::Num;
  }
  out.freq_d = static_cast<double>(out.tail.frequency);
  out.cf = 100.0 * out.tail.rate / out.freq_d;
  return out;
}

double bond_price_multi(double yld, double rate, double redemption, double freq, double e, double a,
                        std::int32_t n) noexcept {
  const double base = (yld / freq) + 1.0;
  const double f = (e - a) / e;
  const double cf = (100.0 * rate) / freq;
  double coupons = 0.0;
  for (std::int32_t k = 1; k <= n; ++k) {
    coupons = (cf / excel_power(base, (static_cast<double>(k) - 1.0) + f)) + coupons;
  }
  const double redemption_term = redemption / excel_power(base, (static_cast<double>(n) - 1.0) + f);
  return (redemption_term + coupons) - ((rate / freq) * (a / e)) * 100.0;
}

// Computes the clean price per 100 face. Returns `#NUM!` on any
// validation or numerical failure. Factored out of the Value-returning
// `Price` so YIELD (Newton iteration over yld in `financial_yield.cpp`)
// can call the same closed form without re-parsing arguments. Declared in
// `financial_clean_price.h` so the YIELD TU can include only that header.
Expected<double, ErrorCode> compute_clean_price(const Value* args, std::uint32_t arity, bool date1904) {
  auto in = read_coupon_bond(args, arity, date1904);
  if (!in) {
    return std::move(in.error());
  }
  const auto& [tail, cd, freq_d, cf] = in.value();
  const double yld = tail.amount;
  const double red = tail.redemption;
  if (yld < 0.0) {
    return ErrorCode::Num;
  }

  // DSC is the rest of the coupon period after the accrued days, E - A, so that
  // `A / E + DSC / E` is one for every basis.
  const double e = cd.period_days;
  const double a = cd.days_bs;
  const double dsc = e - a;
  const std::int32_t n = cd.coupons_remaining;
  const double yfreq = yld / freq_d;

  double price = 0.0;
  if (n == 1) {
    // Single-period branch: simple-interest discounting on the final cash flow.
    const double denom = (yfreq * (dsc / e)) + 1.0;
    if (denom == 0.0) {
      return ErrorCode::Num;
    }
    price = ((cf + red) / denom) - cf * (a / e);
  } else {
    price = bond_price_multi(yld, tail.rate, red, freq_d, e, a, n);
  }
  if (std::isnan(price) || std::isinf(price)) {
    return ErrorCode::Num;
  }
  return price;
}

// --- PRICE(settlement, maturity, rate, yld, redemption, frequency,
//           [basis=0]) -----------------------------------------------------
//
// Clean price per 100 face for a security paying periodic interest.
Value Price(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto p = compute_clean_price(args, arity, date1904);
  if (!p) {
    return Value::error(p.error());
  }
  return finalize(p.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
