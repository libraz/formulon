//
// Implementation of the regular-period bond-pricing built-in:
//
//   * PRICE(settlement, maturity, rate, yld, redemption, frequency,
//           [basis=0])
//
// Returns the clean price per 100 face. The classical (Mac Excel 365 /
// Microsoft documented) closed form has two branches keyed on the number
// of coupons remaining after settlement:
//
//   t1  = days_nc / period_days        (fractional periods to NCD)
//   cf  = 100 * rate / frequency       (per-period coupon cash flow)
//   AI  = 100 * rate * days_bs / (period_days * frequency)
//
//   if n == 1:
//     PRICE = (redemption + cf) / (1 + t1 * (yld / frequency)) - AI
//   else:
//     b     = 1 + yld / frequency
//     PRICE = redemption / b^(n-1+t1) + sum_{k=1..n} cf / b^(k-1+t1) - cf * days_bs / period_days
//
// The two branches do NOT collapse at full precision: the n==1 branch
// uses simple-interest (linear) discounting of the single remaining cash
// flow, matching Microsoft's documented last-period formula. Substituting
// `v^t1` for `1 / (1 + t1*(yld/frequency))` would change the result at
// the sub-penny level on real bonds and break 1-bit oracle parity, so
// the impl deliberately keeps both forms.
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
  CouponBondInputs out{tail.value(), CouponDates{}, 0.0, 0.0, 0.0};
  if (!compute_coupon_dates(settlement.value(), maturity.value(), out.tail.frequency, out.tail.basis, date1904,
                            &out.cd)) {
    return ErrorCode::Num;
  }
  if (out.cd.coupons_remaining <= 0 || out.cd.period_days <= 0.0) {
    return ErrorCode::Num;
  }
  out.freq_d = static_cast<double>(out.tail.frequency);
  out.cf = 100.0 * out.tail.rate / out.freq_d;
  out.ai = 100.0 * out.tail.rate * out.cd.days_bs / (out.cd.period_days * out.freq_d);
  return out;
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
  const auto& [tail, cd, freq_d, cf, ai] = in.value();
  const double yld = tail.amount;
  const double red = tail.redemption;
  if (yld < 0.0) {
    return ErrorCode::Num;
  }

  // `bond_dsc` (not the raw `days_nc`) is the DSC the bond formula
  // needs: it satisfies `days_bs + bond_dsc == period_days` for all
  // bases, including bases 2 / 3 where the raw COUPDAYSNC value would
  // break the identity. See `coupon_schedule.h` for the rationale.
  const double t1 = cd.bond_dsc / cd.period_days;
  const std::int32_t n = cd.coupons_remaining;
  const double yfreq = yld / freq_d;

  // Single-period branch: simple-interest discounting on the final
  // cash flow. See the file header for why this is NOT the same as
  // substituting v^t1 below.
  if (n == 1) {
    const double denom = 1.0 + t1 * yfreq;
    if (denom == 0.0) {
      return ErrorCode::Num;
    }
    const double price = (red + cf) / denom - ai;
    if (std::isnan(price) || std::isinf(price)) {
      return ErrorCode::Num;
    }
    return price;
  }

  // Multi-period branch, summed the way Excel does: the redemption term and
  // the coupon sum are accumulated separately, then added, and the accrued
  // interest is `cf * (days_bs / period_days)`.
  const double base = 1.0 + yfreq;
  const double t_n = static_cast<double>(n - 1) + t1;
  double coupons = 0.0;
  for (std::int32_t k = 1; k <= n; ++k) {
    coupons += cf / excel_power(base, static_cast<double>(k - 1) + t1);
  }
  const double price = red / excel_power(base, t_n) + coupons - cf * cd.days_bs / cd.period_days;
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
