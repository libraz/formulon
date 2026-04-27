// Copyright 2026 libraz. Licensed under the MIT License.
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
//   v   = 1 / (1 + yld / frequency)
//   t1  = days_nc / period_days        (fractional periods to NCD)
//   cf  = 100 * rate / frequency       (per-period coupon cash flow)
//   AI  = 100 * rate * days_bs / (period_days * frequency)
//
//   if n == 1:
//     PRICE = (redemption + cf) / (1 + t1 * (yld / frequency)) - AI
//   else:
//     dirty = sum_{i=0..n-1}  CF_i * v^(t1 + i)
//             where  CF_i  = cf                  for i < n-1
//                    CF_{n-1} = cf + redemption
//     PRICE = dirty - AI
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
#include "eval/coerce.h"
#include "eval/coupon_schedule.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Reads a required date argument, truncating toward zero. Negative serials
// are rejected as `#NUM!`. Mirrors the helper in `financial_duration.cpp`
// and `financial_bond_simple.cpp`; replicated per-TU to keep the helper
// local to its callers (see those files for the established pattern).
Expected<double, ErrorCode> read_date(const Value* args, std::uint32_t index) {
  auto raw = read_required_number(args, index);
  if (!raw) {
    return raw.error();
  }
  const double t = std::trunc(raw.value());
  if (t < 0.0) {
    return ErrorCode::Num;
  }
  return t;
}

}  // namespace

// Computes the clean price per 100 face. Returns `#NUM!` on any
// validation or numerical failure. Factored out of the Value-returning
// `Price` so YIELD (Newton iteration over yld in `financial_yield.cpp`)
// can call the same closed form without re-parsing arguments. Declared in
// `financial_clean_price.h` so the YIELD TU can include only that header.
Expected<double, ErrorCode> compute_clean_price(const Value* args, std::uint32_t arity) {
  auto settlement = read_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto rate = read_required_number(args, 2);
  if (!rate) {
    return rate.error();
  }
  auto yld = read_required_number(args, 3);
  if (!yld) {
    return yld.error();
  }
  auto redemption = read_required_number(args, 4);
  if (!redemption) {
    return redemption.error();
  }
  auto frequency_e = read_required_number(args, 5);
  if (!frequency_e) {
    return frequency_e.error();
  }
  auto basis_e = read_optional_number(args, arity, 6, 0.0);
  if (!basis_e) {
    return basis_e.error();
  }

  // Validation order matches Microsoft's documented contract and the
  // sibling DURATION / PRICE* impls: date ordering -> frequency domain
  // -> basis domain -> rate / yld sign checks -> redemption sign check.
  if (settlement.value() >= maturity.value()) {
    return ErrorCode::Num;
  }
  const int frequency = static_cast<int>(std::trunc(frequency_e.value()));
  if (frequency != 1 && frequency != 2 && frequency != 4) {
    return ErrorCode::Num;
  }
  const int basis = static_cast<int>(std::trunc(basis_e.value()));
  if (basis < 0 || basis > 4) {
    return ErrorCode::Num;
  }
  if (rate.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (yld.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (redemption.value() <= 0.0) {
    return ErrorCode::Num;
  }

  CouponDates cd{};
  if (!compute_coupon_dates(settlement.value(), maturity.value(), frequency, basis, &cd)) {
    return ErrorCode::Num;
  }
  if (cd.coupons_remaining <= 0 || cd.period_days <= 0.0) {
    return ErrorCode::Num;
  }

  const double freq_d = static_cast<double>(frequency);
  const double t1 = cd.days_nc / cd.period_days;
  const double cf = 100.0 * rate.value() / freq_d;
  const double ai = 100.0 * rate.value() * cd.days_bs / (cd.period_days * freq_d);
  const std::int32_t n = cd.coupons_remaining;
  const double red = redemption.value();
  const double yfreq = yld.value() / freq_d;

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

  // Multi-period branch: discount each cash flow by v^(t1+i).
  const double v = 1.0 / (1.0 + yfreq);
  double dirty = 0.0;
  for (std::int32_t i = 0; i < n; ++i) {
    const double cf_i = (i == n - 1) ? (cf + red) : cf;
    const double disc = std::pow(v, t1 + static_cast<double>(i));
    if (std::isnan(disc) || std::isinf(disc)) {
      return ErrorCode::Num;
    }
    dirty += cf_i * disc;
  }
  const double price = dirty - ai;
  if (std::isnan(price) || std::isinf(price)) {
    return ErrorCode::Num;
  }
  return price;
}

// --- PRICE(settlement, maturity, rate, yld, redemption, frequency,
//           [basis=0]) -----------------------------------------------------
//
// Clean price per 100 face for a security paying periodic interest.
Value Price(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto p = compute_clean_price(args, arity);
  if (!p) {
    return Value::error(p.error());
  }
  return finalize(p.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
