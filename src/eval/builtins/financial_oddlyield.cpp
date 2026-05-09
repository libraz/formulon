// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto last_interest = read_financial_date(args, 2);
  if (!last_interest) {
    return last_interest.error();
  }
  auto rate = read_required_number(args, 3);
  if (!rate) {
    return rate.error();
  }
  auto pr = read_required_number(args, 4);
  if (!pr) {
    return pr.error();
  }
  auto redemption = read_required_number(args, 5);
  if (!redemption) {
    return redemption.error();
  }
  auto frequency_e = read_coupon_frequency(args, 6);
  if (!frequency_e) {
    return frequency_e.error();
  }
  auto basis_e = read_day_count_basis(args, arity, 7);
  if (!basis_e) {
    return basis_e.error();
  }

  // Validation order mirrors ODDLPRICE: date ordering -> frequency ->
  // basis -> rate sign -> pr sign (replaces yld sign) -> redemption
  // sign. The `pr <= 0` check is the YIELD-family-specific addition;
  // ODDLPRICE accepts a zero yld but ODDLYIELD rejects a non-positive
  // market price (zero would imply infinite yield).
  if (last_interest.value() >= settlement.value()) {
    return ErrorCode::Num;
  }
  if (settlement.value() >= maturity.value()) {
    return ErrorCode::Num;
  }
  const int frequency = frequency_e.value();
  const int basis = basis_e.value();
  if (rate.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (pr.value() <= 0.0) {
    return ErrorCode::Num;
  }
  if (redemption.value() <= 0.0) {
    return ErrorCode::Num;
  }

  auto sched = compute_odd_last_schedule(settlement.value(), maturity.value(), last_interest.value(), frequency, basis);
  if (!sched) {
    return sched.error();
  }
  const double dc = sched.value().dc_total;
  const double a = sched.value().a_total;
  const double dsc = sched.value().dsc;
  const double e = sched.value().e;

  const double freq_d = static_cast<double>(frequency);
  const double cf = 100.0 * rate.value() / freq_d * (dc / e);
  const double ai = 100.0 * rate.value() / freq_d * (a / e);
  const double pr_v = pr.value();
  const double denom = pr_v + ai;
  if (denom == 0.0 || dsc == 0.0) {
    return ErrorCode::Num;
  }
  const double yld = (freq_d * e / dsc) * ((redemption.value() + cf - pr_v - ai) / denom);
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
