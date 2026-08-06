//
// Implementation of the Macaulay-duration / modified-duration bond
// built-ins:
//
//   * DURATION(settlement, maturity, coupon, yld, frequency, [basis=0])
//   * MDURATION(settlement, maturity, coupon, yld, frequency, [basis=0])
//
// Both share the closed-form formula
//
//   v       = 1 / (1 + yld/frequency)
//   t1      = days_nc / period_days        (fractional periods to NCD)
//   time_i  = (i - 1) + t1                 for i = 1..n
//   CF_i    = coupon / frequency           for i < n
//   CF_n    = coupon / frequency + 1       (face = 1)
//
//   num = sum_i (time_i * CF_i * v^time_i)
//   den = sum_i (CF_i * v^time_i)
//
//   DURATION  = (num / den) / frequency
//   MDURATION = DURATION / (1 + yld/frequency)
//
// The face value is normalised to 1 (instead of the textbook 100); the
// factor cancels in num / den so this saves one multiplication per term
// without changing the result.
//
// Coupon-schedule mechanics are delegated to the shared engine in
// `eval/coupon_schedule.h` (the same engine used by the COUP* family
// and the bond-pricing helpers).

#include "eval/builtins/financial_duration.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/coupon_schedule.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Computes Macaulay duration with face = 1. Returns `#NUM!` on any
// validation or numerical failure. The caller (DURATION) returns this
// value directly; MDURATION divides it by `(1 + yld/frequency)`.
Expected<double, ErrorCode> compute_macaulay(const Value* args, std::uint32_t arity) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto coupon = read_required_number(args, 2);
  if (!coupon) {
    return coupon.error();
  }
  auto yld = read_required_number(args, 3);
  if (!yld) {
    return yld.error();
  }
  auto frequency_e = read_coupon_frequency(args, 4);
  if (!frequency_e) {
    return frequency_e.error();
  }
  auto basis_e = read_day_count_basis(args, arity, 5);
  if (!basis_e) {
    return basis_e.error();
  }

  // Validation order matches Microsoft's documented contract and the
  // sibling COUP* / PRICE* impls: date ordering -> frequency domain ->
  // basis domain -> coupon / yld sign checks.
  if (settlement.value() >= maturity.value()) {
    return ErrorCode::Num;
  }
  const int frequency = frequency_e.value();
  const int basis = basis_e.value();
  if (coupon.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (yld.value() < 0.0) {
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
  const double v = 1.0 / (1.0 + yld.value() / freq_d);
  // `bond_dsc` (not the raw `days_nc`) keeps `A/E + DSC/E == 1` on
  // bases 2 / 3; see `coupon_schedule.h` for the rationale.
  const double t1 = cd.bond_dsc / cd.period_days;
  const double cf_coupon = coupon.value() / freq_d;
  const std::int32_t n = cd.coupons_remaining;

  double num = 0.0;
  double den = 0.0;
  for (std::int32_t i = 1; i <= n; ++i) {
    const double time_i = static_cast<double>(i - 1) + t1;
    const double cf_i = (i == n) ? (cf_coupon + 1.0) : cf_coupon;
    const double disc = std::pow(v, time_i);
    if (std::isnan(disc) || std::isinf(disc)) {
      return ErrorCode::Num;
    }
    const double pv = cf_i * disc;
    num += time_i * pv;
    den += pv;
  }

  if (den == 0.0) {
    return ErrorCode::Num;
  }
  const double duration = (num / den) / freq_d;
  if (std::isnan(duration) || std::isinf(duration)) {
    return ErrorCode::Num;
  }
  return duration;
}

}  // namespace

// --- DURATION(settlement, maturity, coupon, yld, frequency, [basis=0]) -
//
// Macaulay duration in years.
Value Duration(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto d = compute_macaulay(args, arity);
  if (!d) {
    return Value::error(d.error());
  }
  return finalize(d.value());
}

// --- MDURATION(settlement, maturity, coupon, yld, frequency, [basis=0]) -
//
// Modified duration: DURATION / (1 + yld/frequency). Re-reads the yld and
// frequency arguments because `compute_macaulay` does not surface them; the
// argument coercions are idempotent (any error inside an arg has already
// been short-circuited by the dispatcher's `propagate_errors = true`
// path).
Value MDuration(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto d = compute_macaulay(args, arity);
  if (!d) {
    return Value::error(d.error());
  }
  // yld and frequency are guaranteed to be valid here — `compute_macaulay`
  // succeeded, so it already validated them. We only need their numeric
  // values to form the modifier.
  auto yld_e = read_required_number(args, 3);
  if (!yld_e) {
    return Value::error(yld_e.error());
  }
  auto freq_e = read_required_number(args, 4);
  if (!freq_e) {
    return Value::error(freq_e.error());
  }
  const double freq_d = std::trunc(freq_e.value());
  const double modifier = 1.0 + yld_e.value() / freq_d;
  if (modifier == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(d.value() / modifier);
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
