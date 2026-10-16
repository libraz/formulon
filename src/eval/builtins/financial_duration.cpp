//
// Implementation of the Macaulay-duration / modified-duration bond
// built-ins:
//
//   * DURATION(settlement, maturity, coupon, yld, frequency, [basis=0])
//   * MDURATION(settlement, maturity, coupon, yld, frequency, [basis=0])
//
// Both share the closed-form formula, evaluated term by term in this order:
//
//   f = (E - A) / E                  (E = coupon period days, A = days since PCD)
//   q = yld/frequency + 1,  p = (f + n) - 1
//   t_k = (k - 1) + f,  c_k = ((100 * coupon) / frequency) / POWER(q, t_k)
//
//   num = (p * 100) / POWER(q, p) + sum_k c_k * t_k
//   den = 100 / POWER(q, p)       + sum_k c_k
//
//   DURATION  = (num / den) / frequency
//   MDURATION = DURATION / (1 + yld/frequency)
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

// Computes Macaulay duration. Returns `#NUM!` on any
// validation or numerical failure. The caller (DURATION) returns this
// value directly; MDURATION divides it by `(1 + yld/frequency)`.
Expected<double, ErrorCode> compute_macaulay(const Value* args, std::uint32_t arity, bool date1904) {
  double v4[4];
  if (auto read = read_required_numbers(args, "ddnn", v4); !read) {
    return std::move(read.error());
  }
  const double settlement = v4[0];
  const double maturity = v4[1];
  const double coupon = v4[2];
  const double yld = v4[3];
  auto frequency_e = read_coupon_frequency(args, 4);
  if (!frequency_e) {
    return std::move(frequency_e.error());
  }
  auto basis_e = read_day_count_basis(args, arity, 5);
  if (!basis_e) {
    return std::move(basis_e.error());
  }

  // Validation order matches Microsoft's documented contract and the
  // sibling COUP* / PRICE* impls: date ordering -> frequency domain ->
  // basis domain -> coupon / yld sign checks.
  if (settlement >= maturity) {
    return ErrorCode::Num;
  }
  const int frequency = frequency_e.value();
  const int basis = basis_e.value();
  if (coupon < 0.0) {
    return ErrorCode::Num;
  }
  if (yld < 0.0) {
    return ErrorCode::Num;
  }

  CouponDates cd{};
  if (!compute_coupon_dates(settlement, maturity, frequency, basis, date1904, &cd)) {
    return ErrorCode::Num;
  }
  if (cd.coupons_remaining <= 0 || cd.period_days <= 0.0) {
    return ErrorCode::Num;
  }

  const double freq_d = static_cast<double>(frequency);
  const double e = cd.period_days;
  const double f = (e - cd.days_bs) / e;
  const std::int32_t n = cd.coupons_remaining;
  const double p = (f + static_cast<double>(n)) - 1.0;
  const double q = (yld / freq_d) + 1.0;
  const double p1 = excel_power(q, p);

  double sum_num = 0.0;
  double sum_den = 0.0;
  for (std::int32_t k = 1; k <= n; ++k) {
    const double t = (static_cast<double>(k) - 1.0) + f;
    const double c = ((100.0 * coupon) / freq_d) / excel_power(q, t);
    sum_num = (c * t) + sum_num;
    sum_den = c + sum_den;
  }
  const double num = ((p * 100.0) / p1) + sum_num;
  const double den = (100.0 / p1) + sum_den;
  if (std::isnan(num) || std::isnan(den) || den == 0.0) {
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
Value Duration(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto d = compute_macaulay(args, arity, date1904);
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
Value MDuration(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto d = compute_macaulay(args, arity, date1904);
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
