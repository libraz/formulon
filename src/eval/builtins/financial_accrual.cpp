//
// Implementation of the eager accrued-interest built-ins: ACCRINT and
// ACCRINTM. Registered from `financial.cpp` via
// `register_financial_builtins`.
//
// ACCRINTM is simple interest, `par * rate * YEARFRAC(issue, settlement,
// basis)`. ACCRINT instead counts coupon periods on the quasi-coupon grid
// anchored on `first_interest`; see `accrint_periods`.

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/coerce.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Reads a Bool / numeric "calc_method" tail argument. TRUE/non-zero
// selects the issue-to-settlement formula; FALSE/zero selects the
// first-interest-to-settlement branch. Missing -> TRUE (Excel default).
Expected<bool, ErrorCode> read_calc_method(const Value* args, std::uint32_t arity, std::uint32_t index) {
  if (arity <= index) {
    return true;
  }
  const Value& v = args[index];
  if (v.is_boolean()) {
    return v.as_boolean();
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return std::move(coerced.error());
  }
  return coerced.value() != 0.0;
}

// `serial` moved by `months` calendar months; a month-end schedule keeps every
// date on its month's end, otherwise the day clamps to the shorter month.
double add_months(double serial, int months, bool month_end, bool date1904) noexcept {
  const date_time::YMD d = date_time::ymd_from_serial(serial, date1904);
  int y = d.y;
  unsigned m = d.m;
  date_time::shift_year_month(y, m, months);
  return date_time::serial_from_ymd_clamped(y, m, month_end ? 31u : d.d, date1904);
}

// Day count of `[start, end]` for the basis: actual for basis 1, basis 2 when
// `actual` and basis 3, European 30/360 for basis 4, US 30/360 otherwise.
double accrint_days(double start, double end, int basis, bool actual, int mode, bool date1904) noexcept {
  if (basis == 1 || (basis == 2 && actual) || (basis == 3 && actual)) {
    return end - start;
  }
  return days_30_360(start, end, mode, basis != 4, date1904);
}

// Length of the coupon period starting at `pcd` on the schedule counted back
// from `anchor`: 360/freq (365/freq for basis 3), or for basis 1 the actual
// span of the period holding `pcd`.
double accrint_period_length(double anchor, double pcd, int frequency, int basis, bool month_end,
                             bool date1904) noexcept {
  const double freq = static_cast<double>(frequency);
  if (basis == 3) {
    return 365.0 / freq;
  }
  if (basis != 1) {
    return 360.0 / freq;
  }
  double current = anchor;
  while (true) {
    const double previous = add_months(current, -(12 / frequency), month_end, date1904);
    if (previous <= pcd) {
      return current - previous;
    }
    current = previous;
  }
}

// Accrued interest of ACCRINT in coupon units (multiples of par * rate /
// frequency). The period holding settlement contributes A / E; walking back
// from its start, every period before issue's contributes `calc_method` (1 for
// TRUE, 0 for FALSE) and the period holding issue contributes its fraction.
// With calc_method TRUE and settlement after first_interest the settlement
// period is found by walking forward from first_interest.
double accrint_periods(double issue, double first_interest, double settlement, int frequency, int basis,
                       bool calc_method, bool date1904) noexcept {
  const int months = 12 / frequency;
  const date_time::YMD anchor = date_time::ymd_from_serial(first_interest, date1904);
  const bool month_end = anchor.d == date_time::days_in_month(anchor.y, anchor.m);
  double pcd = 0.0;
  if (settlement > first_interest && calc_method) {
    double current = first_interest;
    pcd = first_interest;
    while (current < settlement) {
      pcd = current;
      current = add_months(current, months, month_end, date1904);
    }
  } else {
    pcd = add_months(first_interest, -months, month_end, date1904);
  }
  const bool issue_after = issue > pcd;
  const double start = issue_after ? issue : pcd;
  const double a = accrint_days(start, settlement, basis, true, issue_after ? 0 : 1, date1904);
  const double e = accrint_period_length(first_interest, pcd, frequency, basis, month_end, date1904);
  double total = a / e;
  while (pcd > issue) {
    const double old = pcd;
    pcd = add_months(old, -months, month_end, date1904);
    const bool partial = issue > pcd;
    const double period_start = partial ? issue : pcd;
    const double ai = accrint_days(period_start, old, basis, true, partial ? 2 : 3, date1904);
    const double ei =
        basis == 3 ? 365.0 / static_cast<double>(frequency) : accrint_days(pcd, old, basis, false, 3, date1904);
    total = total + (partial ? ai / ei : (calc_method ? 1.0 : 0.0));
  }
  return total;
}

}  // namespace

// --- ACCRINT(issue, first_interest, settlement, rate, par, frequency,
//             [basis=0], [calc_method=TRUE]) --------------------------------
//
//   ACCRINT = (par * (rate / frequency)) * accrint_periods(...)
//
// Domain:
//   - issue >= settlement            ->  #NUM!
//   - rate <= 0                      ->  #NUM!
//   - par <= 0                       ->  #NUM!
//   - frequency not in {1, 2, 4}     ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}   ->  #NUM!
Value Accrint(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  double v[5];
  if (auto read = read_required_numbers(args, "dddnn", v); !read) {
    return Value::error(read.error());
  }
  const double issue = std::trunc(v[0]);
  const double first_interest = std::trunc(v[1]);
  const double settlement = std::trunc(v[2]);
  const double rate = v[3];
  const double par = v[4];
  auto frequency = read_coupon_frequency(args, 5);
  if (!frequency) {
    return Value::error(frequency.error());
  }
  auto basis = read_day_count_basis(args, arity, 6);
  if (!basis) {
    return Value::error(basis.error());
  }
  auto calc_method = read_calc_method(args, arity, 7);
  if (!calc_method) {
    return Value::error(calc_method.error());
  }
  if (issue >= settlement) {
    return Value::error(ErrorCode::Num);
  }
  if (rate <= 0.0 || par <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double periods = accrint_periods(issue, first_interest, settlement, frequency.value(), basis.value(),
                                         calc_method.value(), date1904);
  return finalize((par * (rate / static_cast<double>(frequency.value()))) * periods);
}

// --- ACCRINTM(issue, settlement, rate, par, [basis=0]) -----------------
//
// Accrued interest at maturity for a security that pays interest only at
// maturity:
//
//   ACCRINTM = par * rate * YEARFRAC(issue, settlement, basis)
//
// Domain:
//   - issue >= settlement            ->  #NUM!
//   - rate <= 0                      ->  #NUM!
//   - par <= 0                       ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}   ->  #NUM!
Value Accrintm(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [issue, settlement, rate, par, basis] = parsed.value();
  auto yf = yearfrac_for_basis(issue, settlement, basis, date1904);
  if (!yf) {
    return Value::error(yf.error());
  }
  return finalize(par * rate * yf.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
