//
// Implementation of the eager accrued-interest built-ins: ACCRINT and
// ACCRINTM. Registered from `financial.cpp` via
// `register_financial_builtins`.
//
// ACCRINTM is simple interest, `par * rate * YEARFRAC(issue, settlement,
// basis)`. ACCRINT instead counts coupon periods on the quasi-coupon grid
// anchored on `first_interest`; see `accrint_periods` for the measured rule.

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

// Quasi-coupon dates anchored on `first_interest`: `at(k)` lies `k` periods
// before the anchor (negative `k` steps forward). A month-end anchor keeps
// every quasi date on its month's end (2024-02-29 -> 2023-08-31).
struct QuasiCouponGrid {
  date_time::YMD anchor;
  unsigned anchor_day;
  int step_months;
  bool date1904;

  double at(int k) const noexcept {
    int y = anchor.y;
    unsigned m = anchor.m;
    shift_months(y, m, -k * step_months);
    return quasi_serial(y, m, anchor_day, date1904);
  }
};

// Length of the quasi-coupon period [start, end]: actual days for basis 1,
// otherwise the nominal 360/freq (365/freq for basis 3).
double quasi_period_length(double start, double end, int basis, int frequency) noexcept {
  if (basis == 1) {
    return end - start;
  }
  return (basis == 3 ? 365.0 : 360.0) / static_cast<double>(frequency);
}

// Accrued interest of ACCRINT in coupon units (multiples of
// par * rate / frequency), as Mac Excel 365 computes it. With quasi dates
// q(k) counted back from first_interest = q(0) (k <= 0 lies after it),
// issue in (q(n), q(n-1)], D the basis day count and E = length of
// (q(1), q(0)]:
//   - calc_method TRUE with settlement after first_interest walks forward:
//     the issue part D(issue, q(n-1)) / len(issue period) and every whole
//     period count in full, and the period holding settlement counts
//     D(start, settlement) / E; issue and settlement in one period give
//     D(issue, settlement) / E.
//   - otherwise n <= 1: D(issue, settlement) / E.
//   - otherwise base = issue part + D(q(1), settlement) / E, plus n - 2
//     whole periods when calc_method is TRUE. The q(1) leg is signed, so
//     settlement before q(1) gives FALSE its negative values.
double accrint_periods(double issue, double first_interest, double settlement, int frequency, int basis,
                       bool calc_method, bool date1904) noexcept {
  const date_time::YMD anchor = date_time::ymd_from_serial(first_interest, date1904);
  const bool month_end = anchor.d == date_time::days_in_month(anchor.y, anchor.m);
  const QuasiCouponGrid q{anchor, month_end ? 31u : anchor.d, 12 / frequency, date1904};
  const auto days = [&](double a, double b) {
    return std::round(date_time::basis_days_between(a, b, basis, date1904));
  };

  int n = 1;
  while (q.at(n) >= issue) {
    ++n;
  }
  while (q.at(n - 1) < issue) {
    --n;
  }
  const double e_last = quasi_period_length(q.at(1), q.at(0), basis, frequency);
  const double issue_part = days(issue, q.at(n - 1)) / quasi_period_length(q.at(n), q.at(n - 1), basis, frequency);

  if (calc_method && settlement > first_interest) {
    int k = n;
    while (q.at(k - 1) < settlement) {
      --k;
    }
    if (k == n) {
      return days(issue, settlement) / e_last;
    }
    return issue_part + static_cast<double>(n - 1 - k) + days(q.at(k), settlement) / e_last;
  }
  if (n <= 1) {
    return days(issue, settlement) / e_last;
  }
  const double from_q1 = issue_part + days(q.at(1), settlement) / e_last;
  return calc_method ? from_q1 + static_cast<double>(n - 2) : from_q1;
}

}  // namespace

// --- ACCRINT(issue, first_interest, settlement, rate, par, frequency,
//             [basis=0], [calc_method=TRUE]) --------------------------------
//
//   ACCRINT = par * rate / frequency * accrint_periods(...)
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
  return finalize(par * rate / static_cast<double>(frequency.value()) * periods);
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
