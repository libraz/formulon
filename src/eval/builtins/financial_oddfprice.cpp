//
// Implementation of the irregular-first-period bond-pricing built-in:
//
//   * ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld,
//               redemption, frequency, [basis=0])
//
// Returns the clean price per 100 face. The schedule facts and the short /
// long first-period price forms are documented in `financial_oddf_helpers.h`;
// this file resolves the schedule once per call (`read_odd_first_bond`) and
// evaluates the price at any yield (`oddf_price_at`), which ODDFYIELD iterates.

#include "eval/builtins/financial_oddfprice.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddf_helpers.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// End-of-month code of a schedule date: 1 for a last day, the day itself for
// the 29th and 30th, otherwise 0 (plain day, clamped to short months).
unsigned eom_code(double serial, bool date1904) noexcept {
  const date_time::YMD d = date_time::ymd_from_serial(serial, date1904);
  if (d.d == date_time::days_in_month(d.y, d.m)) {
    return 1u;
  }
  return d.d >= 29u ? d.d : 0u;
}

// `serial` moved by `months` calendar months under an end-of-month code.
double step_months(double serial, int months, unsigned code, bool date1904) noexcept {
  const date_time::YMD d = date_time::ymd_from_serial(serial, date1904);
  int y = d.y;
  unsigned m = d.m;
  date_time::shift_year_month(y, m, months);
  const unsigned day = code == 1u ? 31u : (code >= 29u ? code : d.d);
  return date_time::serial_from_ymd_clamped(y, m, day, date1904);
}

// Day count of [a, b] for the basis (0 when `a >= b`); `mode` selects the 30/360 variant.
double days_mode(double a, double b, int basis, int mode, bool date1904) noexcept {
  if (!(a < b)) {
    return 0.0;
  }
  if (basis == 0 || basis == 4) {
    return days_30_360(a, b, mode, basis == 0, date1904);
  }
  return b - a;
}

// Coupon dates d on the schedule anchored at `b` (with b's code) satisfying a < d <= b.
std::int32_t coupon_count(double a, double b, int months, bool date1904) noexcept {
  const unsigned code = eom_code(b, date1904);
  std::int32_t n = 0;
  for (double cur = b; cur > a; cur = step_months(cur, -months, code, date1904)) {
    ++n;
  }
  return n;
}

struct QuasiPair {
  double after;
  double before;
};

// Walks back from `anchor` until a date at or before `bound`; returns the last date after it and that date.
QuasiPair quasi_walk(double anchor, double bound, int months, bool date1904) noexcept {
  const unsigned code = eom_code(anchor, date1904);
  double cur = anchor;
  while (true) {
    const double next = step_months(cur, -months, code, date1904);
    if (next <= bound) {
      return QuasiPair{cur, next};
    }
    cur = next;
  }
}

}  // namespace

Expected<OddFirstBond, ErrorCode> read_odd_first_bond(const Value* args, std::uint32_t arity,
                                                      bool slot5_must_be_positive, bool date1904) {
  double dates[4];
  if (auto read = read_required_numbers(args, "dddd", dates); !read) {
    return std::move(read.error());
  }
  auto tail = read_coupon_bond_tail(args, arity, 4);
  if (!tail) {
    return std::move(tail.error());
  }
  const auto [rate, slot5, redemption, frequency, basis] = tail.value();
  OddFirstBond b{};
  b.settlement = dates[0];
  b.maturity = dates[1];
  b.issue = dates[2];
  b.first_coupon = dates[3];
  // issue < settlement < first_coupon < maturity.
  if (b.issue >= b.settlement || b.settlement >= b.first_coupon || b.first_coupon >= b.maturity) {
    return ErrorCode::Num;
  }
  if (slot5_must_be_positive ? (slot5 <= 0.0) : (slot5 < 0.0)) {
    return ErrorCode::Num;
  }
  b.rate = rate;
  b.slot5 = slot5;
  b.redemption = redemption;
  b.freq_d = static_cast<double>(frequency);
  b.basis = basis;

  const int months = 12 / frequency;
  const double s = b.settlement;
  const double i = b.issue;
  const double f = b.first_coupon;
  // first_coupon must lie on the schedule anchored at maturity.
  {
    const unsigned code = eom_code(b.maturity, date1904);
    double cur = b.maturity;
    double next = step_months(cur, -months, code, date1904);
    while (next > f) {
      cur = next;
      next = step_months(cur, -months, code, date1904);
    }
    if (next != f) {
      return ErrorCode::Num;
    }
  }

  // E: the quasi period holding settlement, walking back from first_coupon.
  const QuasiPair holding = quasi_walk(f, s, months, date1904);
  if (basis == 1) {
    b.e = holding.after - holding.before;
  } else {
    b.e = (basis == 3 ? 365.0 : 360.0) / b.freq_d;
  }
  b.n = coupon_count(s, b.maturity, months, date1904);
  b.dfc = days_mode(i, f, basis, 2, date1904);
  b.dsc_first = days_mode(s, f, basis, 2, date1904);
  if (b.dfc == b.e) {
    return ErrorCode::Num;
  }
  b.long_first = b.dfc > b.e;
  if (!b.long_first) {
    b.accrued_days = days_mode(i, s, basis, 0, date1904);
    return b;
  }

  // Long first period: quasi periods stepped back from first_coupon with plain month steps.
  const std::int32_t quasi_count = coupon_count(i, f, months, date1904);
  double late = f;
  for (std::int32_t idx = quasi_count; idx > 0; --idx) {
    const double early = step_months(late, -months, 0u, date1904);
    const double nl = basis == 1 ? days_mode(early, late, basis, 3, date1904) : b.e;
    const double dcnl = idx >= 2 ? nl : days_mode(i, late, basis, 2, date1904);
    b.sum_dcnl = dcnl / nl + b.sum_dcnl;
    const double start = i > early ? i : early;
    const double end = s < late ? s : late;
    b.sum_a = days_mode(start, end, basis, 0, date1904) / nl + b.sum_a;
    late = early;
  }
  if (basis == 2 || basis == 3) {
    b.dsc = days_mode(s, holding.after, basis, 2, date1904);
  } else {
    b.dsc = b.e - days_mode(holding.before, s, basis, 2, date1904);
  }
  // Whole quasi periods from settlement's quasi date to first_coupon; the settlement day is
  // forced by first_coupon's end-of-month code.
  const unsigned code = eom_code(f, date1904);
  const double norm = step_months(s, 0, code, date1904);
  b.nq = s < norm ? 1.0 : 0.0;
  for (double d = step_months(norm, months, code, date1904); d < f; d = step_months(d, months, code, date1904)) {
    b.nq += 1.0;
  }
  b.n2 = coupon_count(f, b.maturity, months, date1904);
  return b;
}

Expected<double, ErrorCode> oddf_price_at(const OddFirstBond& b, double yld) {
  const double q = (yld / b.freq_d) + 1.0;
  const double cf = (100.0 * b.rate) / b.freq_d;
  double price = 0.0;
  if (!b.long_first) {
    const double t1 = b.dsc_first / b.e;
    const double redemption_term = b.redemption / excel_power(q, (static_cast<double>(b.n) - 1.0) + t1);
    const double first_term = ((cf * b.dfc) / b.e) / excel_power(q, t1);
    double regular = 0.0;
    for (std::int32_t k = 2; k <= b.n; ++k) {
      regular = (cf / excel_power(q, (static_cast<double>(k) - 1.0) + t1)) + regular;
    }
    const double accrued = ((b.accrued_days / b.e) * (b.rate / b.freq_d)) * 100.0;
    price = ((redemption_term + first_term) + regular) - accrued;
  } else {
    const double t = b.dsc / b.e;
    const double redemption_term = b.redemption / excel_power(q, (t + b.nq) + static_cast<double>(b.n2));
    const double first_term = (cf * b.sum_dcnl) / excel_power(q, b.nq + t);
    double regular = 0.0;
    for (std::int32_t k = 1; k <= b.n2; ++k) {
      regular = (cf / excel_power(q, (static_cast<double>(k) + b.nq) + t)) + regular;
    }
    price = ((redemption_term + first_term) + regular) - cf * b.sum_a;
  }
  if (std::isnan(price) || std::isinf(price)) {
    return ErrorCode::Num;
  }
  return price;
}

Expected<double, ErrorCode> compute_oddf_clean_price(const Value* args, std::uint32_t arity, bool date1904) {
  auto bond = read_odd_first_bond(args, arity, /*slot5_must_be_positive=*/false, date1904);
  if (!bond) {
    return std::move(bond.error());
  }
  return oddf_price_at(bond.value(), bond.value().slot5);
}

// --- ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld,
//              redemption, frequency, [basis=0]) ----------------------------
//
// Clean price per 100 face for a security whose first coupon period
// is irregular.
Value OddfPrice(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto p = compute_oddf_clean_price(args, arity, date1904);
  if (!p) {
    return Value::error(p.error());
  }
  return finalize(p.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
