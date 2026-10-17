//
// Implementation of the eager security-rate / T-Bill built-ins:
// DISC, INTRATE, RECEIVED, TBILLPRICE, TBILLYIELD, TBILLEQ. Registered
// from `financial.cpp` via `register_financial_builtins`.
//
// All six share the following conventions:
//   * Date arguments are Excel serial numbers (doubles); the integer part
//     is taken with `std::trunc` before use.
//   * `settlement < maturity` is required; otherwise `#NUM!`.
//   * `basis` (where present) must be in {0, 1, 2, 3, 4} after truncation.
//   * T-Bill entries additionally require that maturity falls no more
//     than one *calendar* year after settlement (Excel uses the
//     "same month + day next year" anniversary rule, not a day-count
//     cutoff). A span that crosses the anniversary serial by even one
//     day yields `#NUM!`, even if the raw DSM is 366.
//
// DISC / INTRATE / RECEIVED reuse the day-count helpers in
// `eval/date_time.h` (moved there from `datetime.cpp` so the financial
// family can share them without pulling in the whole calendar module).
// TBILLPRICE / TBILLYIELD / TBILLEQ use an actual-day / 360-day basis
// directly (Excel fixes the convention — no `basis` argument).

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

struct TBillArgs {
  double value;
  double dsm;
};

bool has_direct_bool_tbill_arg(const Value* args) {
  return args[0].kind() == ValueKind::Bool || args[1].kind() == ValueKind::Bool || args[2].kind() == ValueKind::Bool;
}

Expected<double, ErrorCode> t_bill_dsm(double settlement, double maturity, bool date1904) {
  const double dsm = maturity - settlement;
  if (dsm <= 0.0) {
    return ErrorCode::Num;
  }
  // Calendar-year rule: reject maturities past the "same month + day next
  // year" anniversary of settlement, regardless of raw day count. This
  // matches Excel 365 (and is stricter than the naive `dsm > 366` rule,
  // which accepted 366-day spans that actually cross the anniversary).
  const auto s_ymd = date_time::ymd_from_serial(std::floor(settlement), date1904);
  const double anniversary = date_time::serial_from_ymd(s_ymd.y + 1, s_ymd.m, s_ymd.d, date1904);
  if (std::floor(maturity) > anniversary) {
    return ErrorCode::Num;
  }
  return dsm;
}

Expected<TBillArgs, ErrorCode> read_tbill_args(const Value* args, bool date1904) {
  // Excel-quirk: T-Bill functions reject a direct Bool for settlement,
  // maturity, or the numeric rate/price argument with `#VALUE!` rather
  // than coercing TRUE/FALSE to 1/0. See the DEC2BIN precedent in
  // `engineering.cpp::convert_from_dec`.
  if (has_direct_bool_tbill_arg(args)) {
    return ErrorCode::Value;
  }
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return std::move(settlement.error());
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return std::move(maturity.error());
  }
  auto value = read_required_number(args, 2);
  if (!value) {
    return std::move(value.error());
  }
  auto dsm = t_bill_dsm(settlement.value(), maturity.value(), date1904);
  if (!dsm) {
    return std::move(dsm.error());
  }
  return TBillArgs{value.value(), dsm.value()};
}

}  // namespace

// --- DISC(settlement, maturity, pr, redemption, [basis=0]) -------------
//
// Discount rate for a security that doesn't pay periodic interest:
//
//   DISC = (1 - pr / redemption) * (B / DSM)
//
// where DSM is the day count from settlement to maturity and B the year length of the basis.
//
// Domain per Microsoft docs:
//   - settlement >= maturity  ->  #NUM!
//   - pr <= 0 or redemption <= 0  ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}  ->  #NUM!
Value Disc(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [settlement, maturity, pr, redemption, basis] = parsed.value();
  auto span = day_count_span(settlement, maturity, basis, date1904);
  if (!span) {
    return Value::error(span.error());
  }
  if (span.value().days <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (1.0 - pr / redemption) * (span.value().year / span.value().days);
  return finalize(result);
}

// --- INTRATE(settlement, maturity, investment, redemption, [basis=0]) --
//
// Interest rate for a fully invested security:
//
//   INTRATE = (B / DSM) * ((redemption - investment) / investment)
//
// Domain per Microsoft docs:
//   - settlement >= maturity  ->  #NUM!
//   - investment <= 0 or redemption <= 0  ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}  ->  #NUM!
Value Intrate(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [settlement, maturity, investment, redemption, basis] = parsed.value();
  auto span = day_count_span(settlement, maturity, basis, date1904);
  if (!span) {
    return Value::error(span.error());
  }
  if (span.value().days <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (span.value().year / span.value().days) * ((redemption - investment) / investment);
  return finalize(result);
}

// --- RECEIVED(settlement, maturity, investment, discount, [basis=0]) ---
//
// Amount received at maturity for a fully invested security:
//
//   RECEIVED = investment / (1 - (discount * DSM) / B)
//
// Domain per Microsoft docs:
//   - settlement >= maturity  ->  #NUM!
//   - investment <= 0 or discount <= 0  ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}  ->  #NUM!
//   - (discount * DSM) / B >= 1  ->  #NUM!
Value Received(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [settlement, maturity, investment, disc_rate, basis] = parsed.value();
  auto span = day_count_span(settlement, maturity, basis, date1904);
  if (!span) {
    return Value::error(span.error());
  }
  const double x = (disc_rate * span.value().days) / span.value().year;
  if (!(x < 1.0)) {
    // The discount consumes the whole face value (result would be infinite or negative).
    return Value::error(ErrorCode::Num);
  }
  const double result = investment / (-x + 1.0);
  return finalize(result);
}

// --- TBILLPRICE(settlement, maturity, discount) ------------------------
//
// Price per $100 face value of a T-Bill. Uses an actual-day / 360-day
// basis (Excel fixes this — no `basis` argument):
//
//   DSM = maturity - settlement  (actual days)
//   TBILLPRICE = 100 * (1 - discount * DSM / 360)
//
// Domain per Microsoft docs:
//   - settlement >= maturity        ->  #NUM!
//   - maturity > settlement + 1 calendar year  ->  #NUM!
//     (Excel checks against the same-month/day anniversary serial, so
//     e.g. 1902-09-26 -> 1903-09-27 is rejected even though DSM = 366.)
//   - discount <= 0                 ->  #NUM!
//   - result <= 0 (discount * DSM/360 >= 1)  ->  #NUM!
Value TBillPrice(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/, bool date1904) {
  auto parsed = read_tbill_args(args, date1904);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [discount, dsm] = parsed.value();
  if (discount <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = 100.0 * (1.0 - discount * dsm / 360.0);
  if (result <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(result);
}

// --- TBILLYIELD(settlement, maturity, pr) ------------------------------
//
// Yield for a T-Bill at price `pr` (per 100 face value):
//
//   DSM = maturity - settlement
//   TBILLYIELD = ((100 - pr) / pr) * (360 / DSM)
//
// Domain per Microsoft docs:
//   - settlement >= maturity       ->  #NUM!
//   - maturity > settlement + 1 calendar year  ->  #NUM!
//     (same anniversary rule as TBILLPRICE).
//   - pr <= 0                      ->  #NUM!
Value TBillYield(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/, bool date1904) {
  auto parsed = read_tbill_args(args, date1904);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [pr, dsm] = parsed.value();
  if (pr <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = ((100.0 - pr) / pr) * (360.0 / dsm);
  return finalize(result);
}

// --- TBILLEQ(settlement, maturity, discount) ---------------------------
//
// Bond-equivalent yield for a T-Bill. Two branches, split at DSM = 182:
//
//   if DSM <= 182:
//       TBILLEQ = (365 * discount) / (360 - discount * DSM)
//
//   if DSM > 182:
//       a = DSM / Y            (Y = 366 when DSM == 366, else 365)
//       price = 1 - discount * DSM / 360
//       disc = a^2 - (2a - 1) * (1 - 1/price)
//       TBILLEQ = (-a + sqrt(disc)) / (a - 0.5)
//   evaluated in the operation order given at the long branch below.
//
// The long branch solves the quadratic arising from semi-annual
// compounding on the bond-equivalent side; see Microsoft's TBILLEQ
// documentation. A negative discriminant or non-positive price surfaces
// as `#NUM!` (Excel's behaviour when the inputs are incompatible).
//
// Domain:
//   - settlement >= maturity       ->  #NUM!
//   - maturity > settlement + 1 calendar year  ->  #NUM!
//     (same anniversary rule as TBILLPRICE / TBILLYIELD).
//   - discount <= 0                ->  #NUM!
Value TBillEq(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/, bool date1904) {
  auto parsed = read_tbill_args(args, date1904);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [discount, dsm] = parsed.value();
  if (discount <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double rate = discount;
  if (dsm <= 182.0) {
    const double denom = 360.0 - rate * dsm;
    if (denom == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    return finalize((365.0 * rate) / denom);
  }
  // Long branch: 182 < DSM. The semi-annual-compounding quadratic, with
  // Y = 366 only when DSM is exactly 366 and Excel's operation order:
  //
  //   p    = ((rate * 100) * DSM) / 360       P = (100 - p) / 100
  //   A    = ((-2) * DSM) / Y                 Bq = ((2 * DSM) / Y - 1) * (1 - 1 / P)
  //   disc = (DSM / Y)^2 - Bq
  //   TBILLEQ = ((sqrt(disc) * 2) + A) / ((2 * DSM) / Y - 1)
  const double year = dsm == 366.0 ? 366.0 : 365.0;
  const double p = ((rate * 100.0) * dsm) / 360.0;
  const double price_fraction = (100.0 - p) / 100.0;
  if (price_fraction <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double a = ((-2.0) * dsm) / year;
  const double bq = (((2.0 * dsm) / year) - 1.0) * (1.0 - 1.0 / price_fraction);
  const double q = dsm / year;
  const double disc = q * q - bq;
  if (disc < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double denom = ((2.0 * dsm) / year) - 1.0;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = ((std::sqrt(disc) * 2.0) + a) / denom;
  return finalize(result);
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
