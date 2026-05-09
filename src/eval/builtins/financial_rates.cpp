// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

struct SecurityRateArgs {
  double settlement;
  double maturity;
  double amount1;
  double amount2;
  int basis;
};

struct TBillArgs {
  double value;
  double dsm;
};

// Computes YEARFRAC(settlement, maturity, basis) under the same rules as
// the YEARFRAC builtin. Returns `#NUM!` for an unsupported basis (or if
// the helper would yield a non-finite / zero value — the latter would
// otherwise divide to infinity in the callers).
Expected<double, ErrorCode> positive_yearfrac(double settlement, double maturity, int basis) {
  auto yf = yearfrac_for_basis(settlement, maturity, basis);
  if (!yf) {
    return yf.error();
  }
  if (yf.value() <= 0.0) {
    return ErrorCode::Num;
  }
  return yf.value();
}

Expected<SecurityRateArgs, ErrorCode> read_security_rate_args(const Value* args, std::uint32_t arity) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto amount1 = read_required_number(args, 2);
  if (!amount1) {
    return amount1.error();
  }
  auto amount2 = read_required_number(args, 3);
  if (!amount2) {
    return amount2.error();
  }
  auto basis = read_day_count_basis(args, arity, 4);
  if (!basis) {
    return basis.error();
  }
  if (settlement.value() >= maturity.value()) {
    return ErrorCode::Num;
  }
  return SecurityRateArgs{settlement.value(), maturity.value(), amount1.value(), amount2.value(), basis.value()};
}

bool has_direct_bool_tbill_arg(const Value* args) {
  return args[0].kind() == ValueKind::Bool || args[1].kind() == ValueKind::Bool || args[2].kind() == ValueKind::Bool;
}

Expected<double, ErrorCode> t_bill_dsm(double settlement, double maturity) {
  const double dsm = maturity - settlement;
  if (dsm <= 0.0) {
    return ErrorCode::Num;
  }
  // Calendar-year rule: reject maturities past the "same month + day next
  // year" anniversary of settlement, regardless of raw day count. This
  // matches Excel 365 (and is stricter than the naive `dsm > 366` rule,
  // which accepted 366-day spans that actually cross the anniversary).
  const auto s_ymd = date_time::ymd_from_serial(std::floor(settlement));
  const double anniversary = date_time::serial_from_ymd(s_ymd.y + 1, s_ymd.m, s_ymd.d);
  if (std::floor(maturity) > anniversary) {
    return ErrorCode::Num;
  }
  return dsm;
}

Expected<TBillArgs, ErrorCode> read_tbill_args(const Value* args) {
  // Excel-quirk: T-Bill functions reject a direct Bool for settlement,
  // maturity, or the numeric rate/price argument with `#VALUE!` rather
  // than coercing TRUE/FALSE to 1/0. See the DEC2BIN precedent in
  // `engineering.cpp::convert_from_dec`.
  if (has_direct_bool_tbill_arg(args)) {
    return ErrorCode::Value;
  }
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto value = read_required_number(args, 2);
  if (!value) {
    return value.error();
  }
  auto dsm = t_bill_dsm(settlement.value(), maturity.value());
  if (!dsm) {
    return dsm.error();
  }
  return TBillArgs{value.value(), dsm.value()};
}

}  // namespace

// --- DISC(settlement, maturity, pr, redemption, [basis=0]) -------------
//
// Discount rate for a security that doesn't pay periodic interest:
//
//   DISC = ((redemption - pr) / redemption) / YEARFRAC(settlement, maturity, basis)
//
// Domain per Microsoft docs:
//   - settlement >= maturity  ->  #NUM!
//   - pr <= 0 or redemption <= 0  ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}  ->  #NUM!
Value Disc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [settlement, maturity, pr, redemption, basis] = parsed.value();
  if (pr <= 0.0 || redemption <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = positive_yearfrac(settlement, maturity, basis);
  if (!yf) {
    return Value::error(yf.error());
  }
  const double result = ((redemption - pr) / redemption) / yf.value();
  return finalize(result);
}

// --- INTRATE(settlement, maturity, investment, redemption, [basis=0]) --
//
// Interest rate for a fully invested security:
//
//   INTRATE = ((redemption - investment) / investment) / YEARFRAC(settlement, maturity, basis)
//
// Domain per Microsoft docs:
//   - settlement >= maturity  ->  #NUM!
//   - investment <= 0 or redemption <= 0  ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}  ->  #NUM!
Value Intrate(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [settlement, maturity, investment, redemption, basis] = parsed.value();
  if (investment <= 0.0 || redemption <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = positive_yearfrac(settlement, maturity, basis);
  if (!yf) {
    return Value::error(yf.error());
  }
  const double result = ((redemption - investment) / investment) / yf.value();
  return finalize(result);
}

// --- RECEIVED(settlement, maturity, investment, discount, [basis=0]) ---
//
// Amount received at maturity for a fully invested security:
//
//   RECEIVED = investment / (1 - discount * YEARFRAC(settlement, maturity, basis))
//
// Domain per Microsoft docs:
//   - settlement >= maturity  ->  #NUM!
//   - investment <= 0 or discount <= 0  ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}  ->  #NUM!
//   - 1 - discount*yearfrac == 0 (or negative)  ->  #NUM!
Value Received(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [settlement, maturity, investment, disc_rate, basis] = parsed.value();
  if (investment <= 0.0 || disc_rate <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = positive_yearfrac(settlement, maturity, basis);
  if (!yf) {
    return Value::error(yf.error());
  }
  const double denom = 1.0 - disc_rate * yf.value();
  if (denom <= 0.0) {
    // Excel returns #NUM! when the discount consumes the whole face
    // value (result would be infinite or negative).
    return Value::error(ErrorCode::Num);
  }
  const double result = investment / denom;
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
Value TBillPrice(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto parsed = read_tbill_args(args);
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
Value TBillYield(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto parsed = read_tbill_args(args);
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
//       a = DSM / 365
//       price = 1 - discount * DSM / 360
//       disc = a^2 - (2a - 1) * (1 - 1/price)
//       TBILLEQ = (-a + sqrt(disc)) / (a - 0.5)
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
Value TBillEq(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto parsed = read_tbill_args(args);
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
  // Long branch: 182 < DSM. Semi-annual-compounding quadratic per the
  // Microsoft TBILLEQ documentation:
  //
  //   a = DSM / 365   (or DSM / 366 when DSM > 365, i.e. the span is
  //                   exactly a one-year maturity across a leap-year
  //                   anniversary — see D6 in the IronCalc oracle,
  //                   14640 (1940-02-16) -> 15006 (1941-02-16))
  //   price = 1 - rate * DSM / 360
  //   disc = a^2 - (2a - 1) * (1 - 1/price)
  //   TBILLEQ = (-a + sqrt(disc)) / (a - 0.5)
  const double a = dsm / (dsm > 365.0 ? 366.0 : 365.0);
  const double price = 1.0 - rate * dsm / 360.0;
  if (price <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double disc = a * a - (2.0 * a - 1.0) * (1.0 - 1.0 / price);
  if (disc < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double denom = a - 0.5;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (-a + std::sqrt(disc)) / denom;
  return finalize(result);
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
