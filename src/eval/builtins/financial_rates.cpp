// Copyright 2026 libraz. Licensed under the MIT License.
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
//   * T-Bill entries additionally require `maturity - settlement <= 365`
//     (a T-Bill cannot mature more than one year after settlement).
//
// DISC / INTRATE / RECEIVED reuse the day-count helpers in
// `eval/date_time.h` (moved there from `datetime.cpp` so the financial
// family can share them without pulling in the whole calendar module).
// TBILLPRICE / TBILLYIELD / TBILLEQ use an actual-day / 360-day basis
// directly (Excel fixes the convention — no `basis` argument).

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/coerce.h"
#include "eval/date_time.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Computes YEARFRAC(settlement, maturity, basis) under the same rules as
// the YEARFRAC builtin. Returns `#NUM!` for an unsupported basis (or if
// the helper would yield a non-finite / zero value — the latter would
// otherwise divide to infinity in the callers).
Expected<double, ErrorCode> yearfrac(double settlement, double maturity, int basis) {
  if (basis < 0 || basis > 4) {
    return ErrorCode::Num;
  }
  const double s = std::trunc(settlement);
  const double e = std::trunc(maturity);
  const date_time::YMD a = date_time::ymd_from_serial(s);
  const date_time::YMD b = date_time::ymd_from_serial(e);
  double yf = 0.0;
  switch (basis) {
    case 0:
      yf = date_time::yearfrac_us30_360(a.y, a.m, a.d, b.y, b.m, b.d);
      break;
    case 1:
      yf = date_time::yearfrac_actual_actual(a.y, a.m, a.d, b.y, b.m, b.d);
      break;
    case 2:
      yf = (e - s) / 360.0;
      break;
    case 3:
      yf = (e - s) / 365.0;
      break;
    case 4:
      yf = date_time::yearfrac_eu30_360(a.y, a.m, a.d, b.y, b.m, b.d);
      break;
    default:
      return ErrorCode::Num;
  }
  if (std::isnan(yf) || std::isinf(yf) || yf <= 0.0) {
    return ErrorCode::Num;
  }
  return yf;
}

// Reads the `basis` argument at `args[index]` if present, otherwise
// returns the default (0). Truncates toward zero and validates against
// the set {0, 1, 2, 3, 4}.
Expected<int, ErrorCode> read_basis(const Value* args, std::uint32_t arity, std::uint32_t index) {
  if (arity <= index) {
    return 0;
  }
  auto raw = read_required_number(args, index);
  if (!raw) {
    return raw.error();
  }
  const int basis = static_cast<int>(std::trunc(raw.value()));
  if (basis < 0 || basis > 4) {
    return ErrorCode::Num;
  }
  return basis;
}

// Reads a required date argument, truncating toward zero. Negative serials
// are rejected as `#NUM!` (Excel's calendar builtins do the same).
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
  auto settlement = read_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto pr = read_required_number(args, 2);
  if (!pr) {
    return Value::error(pr.error());
  }
  auto redemption = read_required_number(args, 3);
  if (!redemption) {
    return Value::error(redemption.error());
  }
  auto basis = read_basis(args, arity, 4);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (settlement.value() >= maturity.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (pr.value() <= 0.0 || redemption.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = yearfrac(settlement.value(), maturity.value(), basis.value());
  if (!yf) {
    return Value::error(yf.error());
  }
  const double result = ((redemption.value() - pr.value()) / redemption.value()) / yf.value();
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
  auto settlement = read_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto investment = read_required_number(args, 2);
  if (!investment) {
    return Value::error(investment.error());
  }
  auto redemption = read_required_number(args, 3);
  if (!redemption) {
    return Value::error(redemption.error());
  }
  auto basis = read_basis(args, arity, 4);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (settlement.value() >= maturity.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (investment.value() <= 0.0 || redemption.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = yearfrac(settlement.value(), maturity.value(), basis.value());
  if (!yf) {
    return Value::error(yf.error());
  }
  const double result = ((redemption.value() - investment.value()) / investment.value()) / yf.value();
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
  auto settlement = read_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto investment = read_required_number(args, 2);
  if (!investment) {
    return Value::error(investment.error());
  }
  auto disc_rate = read_required_number(args, 3);
  if (!disc_rate) {
    return Value::error(disc_rate.error());
  }
  auto basis = read_basis(args, arity, 4);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (settlement.value() >= maturity.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (investment.value() <= 0.0 || disc_rate.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = yearfrac(settlement.value(), maturity.value(), basis.value());
  if (!yf) {
    return Value::error(yf.error());
  }
  const double denom = 1.0 - disc_rate.value() * yf.value();
  if (denom <= 0.0) {
    // Excel returns #NUM! when the discount consumes the whole face
    // value (result would be infinite or negative).
    return Value::error(ErrorCode::Num);
  }
  const double result = investment.value() / denom;
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
//   - maturity > settlement + 365   ->  #NUM! (T-Bill < 1 year)
//   - discount <= 0                 ->  #NUM!
//   - result <= 0 (discount * DSM/360 >= 1)  ->  #NUM!
Value TBillPrice(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto settlement = read_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto discount = read_required_number(args, 2);
  if (!discount) {
    return Value::error(discount.error());
  }
  const double dsm = maturity.value() - settlement.value();
  if (dsm <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (dsm > 365.0) {
    return Value::error(ErrorCode::Num);
  }
  if (discount.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = 100.0 * (1.0 - discount.value() * dsm / 360.0);
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
//   - maturity > settlement + 365  ->  #NUM!
//   - pr <= 0                      ->  #NUM!
Value TBillYield(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto settlement = read_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto pr = read_required_number(args, 2);
  if (!pr) {
    return Value::error(pr.error());
  }
  const double dsm = maturity.value() - settlement.value();
  if (dsm <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (dsm > 365.0) {
    return Value::error(ErrorCode::Num);
  }
  if (pr.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = ((100.0 - pr.value()) / pr.value()) * (360.0 / dsm);
  return finalize(result);
}

// --- TBILLEQ(settlement, maturity, discount) ---------------------------
//
// Bond-equivalent yield for a T-Bill. Two branches, split at DSM = 182:
//
//   if DSM <= 182:
//       TBILLEQ = (365 * discount) / (360 - discount * DSM)
//
//   if 182 < DSM <= 365:
//       price = 1 - discount * DSM / 360
//       A = DSM / 365
//       discriminant = (2A - 1)^2 - (2A^2 - 1) * (1 - 1/price) * 4A
//       TBILLEQ = (-2A + 1 + sqrt(discriminant)) / (2A - 1)
//
// The long branch solves the quadratic arising from semi-annual
// compounding on the bond-equivalent side; see Microsoft's TBILLEQ
// documentation. A negative discriminant or non-positive price surfaces
// as `#NUM!` (Excel's behaviour when the inputs are incompatible).
//
// Domain:
//   - settlement >= maturity       ->  #NUM!
//   - maturity > settlement + 365  ->  #NUM!
//   - discount <= 0                ->  #NUM!
Value TBillEq(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto settlement = read_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto discount = read_required_number(args, 2);
  if (!discount) {
    return Value::error(discount.error());
  }
  const double dsm = maturity.value() - settlement.value();
  if (dsm <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (dsm > 365.0) {
    return Value::error(ErrorCode::Num);
  }
  if (discount.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double rate = discount.value();
  if (dsm <= 182.0) {
    const double denom = 360.0 - rate * dsm;
    if (denom == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    return finalize((365.0 * rate) / denom);
  }
  // Long branch: 182 < DSM <= 365. Semi-annual-compounding quadratic.
  const double price = 1.0 - rate * dsm / 360.0;
  if (price <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double a = dsm / 365.0;
  const double disc = (2.0 * a - 1.0) * (2.0 * a - 1.0) - (2.0 * a * a - 1.0) * (1.0 - 1.0 / price) * 4.0 * a;
  if (disc < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double denom = 2.0 * a - 1.0;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (-2.0 * a + 1.0 + std::sqrt(disc)) / denom;
  return finalize(result);
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
