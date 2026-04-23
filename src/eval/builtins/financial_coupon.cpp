// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the six coupon-date financial built-ins —
// COUPPCD, COUPNCD, COUPNUM, COUPDAYBS, COUPDAYSNC, COUPDAYS.
//
// All six share the common signature `(settlement, maturity,
// frequency, [basis=0])` and route through the shared
// `compute_coupon_dates` engine in `eval/coupon_schedule.h`. Each
// builtin only reads the relevant sub-field of the resulting
// `CouponDates` struct.

#include "eval/builtins/financial_coupon.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/coupon_schedule.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using financial_detail::finalize;
using financial_detail::read_optional_number;
using financial_detail::read_required_number;

// Shared argument-validation helper for the COUP* family. Truncates
// settlement / maturity to integer serials, validates the frequency
// / basis domains, enforces `settlement < maturity`, and runs the
// coupon-schedule engine. On success the caller receives a populated
// `CouponDates`; on any validation failure a `#NUM!` error is
// returned for the caller to forward to Excel.
Expected<CouponDates, ErrorCode> resolve_coupon(const Value* args, std::uint32_t arity) {
  auto s_e = read_required_number(args, 0);
  if (!s_e) {
    return s_e.error();
  }
  auto m_e = read_required_number(args, 1);
  if (!m_e) {
    return m_e.error();
  }
  auto f_e = read_required_number(args, 2);
  if (!f_e) {
    return f_e.error();
  }
  auto b_e = read_optional_number(args, arity, 3, 0.0);
  if (!b_e) {
    return b_e.error();
  }

  const double s = std::trunc(s_e.value());
  const double m = std::trunc(m_e.value());
  if (s < 0.0 || m < 0.0) {
    return ErrorCode::Num;
  }
  if (s >= m) {
    return ErrorCode::Num;
  }

  const int frequency = static_cast<int>(std::trunc(f_e.value()));
  if (frequency != 1 && frequency != 2 && frequency != 4) {
    return ErrorCode::Num;
  }
  const int basis = static_cast<int>(std::trunc(b_e.value()));
  if (basis < 0 || basis > 4) {
    return ErrorCode::Num;
  }

  CouponDates out{};
  if (!compute_coupon_dates(s, m, frequency, basis, &out)) {
    return ErrorCode::Num;
  }
  return out;
}

// --- COUPPCD(settlement, maturity, frequency, [basis=0]) ---------------
//
// Excel serial of the previous coupon date on or before settlement.
Value CoupPcd(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto ctx = resolve_coupon(args, arity);
  if (!ctx) {
    return Value::error(ctx.error());
  }
  return finalize(ctx.value().pcd);
}

// --- COUPNCD(settlement, maturity, frequency, [basis=0]) ---------------
//
// Excel serial of the next coupon date strictly after settlement.
Value CoupNcd(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto ctx = resolve_coupon(args, arity);
  if (!ctx) {
    return Value::error(ctx.error());
  }
  return finalize(ctx.value().ncd);
}

// --- COUPNUM(settlement, maturity, frequency, [basis=0]) ---------------
//
// Number of coupon dates strictly after settlement and up to (and
// including) maturity. The shared engine already tracks this across
// its backward walk, so the builtin just returns that counter.
Value CoupNum(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto ctx = resolve_coupon(args, arity);
  if (!ctx) {
    return Value::error(ctx.error());
  }
  return finalize(static_cast<double>(ctx.value().coupons_remaining));
}

// --- COUPDAYBS(settlement, maturity, frequency, [basis=0]) -------------
//
// Basis-adjusted days from the start of the current coupon period
// (PCD) to settlement.
Value CoupDayBs(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto ctx = resolve_coupon(args, arity);
  if (!ctx) {
    return Value::error(ctx.error());
  }
  return finalize(ctx.value().days_bs);
}

// --- COUPDAYSNC(settlement, maturity, frequency, [basis=0]) ------------
//
// Basis-adjusted days from settlement to the next coupon date (NCD).
Value CoupDaysNc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto ctx = resolve_coupon(args, arity);
  if (!ctx) {
    return Value::error(ctx.error());
  }
  return finalize(ctx.value().days_nc);
}

// --- COUPDAYS(settlement, maturity, frequency, [basis=0]) --------------
//
// Basis-adjusted total days in the coupon period containing
// settlement. For the 30/360 bases and basis 2/3 this is a fixed
// `360 / freq` or `365 / freq`; for basis 1 (actual/actual) it is the
// raw NCD - PCD gap in days (an integer).
Value CoupDays(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto ctx = resolve_coupon(args, arity);
  if (!ctx) {
    return Value::error(ctx.error());
  }
  return finalize(ctx.value().period_days);
}

}  // namespace

void register_financial_coupon_builtins(FunctionRegistry& registry) {
  // All six: 3 required + 1 optional basis = min 3, max 4. Eager
  // scalar, no range support.
  registry.register_function(FunctionDef{"COUPPCD", 3u, 4u, &CoupPcd});
  registry.register_function(FunctionDef{"COUPNCD", 3u, 4u, &CoupNcd});
  registry.register_function(FunctionDef{"COUPNUM", 3u, 4u, &CoupNum});
  registry.register_function(FunctionDef{"COUPDAYBS", 3u, 4u, &CoupDayBs});
  registry.register_function(FunctionDef{"COUPDAYSNC", 3u, 4u, &CoupDaysNc});
  registry.register_function(FunctionDef{"COUPDAYS", 3u, 4u, &CoupDays});
}

}  // namespace eval
}  // namespace formulon
