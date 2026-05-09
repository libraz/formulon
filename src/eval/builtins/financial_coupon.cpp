// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/coupon_schedule.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using financial_detail::finalize;
using financial_detail::read_coupon_frequency;
using financial_detail::read_day_count_basis;
using financial_detail::read_financial_date;

// Shared argument-validation helper for the COUP* family. Truncates
// settlement / maturity to integer serials, validates the frequency
// / basis domains, enforces `settlement < maturity`, and runs the
// coupon-schedule engine. On success the caller receives a populated
// `CouponDates`; on any validation failure a `#NUM!` error is
// returned for the caller to forward to Excel.
Expected<CouponDates, ErrorCode> resolve_coupon(const Value* args, std::uint32_t arity) {
  auto s_e = read_financial_date(args, 0);
  if (!s_e) {
    return s_e.error();
  }
  auto m_e = read_financial_date(args, 1);
  if (!m_e) {
    return m_e.error();
  }
  auto f_e = read_coupon_frequency(args, 2);
  if (!f_e) {
    return f_e.error();
  }
  auto b_e = read_day_count_basis(args, arity, 3);
  if (!b_e) {
    return b_e.error();
  }

  const double s = s_e.value();
  const double m = m_e.value();
  if (s >= m) {
    return ErrorCode::Num;
  }

  const int frequency = f_e.value();
  const int basis = b_e.value();

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
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"COUPPCD", 3u, 4u, &CoupPcd},     {"COUPNCD", 3u, 4u, &CoupNcd},       {"COUPNUM", 3u, 4u, &CoupNum},
      {"COUPDAYBS", 3u, 4u, &CoupDayBs}, {"COUPDAYSNC", 3u, 4u, &CoupDaysNc}, {"COUPDAYS", 3u, 4u, &CoupDays},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
