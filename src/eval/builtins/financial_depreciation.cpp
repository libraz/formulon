// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the eager depreciation built-ins: SLN, SYD, DDB, DB.
// Registered from `financial.cpp` via `register_financial_builtins`.

#include <algorithm>
#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

// --- SLN(cost, salvage, life) -------------------------------------------
//
// Straight-line depreciation: equal depreciation charge across every
// period of the asset's life.
//
//   SLN = (cost - salvage) / life
//
// Domain:
//   - life == 0  ->  #DIV/0!
//
// Oracle-verified: Excel 365 Mac imposes no further sign constraints —
// negative `life` simply yields a negative depreciation (the raw
// algebraic answer), and any signed values for `cost`/`salvage` are
// accepted (including `salvage > cost`, which produces a negative
// charge).
Value Sln(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto cost_e = read_required_number(args, 0);
  if (!cost_e) {
    return Value::error(cost_e.error());
  }
  auto salvage_e = read_required_number(args, 1);
  if (!salvage_e) {
    return Value::error(salvage_e.error());
  }
  auto life_e = read_required_number(args, 2);
  if (!life_e) {
    return Value::error(life_e.error());
  }
  const double cost = cost_e.value();
  const double salvage = salvage_e.value();
  const double life = life_e.value();
  if (life == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  return finalize((cost - salvage) / life);
}

// --- SYD(cost, salvage, life, period) -----------------------------------
//
// Sum-of-years'-digits depreciation:
//
//   SYD = (cost - salvage) * (life - period + 1) * 2 / (life * (life + 1))
//
// Domain:
//   - life   <= 0                         ->  #NUM!
//   - period <  1  or  period > life      ->  #NUM!
//
// `period` and `life` are used as-is (non-integer values are accepted);
// Excel does not floor them before the arithmetic.
Value Syd(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto cost_e = read_required_number(args, 0);
  if (!cost_e) {
    return Value::error(cost_e.error());
  }
  auto salvage_e = read_required_number(args, 1);
  if (!salvage_e) {
    return Value::error(salvage_e.error());
  }
  auto life_e = read_required_number(args, 2);
  if (!life_e) {
    return Value::error(life_e.error());
  }
  auto period_e = read_required_number(args, 3);
  if (!period_e) {
    return Value::error(period_e.error());
  }
  const double cost = cost_e.value();
  const double salvage = salvage_e.value();
  const double life = life_e.value();
  const double period = period_e.value();
  if (life <= 0.0 || period < 1.0 || period > life) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (cost - salvage) * (life - period + 1.0) * 2.0 / (life * (life + 1.0));
  return finalize(result);
}

// --- DDB(cost, salvage, life, period, [factor=2]) -----------------------
//
// Declining-balance depreciation generalised to an arbitrary accelerator
// `factor` (default 2 = double-declining balance).
//
//   rate = factor / life       (clamped to 1 if rate >= 1 — everything
//                               depreciates in period 1)
//
// Iterate from 1..period, maintaining a running book value. The salvage
// floor prevents the charge in any period from pushing the book value
// below `salvage`:
//
//   dep_i = max(0, min(book_value * rate, book_value - salvage))
//   book_value -= dep_i
//
// Only the charge for the requested `period` is returned.
//
// Domain:
//   - cost    <  0  ->  #NUM!
//   - salvage <  0  ->  #NUM!
//   - life   <= 0   ->  #NUM!
//   - period <  1  or period > life  ->  #NUM!
//   - factor <= 0   ->  #NUM!
//
// Period is expected to be a positive integer; Excel's documented
// behaviour for non-integer period is the integer-iteration schedule
// (which is what we implement). If oracle testing ever surfaces a
// fractional-period case that diverges, revisit here.
Value Ddb(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto cost_e = read_required_number(args, 0);
  if (!cost_e) {
    return Value::error(cost_e.error());
  }
  auto salvage_e = read_required_number(args, 1);
  if (!salvage_e) {
    return Value::error(salvage_e.error());
  }
  auto life_e = read_required_number(args, 2);
  if (!life_e) {
    return Value::error(life_e.error());
  }
  auto period_e = read_required_number(args, 3);
  if (!period_e) {
    return Value::error(period_e.error());
  }
  auto factor_e = read_optional_number(args, arity, 4, 2.0);
  if (!factor_e) {
    return Value::error(factor_e.error());
  }
  const double cost = cost_e.value();
  const double salvage = salvage_e.value();
  const double life = life_e.value();
  const double period = period_e.value();
  const double factor = factor_e.value();
  if (cost < 0.0 || salvage < 0.0 || life <= 0.0 || period < 1.0 || period > life || factor <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  double rate = factor / life;
  if (rate >= 1.0) {
    rate = 1.0;
  }
  // Iterate integer periods up to (and including) the requested period.
  // Using ceil() on the requested period lets a fractional period map
  // to the same integer-iteration answer Excel produces in practice.
  const auto iter_count = static_cast<std::int64_t>(std::ceil(period));
  double book_value = cost;
  double depreciation = 0.0;
  for (std::int64_t i = 1; i <= iter_count; ++i) {
    double dep_i = book_value * rate;
    const double floor_cap = book_value - salvage;
    if (dep_i > floor_cap) {
      dep_i = floor_cap;
    }
    if (dep_i < 0.0) {
      dep_i = 0.0;
    }
    depreciation = dep_i;
    book_value -= dep_i;
  }
  return finalize(depreciation);
}

// --- DB(cost, salvage, life, period, [month=12]) ------------------------
//
// Fixed-rate declining-balance depreciation with Excel's quirky
// 3-decimal rate rounding and partial-first-year proration:
//
//   rate   = round((1 - (salvage/cost)^(1/life)) * 1000) / 1000
//   dep_1  = cost * rate * (month / 12)
//   dep_i  = (cost - total_prior_dep) * rate               (2 <= i <= life)
//   dep_L+1= (cost - total_prior_dep) * rate * (12 - month) / 12
//                                                           (only when
//                                                            month < 12)
//
// Domain:
//   - cost   <= 0    ->  #NUM!
//   - salvage<  0    ->  #NUM!
//   - life  <= 0     ->  #NUM!
//   - period<  1     ->  #NUM!
//   - month <  1  or month > 12  ->  #NUM!
//   - period == life+1 is only valid when month < 12 (the partial first
//     year leaves a partial last year); otherwise #NUM!.
//   - period >  life+1 -> #NUM! regardless of month.
//
// Edge cases:
//   - `salvage == 0`: `pow(0, 1/life) = 0`, so rate = round(1) = 1.000
//     and the entire cost depreciates in the first period (prorated by
//     `month / 12`). Subsequent periods charge nothing.
//   - `salvage > cost`: the inner base is > 1, so the raw rate is
//     negative; rounding then yields a negative rate and Excel returns
//     a negative depreciation. No special-cased error path here — we
//     let the math speak, matching observed Excel behaviour.
Value Db(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto cost_e = read_required_number(args, 0);
  if (!cost_e) {
    return Value::error(cost_e.error());
  }
  auto salvage_e = read_required_number(args, 1);
  if (!salvage_e) {
    return Value::error(salvage_e.error());
  }
  auto life_e = read_required_number(args, 2);
  if (!life_e) {
    return Value::error(life_e.error());
  }
  auto period_e = read_required_number(args, 3);
  if (!period_e) {
    return Value::error(period_e.error());
  }
  auto month_e = read_optional_number(args, arity, 4, 12.0);
  if (!month_e) {
    return Value::error(month_e.error());
  }
  const double cost = cost_e.value();
  const double salvage = salvage_e.value();
  const double life = life_e.value();
  const double period = period_e.value();
  const double month = month_e.value();
  if (cost <= 0.0 || salvage < 0.0 || life <= 0.0 || period < 1.0 || month < 1.0 || month > 12.0) {
    return Value::error(ErrorCode::Num);
  }
  // Only valid period values: [1, life] for any month; life+1 only when
  // month < 12 (partial-first-year case); anything beyond is #NUM!.
  if (period > life + 1.0) {
    return Value::error(ErrorCode::Num);
  }
  if (period > life && month >= 12.0) {
    return Value::error(ErrorCode::Num);
  }
  // Rate: (1 - (salvage/cost)^(1/life)), rounded to 3 decimals. When
  // salvage == 0, pow(0, 1/life) == 0 -> rate = 1.000 exactly.
  double rate_raw = 1.0 - std::pow(salvage / cost, 1.0 / life);
  const double rate = std::floor(rate_raw * 1000.0 + 0.5) / 1000.0;
  // First period: prorated by (month/12).
  const double dep_1 = cost * rate * month / 12.0;
  if (period < 2.0) {
    return finalize(dep_1);
  }
  double total = dep_1;
  double dep_i = dep_1;
  // Middle periods: i in [2, min(period, life)].
  const auto last_full = static_cast<std::int64_t>(std::floor(std::min(period, life)));
  for (std::int64_t i = 2; i <= last_full; ++i) {
    dep_i = (cost - total) * rate;
    total += dep_i;
  }
  if (period <= life) {
    return finalize(dep_i);
  }
  // period == life + 1 (partial last year). Guarded above so month < 12.
  const double dep_last = (cost - total) * rate * (12.0 - month) / 12.0;
  return finalize(dep_last);
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
