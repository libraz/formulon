// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the eager depreciation built-ins: SLN, SYD, DDB, DB,
// VDB, AMORDEGRC, AMORLINC. Registered from `financial.cpp` via
// `register_financial_builtins`.
//
// VDB extends DDB with (a) fractional period ranges and (b) a one-way
// switch from declining-balance to straight-line when the latter would
// deplete the remaining book value faster. AMORDEGRC / AMORLINC are the
// French-accounting degressive and linear methods; they share the
// "first period is YEARFRAC-prorated" pattern but diverge in later
// periods.

#include <algorithm>
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

// Computes YEARFRAC(start, end, basis) following the YEARFRAC builtin's
// rules. Returns `#NUM!` on an unsupported basis or non-finite result.
// Zero-length spans are allowed (unlike DISC/INTRATE which divide by
// the year-fraction).
Expected<double, ErrorCode> yearfrac_for_depreciation(double start, double end, int basis) {
  if (basis < 0 || basis > 4) {
    return ErrorCode::Num;
  }
  const double s = std::trunc(start);
  const double e = std::trunc(end);
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
  if (std::isnan(yf) || std::isinf(yf)) {
    return ErrorCode::Num;
  }
  return yf;
}

// Reads the `basis` argument at `args[index]` if present, otherwise
// returns 0. Truncates toward zero and validates against {0,1,2,3,4}.
Expected<int, ErrorCode> read_basis_dep(const Value* args, std::uint32_t arity, std::uint32_t index) {
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

// Reads a required date argument, truncating toward zero. Negative
// serials are rejected as `#NUM!`.
Expected<double, ErrorCode> read_date_dep(const Value* args, std::uint32_t index) {
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

// The AMORDEGRC depreciation coefficient (French accounting code). The
// "life" used here is 1/rate — a proxy for the estimated useful life of
// the asset. The three buckets match the official French tax rule:
//   * life in [3, 4]    -> 1.5
//   * life in [5, 6]    -> 2.0
//   * life >  6         -> 2.5
// Excel rejects life < 3 or life in (4, 5) via #NUM!; callers surface
// a 0.0 coefficient as the sentinel for that rejection.
double amordegrc_coefficient(double life) noexcept {
  if (life >= 3.0 && life <= 4.0) {
    return 1.5;
  }
  if (life >= 5.0 && life <= 6.0) {
    return 2.0;
  }
  if (life > 6.0) {
    return 2.5;
  }
  return 0.0;
}

}  // namespace

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
//   - period <= 0  or  period > life      ->  #NUM!
//
// `period` and `life` are used as-is (non-integer values are accepted);
// Excel does not floor them before the arithmetic. Excel 365 accepts
// fractional `period` values like 0.1 (the formula is a simple linear
// schedule so fractional periods evaluate naturally).
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
  if (life <= 0.0 || period <= 0.0 || period > life) {
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
//   month_int = INT(month)              (floored to an integer; Excel 365
//                                        Mac passes INT(month) into the
//                                        9-decimal DB formula even though
//                                        the argument boundary accepts a
//                                        fractional value.)
//   rate      = round((1 - (salvage/cost)^(1/life)) * 1000) / 1000
//   dep_1     = cost * rate * (month_int / 12)
//   dep_i     = (cost - total_prior_dep) * rate            (2 <= i <= life)
//   dep_L+1   = (cost - total_prior_dep) * rate * (12 - month_int) / 12
//                                                           (only when
//                                                            month_int < 12)
//
// Domain (checks use month_int, the floored value):
//   - cost   <  0    ->  #NUM!
//   - salvage<  0    ->  #NUM!
//   - life  <= 0     ->  #NUM!
//   - period<  1     ->  #NUM!
//   - month_int <  1  or month_int > 12  ->  #NUM!
//   - period == life+1 is only valid when month_int < 12 (the partial
//     first year leaves a partial last year); otherwise #NUM!.
//   - period >  life+1 -> #NUM! regardless of month.
//
// Edge cases:
//   - `cost == 0`: short-circuits to 0 depreciation. Computing the rate
//     as `pow(salvage/0, 1/life)` would go through `pow(inf, 1/life) =
//     inf` and poison the result with #NUM!, so we return a zero charge
//     directly — matching Mac Excel 365's observed behaviour for
//     zero-cost assets (including boolean FALSE / numeric 0).
//   - `salvage == 0`: `pow(0, 1/life) = 0`, so rate = round(1) = 1.000
//     and the entire cost depreciates in the first period (prorated by
//     `month_int / 12`). Subsequent periods charge nothing.
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
  // Excel 365 Mac floors `month` before every downstream use (domain
  // checks, first-period proration, and the partial-last-year factor).
  const double month_int = std::floor(month_e.value());
  if (cost < 0.0 || salvage < 0.0 || life <= 0.0 || period < 1.0 || month_int < 1.0 || month_int > 12.0) {
    return Value::error(ErrorCode::Num);
  }
  // Zero-cost asset: salvage/cost would be +inf and pow(inf, 1/life) also
  // +inf, poisoning the subsequent arithmetic with #NUM!. Mac Excel 365
  // returns a zero depreciation here (verified via G30/G32/G33/G34 in
  // the IronCalc DB oracle), so short-circuit before the rate formula.
  if (cost == 0.0) {
    return Value::number(0.0);
  }
  // Only valid period values: [1, life] for any month; life+1 only when
  // month_int < 12 (partial-first-year case); anything beyond is #NUM!.
  if (period > life + 1.0) {
    return Value::error(ErrorCode::Num);
  }
  if (period > life && month_int >= 12.0) {
    return Value::error(ErrorCode::Num);
  }
  // Rate: (1 - (salvage/cost)^(1/life)), rounded to 3 decimals. When
  // salvage == 0, pow(0, 1/life) == 0 -> rate = 1.000 exactly.
  double rate_raw = 1.0 - std::pow(salvage / cost, 1.0 / life);
  const double rate = std::floor(rate_raw * 1000.0 + 0.5) / 1000.0;
  // First period: prorated by (month_int/12).
  const double dep_1 = cost * rate * month_int / 12.0;
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
  // period == life + 1 (partial last year). Guarded above so month_int < 12.
  const double dep_last = (cost - total) * rate * (12.0 - month_int) / 12.0;
  return finalize(dep_last);
}

// --- VDB(cost, salvage, life, start_period, end_period, [factor=2],
//         [no_switch=FALSE]) ------------------------------------------------
//
// Variable-declining-balance depreciation summed over the fractional
// period range [start_period, end_period]. Two variants depending on
// `no_switch`:
//
//   * no_switch == FALSE (default): declining-balance until the
//     straight-line depreciation of the remaining book value over the
//     remaining life exceeds the declining-balance charge, then switch
//     (one-way) to straight-line for every subsequent period.
//   * no_switch == TRUE: pure declining-balance for every period,
//     capped by `book - salvage` in each period.
//
// Fractional endpoints are handled by linear prorating of the affected
// period's full-period depreciation (this matches Excel's observed
// behaviour, per the `VDB(2400, 300, 10, 0, 0.875, 1.5, FALSE)` = 315.0
// reference).
//
// Domain:
//   - cost     <  0                        ->  #NUM!
//   - salvage  <  0                        ->  #NUM!
//   - life    <=  0                        ->  #NUM!
//   - start_period < 0                     ->  #NUM!
//   - start_period >= end_period           ->  #NUM!
//   - end_period  > life                   ->  #NUM!
//   - factor  <=  0                        ->  #NUM!
Value Vdb(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
  auto start_e = read_required_number(args, 3);
  if (!start_e) {
    return Value::error(start_e.error());
  }
  auto end_e = read_required_number(args, 4);
  if (!end_e) {
    return Value::error(end_e.error());
  }
  auto factor_e = read_optional_number(args, arity, 5, 2.0);
  if (!factor_e) {
    return Value::error(factor_e.error());
  }
  auto no_switch_e = read_optional_number(args, arity, 6, 0.0);
  if (!no_switch_e) {
    return Value::error(no_switch_e.error());
  }
  const double cost = cost_e.value();
  const double salvage = salvage_e.value();
  const double life = life_e.value();
  const double start_period = start_e.value();
  const double end_period = end_e.value();
  const double factor = factor_e.value();
  const bool no_switch = no_switch_e.value() != 0.0;
  if (cost < 0.0 || salvage < 0.0 || life <= 0.0 || factor <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (start_period < 0.0 || start_period >= end_period || end_period > life) {
    return Value::error(ErrorCode::Num);
  }

  // Walk integer periods 1..ceil(end_period), maintaining the running
  // book value. Each integer period `t` covers the half-open segment
  // (t-1, t]; we clip that segment to [start_period, end_period] and
  // linearly prorate the full-period depreciation by the clipped
  // length.
  const auto last_iter = static_cast<std::int64_t>(std::ceil(end_period));
  double book = cost;
  double accumulated = 0.0;
  bool switched = false;  // Sticky once we flip to straight-line.
  double straight_line_dep = 0.0;

  for (std::int64_t t = 1; t <= last_iter; ++t) {
    // Full-period depreciation under the active method. Under
    // no_switch=FALSE we compare DDB against the straight-line charge
    // spread over the remaining life; once SL wins we stay on SL.
    double dep = 0.0;
    if (no_switch) {
      dep = book * factor / life;
    } else {
      if (switched) {
        dep = straight_line_dep;
      } else {
        const double ddb = book * factor / life;
        const double remaining_life = life - static_cast<double>(t - 1);
        const double sl_candidate = remaining_life > 0.0 ? (book - salvage) / remaining_life : 0.0;
        if (sl_candidate > ddb) {
          switched = true;
          straight_line_dep = sl_candidate;
          dep = sl_candidate;
        } else {
          dep = ddb;
        }
      }
    }
    // Salvage floor: never depreciate below salvage.
    const double cap = book - salvage;
    if (dep > cap) {
      dep = cap;
    }
    if (dep < 0.0) {
      dep = 0.0;
    }

    // Clip integer-period segment (t-1, t] to [start_period, end_period].
    const double seg_lo = std::max(start_period, static_cast<double>(t - 1));
    const double seg_hi = std::min(end_period, static_cast<double>(t));
    if (seg_hi > seg_lo) {
      accumulated += dep * (seg_hi - seg_lo);
    }
    book -= dep;
    if (book <= salvage) {
      // No further depreciation possible; remaining periods contribute 0.
      break;
    }
  }
  return finalize(accumulated);
}

// --- AMORDEGRC(cost, date_purchased, first_period, salvage, period,
//               rate, [basis=0]) --------------------------------------------
//
// French degressive (declining-balance) depreciation. The algorithm
// applies a life-dependent coefficient to `rate`:
//
//   life          = 1 / rate
//   coefficient   = 1.5 if 3 <= life <= 4
//                 = 2.0 if 5 <= life <= 6
//                 = 2.5 if life > 6
//   applied_rate  = rate * coefficient
//
// Depreciation schedule (period index is 0-based):
//
//   dep_0          = round(cost * applied_rate *
//                          YEARFRAC(date_purchased, first_period, basis))
//   dep_i (i>=1)   = round(book_i * applied_rate)
//
// The second-to-last full period's charge is forced to `book / 2` and
// the final period zeros out the remainder — matching Excel's
// observable behaviour for long asset lives.
//
// Domain:
//   - cost <= 0 or salvage >= cost or rate <= 0  ->  #NUM!
//   - period < 0                                 ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}               ->  #NUM!
//   - life (= 1/rate) in the rejected buckets    ->  #NUM!
Value Amordegrc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto cost_e = read_required_number(args, 0);
  if (!cost_e) {
    return Value::error(cost_e.error());
  }
  auto date_purchased_e = read_date_dep(args, 1);
  if (!date_purchased_e) {
    return Value::error(date_purchased_e.error());
  }
  auto first_period_e = read_date_dep(args, 2);
  if (!first_period_e) {
    return Value::error(first_period_e.error());
  }
  auto salvage_e = read_required_number(args, 3);
  if (!salvage_e) {
    return Value::error(salvage_e.error());
  }
  auto period_e = read_required_number(args, 4);
  if (!period_e) {
    return Value::error(period_e.error());
  }
  auto rate_e = read_required_number(args, 5);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  auto basis_e = read_basis_dep(args, arity, 6);
  if (!basis_e) {
    return Value::error(basis_e.error());
  }
  const double cost = cost_e.value();
  const double salvage = salvage_e.value();
  const double period = period_e.value();
  const double rate = rate_e.value();
  if (cost <= 0.0 || rate <= 0.0 || salvage >= cost || period < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  // Computed life drives the French coefficient table. A coefficient of
  // 0 means the life bucket is invalid (Excel's #NUM! territory).
  const double life = 1.0 / rate;
  const double coef = amordegrc_coefficient(life);
  if (coef == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double applied_rate = rate * coef;

  // First-period year fraction (YEARFRAC from purchase to first period).
  auto yf = yearfrac_for_depreciation(date_purchased_e.value(), first_period_e.value(), basis_e.value());
  if (!yf) {
    return Value::error(yf.error());
  }

  // Walk periods 0..period, rounding each period's depreciation to the
  // nearest integer (Excel's observed AMORDEGRC behaviour). The schedule
  // terminates early once book value reaches salvage.
  const auto requested = static_cast<std::int64_t>(std::floor(period));
  double book = cost;
  double dep_i = 0.0;
  // Integer estimate of the asset's life (in periods) used by the
  // two-step end-of-schedule rule. Excel rounds life to the nearest
  // integer for this calculation.
  const auto life_int = static_cast<std::int64_t>(std::floor(life + 0.5));
  for (std::int64_t i = 0; i <= requested; ++i) {
    // End-of-schedule rule: the penultimate full period depreciates
    // half of the remaining book value; the last period finishes the
    // remainder. life_int here is a conservative integer proxy.
    if (i == life_int - 1) {
      dep_i = std::floor(book * 0.5 + 0.5);
    } else if (i >= life_int) {
      dep_i = book;
    } else if (i == 0) {
      dep_i = std::floor(cost * applied_rate * yf.value() + 0.5);
    } else {
      dep_i = std::floor(book * applied_rate + 0.5);
    }
    // Cap against the remaining depreciable book value (book - salvage).
    const double cap = book - salvage;
    if (dep_i > cap) {
      dep_i = cap;
    }
    if (dep_i < 0.0) {
      dep_i = 0.0;
    }
    book -= dep_i;
    if (book <= salvage && i < requested) {
      // Asset is fully depreciated; every remaining period returns 0.
      dep_i = 0.0;
      break;
    }
  }
  return finalize(dep_i);
}

// --- AMORLINC(cost, date_purchased, first_period, salvage, period, rate,
//              [basis=0]) ---------------------------------------------------
//
// French straight-line (linear) depreciation. The first period is
// prorated by the year-fraction from the purchase date to the end of
// the first accounting period; every subsequent period charges the
// same flat amount (`cost * rate`) until the asset reaches salvage.
//
//   dep_0          = cost * rate * YEARFRAC(date_purchased, first_period, basis)
//   dep_i (i>=1)   = cost * rate
//
// The last period is whatever remains to bring the book value down to
// salvage. Unlike AMORDEGRC this implementation does not round each
// period to an integer — AMORLINC returns the raw straight-line
// amount, matching Excel.
//
// Domain: same as AMORDEGRC minus the life-bucket rejection (AMORLINC
// is valid for every positive rate).
Value Amorlinc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto cost_e = read_required_number(args, 0);
  if (!cost_e) {
    return Value::error(cost_e.error());
  }
  auto date_purchased_e = read_date_dep(args, 1);
  if (!date_purchased_e) {
    return Value::error(date_purchased_e.error());
  }
  auto first_period_e = read_date_dep(args, 2);
  if (!first_period_e) {
    return Value::error(first_period_e.error());
  }
  auto salvage_e = read_required_number(args, 3);
  if (!salvage_e) {
    return Value::error(salvage_e.error());
  }
  auto period_e = read_required_number(args, 4);
  if (!period_e) {
    return Value::error(period_e.error());
  }
  auto rate_e = read_required_number(args, 5);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  auto basis_e = read_basis_dep(args, arity, 6);
  if (!basis_e) {
    return Value::error(basis_e.error());
  }
  const double cost = cost_e.value();
  const double salvage = salvage_e.value();
  const double period = period_e.value();
  const double rate = rate_e.value();
  if (cost <= 0.0 || rate <= 0.0 || salvage >= cost || period < 0.0) {
    return Value::error(ErrorCode::Num);
  }

  auto yf = yearfrac_for_depreciation(date_purchased_e.value(), first_period_e.value(), basis_e.value());
  if (!yf) {
    return Value::error(yf.error());
  }

  const double dep_first = cost * rate * yf.value();
  const double dep_flat = cost * rate;

  const auto requested = static_cast<std::int64_t>(std::floor(period));
  double book = cost;
  double dep_i = 0.0;
  for (std::int64_t i = 0; i <= requested; ++i) {
    if (i == 0) {
      dep_i = dep_first;
    } else {
      dep_i = dep_flat;
    }
    // Never push book value below salvage.
    const double cap = book - salvage;
    if (dep_i > cap) {
      dep_i = cap;
    }
    if (dep_i < 0.0) {
      dep_i = 0.0;
    }
    book -= dep_i;
    if (book <= salvage && i < requested) {
      dep_i = 0.0;
      break;
    }
  }
  return finalize(dep_i);
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
