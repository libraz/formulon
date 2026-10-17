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
#include "eval/builtins/numeric_helpers.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

struct DepreciationArgs4 {
  double cost;
  double salvage;
  double life;
  double period;
};

struct DepreciationArgs3 {
  double cost;
  double salvage;
  double life;
};

struct VdbArgs {
  double cost;
  double salvage;
  double life;
  double start_period;
  double end_period;
  double factor;
  bool no_switch;
};

struct AmorArgs {
  double cost;
  double date_purchased;
  double first_period;
  double salvage;
  double period;
  double rate;
  int basis;
};

Expected<DepreciationArgs3, ErrorCode> read_depreciation_args3(const Value* args) {
  double v[3];
  if (auto read = read_required_numbers(args, "nnn", v); !read) {
    return std::move(read.error());
  }
  return DepreciationArgs3{v[0], v[1], v[2]};
}

Expected<DepreciationArgs4, ErrorCode> read_depreciation_args4(const Value* args) {
  double v[4];
  if (auto read = read_required_numbers(args, "nnnn", v); !read) {
    return std::move(read.error());
  }
  return DepreciationArgs4{v[0], v[1], v[2], v[3]};
}

Expected<VdbArgs, ErrorCode> read_vdb_args(const Value* args, std::uint32_t arity) {
  double v[5];
  if (auto read = read_required_numbers(args, "nnnnn", v); !read) {
    return std::move(read.error());
  }
  auto factor = read_optional_number(args, arity, 5, 2.0);
  if (!factor) {
    return std::move(factor.error());
  }
  auto no_switch = read_optional_number(args, arity, 6, 0.0);
  if (!no_switch) {
    return std::move(no_switch.error());
  }
  return VdbArgs{v[0], v[1], v[2], v[3], v[4], factor.value(), no_switch.value() != 0.0};
}

// Reads the AMORDEGRC / AMORLINC arguments and applies their shared domain checks.
// `salvage_may_equal_cost` admits a salvage equal to the cost (AMORDEGRC returns 0 there).
Expected<AmorArgs, ErrorCode> read_amor_args(const Value* args, std::uint32_t arity, bool salvage_may_equal_cost) {
  double v[6];
  if (auto read = read_required_numbers(args, "nddnnn", v); !read) {
    return std::move(read.error());
  }
  auto basis = read_day_count_basis(args, arity, 6);
  if (!basis) {
    return std::move(basis.error());
  }
  const double cost = v[0];
  const double salvage = v[3];
  const double period = v[4];
  const double rate = v[5];
  if (cost <= 0.0 || rate <= 0.0 || (salvage_may_equal_cost ? salvage > cost : salvage >= cost) || period < 0.0 ||
      period > kMaxDepreciationPeriods) {
    return ErrorCode::Num;
  }
  return AmorArgs{cost, v[1], v[2], salvage, period, rate, basis.value()};
}

// Declining-balance charge of a (possibly fractional) `period`:
//
//   r = factor / life, q = 1 - r
//   t = cost * q^(period - 1)  (t = cost for period <= 1; 0 when q < 0 and period > 1)
//   charge = t * r, capped at t - salvage and floored at 0
double ddb_charge(double cost, double salvage, double life, double period, double factor) noexcept {
  const double r = factor / life;
  const double q = 1.0 - r;
  if (q < 0.0 && period > 1.0) {
    return 0.0;
  }
  const double power = period <= 1.0 ? 1.0 : excel_power(q, period - 1.0);
  const double book = cost * power;
  const double charge = book * r;
  const double remaining = book - salvage;
  if (remaining - charge >= 0.0) {
    return charge;
  }
  return remaining >= 0.0 ? remaining : 0.0;
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
  auto parsed = read_depreciation_args3(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [cost, salvage, life] = parsed.value();
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
  auto parsed = read_depreciation_args4(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [cost, salvage, life, period] = parsed.value();
  if (life <= 0.0 || period <= 0.0 || period > life) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (cost - salvage) * (life - period + 1.0) * 2.0 / (life * (life + 1.0));
  return finalize(result);
}

// --- DDB(cost, salvage, life, period, [factor=2]) -----------------------
//
// Declining-balance depreciation with accelerator `factor` (default 2).
// `period` may be fractional. See `ddb_charge` for the formula.
//
// Domain: cost < 0, salvage < 0, life <= 0, period <= 0, period > life or
// factor <= 0 is #NUM!.
Value Ddb(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto parsed = read_depreciation_args4(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  auto factor_e = read_optional_number(args, arity, 4, 2.0);
  if (!factor_e) {
    return Value::error(factor_e.error());
  }
  const auto [cost, salvage, life, period] = parsed.value();
  const double factor = factor_e.value();
  if (cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || period > life || factor <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(ddb_charge(cost, salvage, life, period, factor));
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
//   rate      = ROUND(1 - POWER(salvage/cost, 1/life), 3)
//   dep_1     = (cost * rate) * (month_int / 12)
//   dep_i     = book * rate, book -= dep_i                 (2 <= i <= life)
//   dep_L+1   = (12 - month_int) * (rate * (book / 12))     (partial last year)
//
// Life and period count whole years (truncated; 2^32 or more is #NUM!). A
// period past the life is valid only one year past it, and only when
// month_int < 12.
//
// Domain (checks use month_int, the floored value):
//   - cost < 0, salvage < 0, life <= 0, period <= 0  ->  #NUM!
//   - month_int < 1 or month_int > 12                ->  #NUM!
//   - cost == 0 gives a zero charge.
Value Db(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto parsed = read_depreciation_args4(args);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  auto month_e = read_optional_number(args, arity, 4, 12.0);
  if (!month_e) {
    return Value::error(month_e.error());
  }
  const auto [cost, salvage, life, period] = parsed.value();
  // Excel 365 Mac floors `month` before every downstream use (domain
  // checks, first-period proration, and the partial-last-year factor).
  const double month_int = std::floor(month_e.value());
  if (cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || month_int < 1.0 || month_int > 12.0) {
    return Value::error(ErrorCode::Num);
  }
  // Life and period count whole years; either at 2^32 or beyond is #NUM!.
  constexpr double kMaxWholeYears = 4294967296.0;
  const double whole_life = std::trunc(life);
  const double whole_period = std::trunc(period);
  if (whole_life >= kMaxWholeYears || whole_period >= kMaxWholeYears) {
    return Value::error(ErrorCode::Num);
  }
  // A period past the life exists only as the partial last year of a first year
  // shorter than 12 months, and only one period past.
  double last = whole_period;
  const bool partial_last = whole_period > whole_life;
  if (partial_last) {
    if (month_int == 12.0 || whole_period > whole_life + 1.0) {
      return Value::error(ErrorCode::Num);
    }
    last = whole_period - 1.0;
  }
  if (cost == 0.0) {
    return Value::number(0.0);
  }
  // Rate: 1 - (salvage/cost)^(1/life) through POWER, rounded to 3 decimals.
  const double rate = builtins_detail::round_half_away_decimals(1.0 - excel_power(salvage / cost, 1.0 / life), 3);
  double dep = (cost * rate) * (month_int / 12.0);
  double book = cost - dep;
  for (double i = 2.0; i <= last; i += 1.0) {
    dep = book * rate;
    book = book - dep;
    if (dep == 0.0) {
      break;  // the book value no longer changes
    }
  }
  if (partial_last) {
    return finalize((12.0 - month_int) * (rate * (book / 12.0)));
  }
  return finalize(dep);
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
  auto parsed = read_vdb_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const VdbArgs& in = parsed.value();
  const double cost = in.cost;
  const double salvage = in.salvage;
  const double life_arg = in.life;
  const double start_period = in.start_period;
  const double end_period = in.end_period;
  const double factor = in.factor;
  const bool no_switch = in.no_switch;
  if (cost < 0.0 || salvage < 0.0 || life_arg <= 0.0 || factor < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (start_period < 0.0 || start_period > end_period || end_period > life_arg) {
    return Value::error(ErrorCode::Num);
  }
  // The schedule walks one iteration per integer period, and the
  // switch-to-straight-line state machine has no closed form to fall back
  // on, so an over-long schedule is refused outright.
  if (end_period > kMaxDepreciationPeriods) {
    return Value::error(ErrorCode::Num);
  }

  const auto ddb = [&](double n) { return ddb_charge(cost, salvage, life_arg, n, factor); };

  if (no_switch) {
    double n = std::trunc(std::fabs(start_period)) + 1.0;
    double charge = ddb(n);
    double total = ((n - 1.0) - start_period) * charge;
    while (end_period - n > 0.0) {
      total = charge + total;
      n = n + 1.0;
      charge = ddb(n);
    }
    return finalize(((end_period - n) + 1.0) * charge + total);
  }

  double life = life_arg;
  const double rate = factor / life;
  double total = 0.0;
  double book = cost;
  double start = start_period;
  const double whole = std::trunc(std::fabs(start));
  if (end_period - whole != 0.0) {
    const double first = ddb(1.0);
    const double straight = (cost - salvage) / life;
    if (straight - first > 0.0) {
      return finalize((end_period - start) * straight + total);
    }
    const double fraction = start - whole;
    book = -(fraction * first) + cost;
    life = life - fraction;
  }
  // One step of the declining-balance walk: the charge for the current
  // period, or the straight-line charge `slope` once that beats it.
  const auto walk = [&](double* charge, double* slope) {
    *charge = book * rate;
    const double remaining = book - salvage;
    if (!(remaining - *charge > 0.0)) {
      *charge = remaining;
      return false;
    }
    *slope = remaining / life;
    return *slope - *charge > 0.0;
  };
  double charge = 0.0;
  double slope = 0.0;
  for (auto i = static_cast<std::int64_t>(whole); i > 0; --i) {
    if (walk(&charge, &slope)) {
      return finalize((end_period - start) * slope + total);
    }
    book = book - charge;
    life = life - 1.0;
  }
  while (true) {
    if (walk(&charge, &slope)) {
      return finalize((end_period - start) * slope + total);
    }
    if (!((end_period - start) - 1.0 > 0.0)) {
      return finalize((end_period - start) * charge + total);
    }
    total = total + charge;
    book = book - charge;
    start = start + 1.0;
    life = life - 1.0;
  }
}

// --- AMORDEGRC(cost, date_purchased, first_period, salvage, period,
//               rate, [basis=0]) --------------------------------------------
//
// French degressive depreciation. With life L = ceil(1/rate) (L < 3 is #NUM!),
// the rate is multiplied by 1.5 (L of 3-4), 2 (5-6) or 2.5 (7 and up).
//
//   first charge = min(((days / year) * r) * cost, cost - salvage), rounded;
//                  a zero-length first period charges a full year and leaves L as is,
//                  a nonzero one adds 1 to L
//   later charges: unrounded `r * book`, the whole of the next charge being
//                  half the book when L - n reaches 2, floored at the salvage
//
// The returned period's charge is rounded. A period beyond L, or a salvage equal
// to the cost, is 0. Basis 2 and a salvage above the cost are #NUM!.
//
// The first-period day count is 30/360 (US with the February-end rule for
// basis 0, European for basis 4), actual/365 (basis 3), or actual days over
// 366 / 365 by whether the purchase year is a leap year (basis 1).
double amordegrc_first_period_fraction(double purchased, double first_period, int basis, bool date1904) noexcept {
  const date_time::YMD a = date_time::ymd_from_serial(purchased, date1904);
  const date_time::YMD b = date_time::ymd_from_serial(first_period, date1904);
  if (basis == 3) {
    return (first_period - purchased) / 365.0;
  }
  if (basis == 1) {
    return (first_period - purchased) / (date_time::days_in_month(a.y, 2u) == 29u ? 366.0 : 365.0);
  }
  unsigned d1 = a.d;
  unsigned d2 = b.d;
  if (basis == 0) {
    if (is_feb_end(a) && is_feb_end(b)) {
      d2 = 30u;
    }
    if (is_feb_end(a)) {
      d1 = 30u;
    }
    if (d1 == 31u) {
      d1 = 30u;
    }
    if (d2 == 31u && d1 >= 30u) {
      d2 = 30u;
    }
  } else {
    d1 = d1 == 31u ? 30u : d1;
    d2 = d2 == 31u ? 30u : d2;
  }
  const double days = static_cast<double>(b.y - a.y) * 360.0 +
                      (static_cast<double>(b.m) - static_cast<double>(a.m)) * 30.0 +
                      (static_cast<double>(d2) - static_cast<double>(d1));
  return days / 360.0;
}

Value Amordegrc(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_amor_args(args, arity, /*salvage_may_equal_cost=*/true);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const double cost = parsed.value().cost;
  const double salvage = parsed.value().salvage;
  const double rate = parsed.value().rate;
  const int basis = parsed.value().basis;
  const double purchased = std::trunc(parsed.value().date_purchased);
  const double first_period = std::trunc(parsed.value().first_period);
  if (basis == 2 || first_period < purchased) {
    return Value::error(ErrorCode::Num);
  }
  double life = std::ceil(1.0 / rate);
  if (life < 3.0) {
    return Value::error(ErrorCode::Num);
  }
  if (cost == salvage) {
    return Value::number(0.0);
  }
  const double period = std::floor(parsed.value().period);
  if (period > life) {
    return Value::number(0.0);
  }
  const double coefficient = life <= 4.0 ? 1.5 : (life <= 6.0 ? 2.0 : 2.5);
  double r = rate * coefficient;
  double first = (amordegrc_first_period_fraction(purchased, first_period, basis, date1904) * r) * cost;
  if (first != 0.0) {
    life += 1.0;
  } else {
    first = cost * r;
  }
  first = builtins_detail::round_half_away_decimals(std::min(first, cost - salvage), 0);
  if (period == 0.0) {
    return finalize(first);
  }
  double book = cost - first;
  double dep = 0.0;
  for (double n = 1.0; n <= period;) {
    dep = r * book;
    n += 1.0;
    if (life - n == 2.0) {
      dep = book * 0.5;
      r = 1.0;
    }
    if (book < salvage) {
      dep = std::max(book - salvage, 0.0);
    }
    book -= dep;
  }
  return finalize(builtins_detail::round_half_away_decimals(dep, 0));
}

// --- AMORLINC(cost, date_purchased, first_period, salvage, period, rate,
//              [basis=0]) ---------------------------------------------------
//
// French straight-line (linear) depreciation. The first period is
// prorated by the year-fraction from the purchase date to the end of
// the first accounting period; every subsequent period charges the
// same flat amount (`cost * rate`) until the asset reaches salvage.
//
//   dep_0          = ((days / B) * rate) * cost       (B = year length of the basis)
//   dep_i (i>=1)   = cost * rate
//
// The last period is whatever remains to bring the book value down to
// salvage. Unlike AMORDEGRC this implementation does not round each
// period to an integer — AMORLINC returns the raw straight-line
// amount, matching Excel.
//
// Domain: same as AMORDEGRC minus the life-bucket rejection (AMORLINC
// is valid for every positive rate); basis 2 and first_period < date_purchased are #NUM!.
Value Amorlinc(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_amor_args(args, arity, /*salvage_may_equal_cost=*/false);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const double cost = parsed.value().cost;
  const double salvage = parsed.value().salvage;
  const double rate = parsed.value().rate;
  const int basis = parsed.value().basis;
  const double purchased = std::trunc(parsed.value().date_purchased);
  const double first_period = std::trunc(parsed.value().first_period);
  if (basis == 2 || first_period < purchased) {
    return Value::error(ErrorCode::Num);
  }

  const date_time::YMD dp = date_time::ymd_from_serial(purchased, date1904);
  // Actual-day bases clip a February 29 to the 28th on both ends.
  const auto clipped = [&](double serial) {
    const date_time::YMD d = date_time::ymd_from_serial(serial, date1904);
    return d.m == 2u && d.d > 28u ? date_time::serial_from_ymd(d.y, 2u, 28u, date1904) : serial;
  };
  double days = 0.0;
  if (basis == 0 || basis == 4) {
    days = days_30_360(purchased, first_period, 2, basis == 0, date1904);
  } else {
    days = clipped(first_period) - clipped(purchased);
  }
  double year = 360.0;
  if (basis == 3) {
    year = 365.0;
  } else if (basis == 1) {
    year = date_time::days_in_month(dp.y, 2u) == 29u ? 366.0 : 365.0;
  }
  double first = ((days / year) * rate) * cost;
  if (first == 0.0) {
    first = cost * rate;
  }
  first = std::min(first, cost - salvage);
  const double period = std::trunc(parsed.value().period);
  if (period == 0.0) {
    return finalize(first);
  }
  double remaining = (cost - salvage) - first;
  double charge = rate * cost;
  for (double p = 1.0; p <= period; p += 1.0) {
    charge = std::min(charge, remaining);
    remaining = std::max(remaining - charge, 0.0);
  }
  return finalize(charge);
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
