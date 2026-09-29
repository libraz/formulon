//
// Internal header — do not include outside `src/eval/builtins/financial*`.
//
// Shared argument-coercion helpers, scalar TVM primitives, and forward
// declarations of the Value-returning builtins that live in the sibling
// `financial_depreciation.cpp` and `financial_misc.cpp` translation units.
// Keeping these `inline` in the header (rather than duplicating them
// across TUs) ensures the time-value-of-money scalar formulas stay in
// lock-step.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_HELPERS_H_

#include <cmath>
#include <cstdint>
#include <limits>

#include "eval/builtins/numeric_helpers.h"
#include "eval/coerce.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

// Reads an optional trailing numeric argument at position `index`, falling
// back to `default_value` when `arity <= index`. Non-finite coerced values
// surface as `#NUM!` to match the rest of the math-family argument-handling
// conventions. The default is returned by value rather than by reference so
// callers never need to worry about lifetime of a sentinel. Thin alias to
// the shared `builtins_detail::read_optional_number` helper.
inline Expected<double, ErrorCode> read_optional_number(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                        double default_value) {
  return builtins_detail::read_optional_number(args, arity, index, default_value);
}

// Reads a required numeric argument at position `index`. Parallels
// `read_optional_number` above; kept as a separate helper to keep the
// call sites readable (no "magic default that won't be used"). Thin
// alias to the shared `builtins_detail::read_required_number` helper.
inline Expected<double, ErrorCode> read_required_number(const Value* args, std::uint32_t index) {
  return builtins_detail::read_required_number(args, index);
}

// Ceiling on how many periods a depreciation schedule may be stepped
// through. `life` is an unconstrained double, and DB / VDB walk one
// iteration per period, so `=DB(1, 0, 1E18, 1E18)` would otherwise spin for
// longer than the process will live. One Excel column's worth of periods is
// far beyond any real schedule — a period index past it could not even be
// laid out in a worksheet — so a longer request is rejected as `#NUM!`
// rather than accepted and never answered. DDB needs no such cap: it has a
// closed form for the same schedule.
inline constexpr double kMaxDepreciationPeriods = 1048576.0;

// Reads a required date argument, truncating toward zero. Dates outside
// Excel's supported serial range [0, 2958465] (9999-12-31) are rejected as
// `#NUM!` before coupon / day-count loops consume them.
inline Expected<double, ErrorCode> read_financial_date(const Value* args, std::uint32_t index) {
  auto raw = read_required_number(args, index);
  if (!raw) {
    return raw.error();
  }
  const double t = std::trunc(raw.value());
  constexpr double kExcelMaxSerial = 2958465.0;
  if (!std::isfinite(t) || t < 0.0 || t > kExcelMaxSerial) {
    return ErrorCode::Num;
  }
  return t;
}

// Reads an optional day-count basis argument. Excel truncates numeric basis
// values toward zero and accepts only {0, 1, 2, 3, 4}.
inline Expected<int, ErrorCode> read_day_count_basis(const Value* args, std::uint32_t arity, std::uint32_t index) {
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

// Reads a required coupon frequency argument. Excel truncates the numeric
// value and accepts annual, semi-annual, and quarterly schedules only.
inline Expected<int, ErrorCode> read_coupon_frequency(const Value* args, std::uint32_t index) {
  auto raw = read_required_number(args, index);
  if (!raw) {
    return raw.error();
  }
  const int frequency = static_cast<int>(std::trunc(raw.value()));
  if (frequency != 1 && frequency != 2 && frequency != 4) {
    return ErrorCode::Num;
  }
  return frequency;
}

// The `(settlement, maturity, amount1, amount2, [basis])` arguments shared by
// the discounted-security family (DISC, INTRATE, RECEIVED, PRICEDISC,
// YIELDDISC, and ACCRINTM with issue / settlement as its two dates).
struct SecurityRateArgs {
  double settlement;
  double maturity;
  double amount1;
  double amount2;
  int basis;
};

// Reads `SecurityRateArgs` from positions 0..4. `#NUM!` unless
// `settlement < maturity` and both amounts are positive.
inline Expected<SecurityRateArgs, ErrorCode> read_security_rate_args(const Value* args, std::uint32_t arity) {
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
  if (amount1.value() <= 0.0 || amount2.value() <= 0.0) {
    return ErrorCode::Num;
  }
  return SecurityRateArgs{settlement.value(), maturity.value(), amount1.value(), amount2.value(), basis.value()};
}

// The `(rate, amount, redemption, frequency, [basis])` argument tail shared
// by PRICE / YIELD and the ODDF* / ODDL* bond family, where `amount` is the
// yld or pr slot.
struct CouponBondTail {
  double rate;
  double amount;
  double redemption;
  int frequency;
  int basis;
};

// Reads `CouponBondTail` starting at `args[rate_index]`. `#NUM!` on a
// negative rate or a non-positive redemption; the amount's sign rule is
// the caller's.
inline Expected<CouponBondTail, ErrorCode> read_coupon_bond_tail(const Value* args, std::uint32_t arity,
                                                                 std::uint32_t rate_index) {
  auto rate = read_required_number(args, rate_index);
  if (!rate) {
    return rate.error();
  }
  auto amount = read_required_number(args, rate_index + 1);
  if (!amount) {
    return amount.error();
  }
  auto redemption = read_required_number(args, rate_index + 2);
  if (!redemption) {
    return redemption.error();
  }
  auto frequency = read_coupon_frequency(args, rate_index + 3);
  if (!frequency) {
    return frequency.error();
  }
  auto basis = read_day_count_basis(args, arity, rate_index + 4);
  if (!basis) {
    return basis.error();
  }
  if (rate.value() < 0.0 || redemption.value() <= 0.0) {
    return ErrorCode::Num;
  }
  return CouponBondTail{rate.value(), amount.value(), redemption.value(), frequency.value(), basis.value()};
}

// Computes YEARFRAC(start, end, basis) under the same rules as the
// YEARFRAC builtin. Allows a zero result; callers that divide by the
// year fraction should reject zero before use. `date1904` must be the
// calling workbook's date system so `start` / `end` decode to the right
// calendar day.
inline Expected<double, ErrorCode> yearfrac_for_basis(double start, double end, int basis, bool date1904) {
  if (basis < 0 || basis > 4) {
    return ErrorCode::Num;
  }
  const double s = std::trunc(start);
  const double e = std::trunc(end);
  const date_time::YMD a = date_time::ymd_from_serial(s, date1904);
  const date_time::YMD b = date_time::ymd_from_serial(e, date1904);
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

// Finalises a scalar financial result: a non-finite value becomes `#NUM!`,
// otherwise wraps the double in a `Value::number`. Thin alias to the
// shared `builtins_detail::to_finite_value` helper.
inline Value finalize(double r) {
  return builtins_detail::to_finite_value(r);
}

// Normalises the `type` argument (end-vs-begin of period). Excel accepts
// any numeric value; zero means end-of-period, any non-zero value means
// begin-of-period. We mirror that here so the impl formulas can use `type`
// as a 0-or-1 multiplier without a secondary branch.
inline double normalize_type(double t) noexcept {
  return t == 0.0 ? 0.0 : 1.0;
}

// --- Internal scalar helpers -------------------------------------------
//
// The IPMT / PPMT / CUMIPMT / CUMPRINC / RATE impls all need to compute
// PMT and FV at arbitrary (rate, nper[, pv, fv, type]) tuples without
// going back through the argument-coercion machinery. These helpers
// compute directly from doubles and return NaN on degenerate cases; the
// callers translate NaN into `#NUM!` via `finalize`.

// Internal PMT formula; returns NaN if the inputs are degenerate.
// Callers must check `std::isnan(result)` before returning to the user.
inline double pmt_scalar(double rate, double nper, double pv, double fv, double type) noexcept {
  if (rate == -1.0) {
    // Denominator (1+rate)*(pow_term-1) vanishes; caller surfaces as #NUM!
    // (scalar layer cannot distinguish DIV/0 from NUM — the Value-level
    // impls catch the rate==-1 case explicitly to return #DIV/0!).
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (rate == 0.0) {
    if (nper == 0.0) {
      return std::numeric_limits<double>::quiet_NaN();
    }
    return -(pv + fv) / nper;
  }
  const double pow_term = std::pow(1.0 + rate, nper);
  const double denom = (1.0 + rate * type) * (pow_term - 1.0);
  if (denom == 0.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  return -(pv * pow_term + fv) * rate / denom;
}

// Internal FV formula. Matches the Fv() public impl above; extracted so
// IPMT / PPMT can compute balances mid-schedule without re-parsing args.
inline double fv_scalar(double rate, double nper, double pmt, double pv, double type) noexcept {
  if (rate == -1.0) {
    // Division by rate in the closed form; caller surfaces as #NUM!
    // (scalar callers — IPMT/PPMT/CUMIPMT/CUMPRINC — treat the degenerate
    // rate identically to other NaN-producing inputs).
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (rate == 0.0) {
    return -(pv + pmt * nper);
  }
  const double pow_term = std::pow(1.0 + rate, nper);
  return -(pv * pow_term + pmt * (1.0 + rate * type) * (pow_term - 1.0) / rate);
}

// Internal IPMT formula used by IPMT / PPMT / CUMIPMT / CUMPRINC. Per
// Microsoft's documentation:
//
//   - type == 0: the balance at the start of period p is the running
//     `FV(rate, p-1, pmt, pv, 0)` — a signed Excel FV, so for a loan
//     with positive pv the value comes back *negative* (representing
//     the remaining debt from the lender's perspective). Interest
//     charged for the period is `balance * rate` — negative when pv is
//     positive, matching Excel's "interest is cash out from the
//     borrower" sign convention.
//   - type == 1 && p == 1: no interest has accrued on the first period's
//     start-of-period payment, so IPMT is 0.
//   - type == 1 && p  > 1: the balance-at-start is `FV(rate, p-1, pmt,
//     pv, 1)`, and Excel divides interest by `(1 + rate)` because part
//     of the accrued interest is paid at the start of the period rather
//     than the end.
//
// If `rate == 0`, no interest ever accrues, so IPMT is 0 for any period.
// Returns NaN on degenerate inputs (propagated up as `#NUM!`).
//
// Note: Mac Excel 365 (and the IronCalc oracle) reject `per < 1` and
// integer `per > nper` (an "out-of-schedule" amortisation period) with
// `#NUM!`, but evaluate the closed-form formula directly when `per` is
// fractional — even for `per > nper`. We mirror that: a fractional `per`
// is always accepted as long as it is >= 1. CUMIPMT / CUMPRINC apply their
// own tighter range checks before calling this helper.
inline double ipmt_scalar(double rate, double per, double nper, double pv, double fv, double type) noexcept {
  if (per < 1.0) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  // Reject integer per > nper as an out-of-schedule period. Fractional per
  // is accepted for any value >= 1 (matches Mac Excel 365 oracle).
  if (per > nper && std::floor(per) == per) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (rate == 0.0) {
    return 0.0;
  }
  const double pmt = pmt_scalar(rate, nper, pv, fv, type);
  if (std::isnan(pmt)) {
    return pmt;
  }
  if (type == 1.0 && per == 1.0) {
    return 0.0;
  }
  // fv_scalar follows Excel's signed convention, so for a loan with
  // positive pv the returned "balance" is already negative. Multiplying
  // by rate directly gives Excel's signed interest (negative for a
  // borrower paying interest out).
  const double balance = fv_scalar(rate, per - 1.0, pmt, pv, type);
  const double interest = balance * rate;
  if (type == 1.0) {
    // per > 1 here by the branch above.
    return interest / (1.0 + rate);
  }
  return interest;
}

// Value-returning builtins implemented in `financial_depreciation.cpp`.
Value Sln(const Value* args, std::uint32_t arity, Arena& arena);
Value Syd(const Value* args, std::uint32_t arity, Arena& arena);
Value Ddb(const Value* args, std::uint32_t arity, Arena& arena);
Value Db(const Value* args, std::uint32_t arity, Arena& arena);
Value Vdb(const Value* args, std::uint32_t arity, Arena& arena);
Value Amordegrc(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value Amorlinc(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

// A bond price kernel evaluated at one candidate yield, in the shape the
// yield solvers below invert: the call's own argument vector with the
// yield slot replaced by `yld`. `date1904` is the calling workbook's date
// system, needed by any day-count decomposition the kernel performs.
using BondPriceFn = Expected<double, ErrorCode> (*)(const Value* args, std::uint32_t arity, double yld, bool date1904);

// Solves `price_at(yld) == target_price` by Newton-Raphson, shared by
// YIELD and ODDFYIELD. The derivative is a central difference with step
// `1e-7 * max(1, |yld|)`, one-sided while the iterate sits within one
// step of zero, because every price kernel here has domain `yld >= 0`
// and a damped step keeps the iterate on that side of the boundary.
//
// Converges on `|price - target| < 1e-12 * (|target| + 1)` or on a step
// below `1e-15`. A degenerate derivative, a non-finite iterate, or
// exhausting the iteration cap surfaces `#NUM!`, which is what Excel
// reports for a price the bond cannot reach.
inline Expected<double, ErrorCode> solve_yield_by_newton(BondPriceFn price_at, const Value* args, std::uint32_t arity,
                                                         double target_price, double initial_guess, bool date1904) {
  constexpr int kMaxIter = 100;
  constexpr double kStepTol = 1.0e-15;
  const double f_tol = 1.0e-12 * (std::fabs(target_price) + 1.0);
  double yld = initial_guess;
  for (int iter = 0; iter < kMaxIter; ++iter) {
    auto f0 = price_at(args, arity, yld, date1904);
    if (!f0) {
      return f0.error();
    }
    const double residual = f0.value() - target_price;
    if (std::fabs(residual) < f_tol) {
      if (std::isnan(yld) || std::isinf(yld)) {
        return ErrorCode::Num;
      }
      return yld;
    }
    const double step = 1.0e-7 * std::fmax(1.0, std::fabs(yld));
    auto f_plus = price_at(args, arity, yld + step, date1904);
    if (!f_plus) {
      return f_plus.error();
    }
    double df = 0.0;
    if (yld < step) {
      df = (f_plus.value() - f0.value()) / step;
    } else {
      auto f_minus = price_at(args, arity, yld - step, date1904);
      if (!f_minus) {
        return f_minus.error();
      }
      df = (f_plus.value() - f_minus.value()) / (2.0 * step);
    }
    if (df == 0.0 || std::isnan(df) || std::isinf(df)) {
      return ErrorCode::Num;
    }
    double damped_delta = residual / df;
    double new_yld = yld - damped_delta;
    while (new_yld < 0.0 && std::fabs(damped_delta) >= kStepTol) {
      damped_delta *= 0.5;
      new_yld = yld - damped_delta;
    }
    new_yld = std::fmax(0.0, new_yld);
    if (std::isnan(new_yld) || std::isinf(new_yld)) {
      return ErrorCode::Num;
    }
    if (std::fabs(damped_delta) < kStepTol) {
      // The step underflowed without the top-of-loop residual check ever
      // firing -- typically because the iterate is pinned at the yld >= 0
      // domain boundary and can shrink no further. Step-size convergence
      // alone does not mean the target price was reached, so certify the
      // residual here too before accepting the answer.
      auto f_new = price_at(args, arity, new_yld, date1904);
      if (!f_new) {
        return f_new.error();
      }
      if (std::fabs(f_new.value() - target_price) < f_tol) {
        return new_yld;
      }
      return ErrorCode::Num;
    }
    yld = new_yld;
  }
  // Iteration cap reached without convergence.
  return ErrorCode::Num;
}

// Value-returning builtins implemented in `financial_accrual.cpp`.
Value Accrint(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value Accrintm(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

// Value-returning builtins implemented in `financial_misc.cpp`.
Value DollarDe(const Value* args, std::uint32_t arity, Arena& arena);
Value DollarFr(const Value* args, std::uint32_t arity, Arena& arena);
Value Effect(const Value* args, std::uint32_t arity, Arena& arena);
Value Nominal(const Value* args, std::uint32_t arity, Arena& arena);
Value FvSchedule(const Value* args, std::uint32_t arity, Arena& arena);
Value PDuration(const Value* args, std::uint32_t arity, Arena& arena);
Value Rri(const Value* args, std::uint32_t arity, Arena& arena);
Value IsPmt(const Value* args, std::uint32_t arity, Arena& arena);

// Value-returning builtins implemented in `financial_rates.cpp`
// (security-rate and T-Bill family).
Value Disc(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value Intrate(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value Received(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value TBillPrice(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value TBillYield(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value TBillEq(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_HELPERS_H_
