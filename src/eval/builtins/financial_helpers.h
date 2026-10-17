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
#include "eval/excel_math.h"
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
    return std::move(raw.error());
  }
  const double t = std::trunc(raw.value());
  constexpr double kExcelMaxSerial = 2958465.0;
  if (!std::isfinite(t) || t < 0.0 || t > kExcelMaxSerial) {
    return ErrorCode::Num;
  }
  return t;
}

// Reads `args[0..n)` into `out`, where `kinds` holds one character per slot:
// 'n' for a plain number, 'd' for a financial date. All slots are read before
// the caller applies any domain check, and the left-most error wins.
inline Expected<void, ErrorCode> read_required_numbers(const Value* args, const char* kinds, double* out) {
  for (std::uint32_t i = 0; kinds[i] != '\0'; ++i) {
    auto v = kinds[i] == 'd' ? read_financial_date(args, i) : read_required_number(args, i);
    if (!v) {
      return Expected<void, ErrorCode>::Err(v.error());
    }
    out[i] = v.value();
  }
  return Expected<void, ErrorCode>::Ok();
}

// Reads an optional day-count basis argument. Excel truncates numeric basis
// values toward zero and accepts only {0, 1, 2, 3, 4}.
inline Expected<int, ErrorCode> read_day_count_basis(const Value* args, std::uint32_t arity, std::uint32_t index) {
  if (arity <= index) {
    return 0;
  }
  auto raw = read_required_number(args, index);
  if (!raw) {
    return std::move(raw.error());
  }
  const int basis = truncate_saturated_int(raw.value());
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
    return std::move(raw.error());
  }
  const int frequency = truncate_saturated_int(raw.value());
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
  double v[4];
  if (auto read = read_required_numbers(args, "ddnn", v); !read) {
    return std::move(read.error());
  }
  auto basis = read_day_count_basis(args, arity, 4);
  if (!basis) {
    return std::move(basis.error());
  }
  if (v[0] >= v[1]) {
    return ErrorCode::Num;
  }
  if (v[2] <= 0.0 || v[3] <= 0.0) {
    return ErrorCode::Num;
  }
  return SecurityRateArgs{v[0], v[1], v[2], v[3], basis.value()};
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
  double v[3];
  if (auto read = read_required_numbers(args + rate_index, "nnn", v); !read) {
    return std::move(read.error());
  }
  auto frequency = read_coupon_frequency(args, rate_index + 3);
  if (!frequency) {
    return std::move(frequency.error());
  }
  auto basis = read_day_count_basis(args, arity, rate_index + 4);
  if (!basis) {
    return std::move(basis.error());
  }
  if (v[0] < 0.0 || v[2] <= 0.0) {
    return ErrorCode::Num;
  }
  return CouponBondTail{v[0], v[1], v[2], frequency.value(), basis.value()};
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
      yf = date_time::yearfrac_actual_actual(e - s, a.y, a.m, a.d, b.y, b.m, b.d);
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

// `base^exponent` as Excel's financial kernels evaluate it: the shared IEEE
// POWER evaluation, NaN where it reports an error.
inline double excel_power(double base, double exponent) noexcept {
  const auto r = apply_pow_ieee(base, exponent);
  return r ? r.value() : std::numeric_limits<double>::quiet_NaN();
}

inline bool is_feb_end(const date_time::YMD& d) noexcept {
  return d.m == 2u && d.d == date_time::days_in_month(d.y, 2u);
}

// 30/360 day count from `start` to `end`. `mode` 3 forces a February-end or
// 31st end day to 30; other modes follow the US rules. `us == false` is the
// European count.
inline double days_30_360(double start, double end, int mode, bool us, bool date1904) noexcept {
  const date_time::YMD a = date_time::ymd_from_serial(start, date1904);
  const date_time::YMD b = date_time::ymd_from_serial(end, date1904);
  unsigned d1 = a.d;
  unsigned d2 = b.d;
  if (!us) {
    d1 = d1 == 31u ? 30u : d1;
    d2 = d2 == 31u ? 30u : d2;
  } else {
    if (is_feb_end(b) && (is_feb_end(a) || mode == 3)) {
      d2 = 30u;
    }
    if (d2 == 31u && (mode == 3 || d1 >= 30u)) {
      d2 = 30u;
    }
    if (d1 == 31u) {
      d1 = 30u;
    }
    if (is_feb_end(a)) {
      d1 = 30u;
    }
  }
  return static_cast<double>(b.y - a.y) * 360.0 + (static_cast<double>(b.m) - static_cast<double>(a.m)) * 30.0 +
         (static_cast<double>(d2) - static_cast<double>(d1));
}

// A span between two dates as DISC / YIELDMAT count it: the day count and the
// year length it is measured against (`days / year` is the YEARFRAC value).
struct DayCountSpan {
  double days;
  double year;
};

inline Expected<DayCountSpan, ErrorCode> day_count_span(double start, double end, int basis, bool date1904) {
  if (basis < 0 || basis > 4) {
    return ErrorCode::Num;
  }
  const double s = std::trunc(start);
  const double e = std::trunc(end);
  switch (basis) {
    case 0:
    case 4:
      return DayCountSpan{date_time::basis_days_between(s, e, basis, date1904), 360.0};
    case 1: {
      const date_time::YMD a = date_time::ymd_from_serial(s, date1904);
      const date_time::YMD b = date_time::ymd_from_serial(e, date1904);
      return DayCountSpan{e - s, date_time::actual_actual_year_length(a.y, a.m, a.d, b.y, b.m, b.d)};
    }
    case 2:
      return DayCountSpan{e - s, 360.0};
    default:
      return DayCountSpan{e - s, 365.0};
  }
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

// Annuity kernels on Excel's own log1p / expm1.
//
//   H(r, m) = expm1(m * log1p(r))      V(r, m) = exp(m * log1p(r))
//   f       = r (type 0), 1 / ((1 / r) + 1) (type 1)
//   K       = -((pv + fv) / H(r, -n))
inline double annuity_h(double rate, double m) noexcept {
  return excel_expm1(excel_log1p(rate) * m);
}

inline double annuity_v(double rate, double m) noexcept {
  return std::exp(excel_log1p(rate) * m);
}

inline double payment_factor(double rate, double type) noexcept {
  return type != 0.0 ? 1.0 / ((1.0 / rate) + 1.0) : rate;
}

inline double annuity_k(double rate, double nper, double pv, double fv) noexcept {
  return -((pv + fv) / annuity_h(rate, 0.0 - nper));
}

// Internal PMT formula; returns NaN if the inputs are degenerate.
// Callers must check `std::isnan(result)` before returning to the user.
//
//   rate == 0  ->  PMT = (pv + fv) / -nper
//   rate != 0  ->  PMT = (fv - K) * f
inline double pmt_scalar(double rate, double nper, double pv, double fv, double type) noexcept {
  if (rate == 0.0) {
    return (pv + fv) / (0.0 - nper);
  }
  if (!(rate > -1.0)) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  return (fv - annuity_k(rate, nper, pv, fv)) * payment_factor(rate, type);
}

// Shared `per` domain of IPMT / PPMT.
//
// Mac Excel 365 rejects `per < 1` and integer `per > nper` (an
// "out-of-schedule" amortisation period) with `#NUM!`, but evaluates the
// closed form when `per` is fractional -- even for `per > nper`. CUMIPMT /
// CUMPRINC apply their own tighter range checks.
inline bool per_in_schedule(double per, double nper) noexcept {
  return per >= 1.0 && !(per > nper && std::floor(per) == per);
}

// Internal IPMT formula:
//
//   IPMT = (K * H(r, per - (n + 1)) + fv) * f
//
// `rate == 0` accrues no interest, and type 1 pays nothing at per == 1.
// Returns NaN on degenerate inputs (propagated up as `#NUM!`).
inline double ipmt_scalar(double rate, double per, double nper, double pv, double fv, double type) noexcept {
  if (!per_in_schedule(per, nper)) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (rate == 0.0 || (type != 0.0 && per == 1.0)) {
    return 0.0;
  }
  if (!(rate > -1.0)) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  const double k = annuity_k(rate, nper, pv, fv);
  return (k * annuity_h(rate, per - (nper + 1.0)) + fv) * payment_factor(rate, type);
}

// Internal PPMT formula:
//
//   PPMT = ((K * V(r, per - (n + 1))) * -1) * f
//
// and type 1 at per == 1 is PMT itself. Returns NaN on degenerate inputs.
inline double ppmt_scalar(double rate, double per, double nper, double pv, double fv, double type) noexcept {
  if (!per_in_schedule(per, nper)) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  if (rate == 0.0 || (type != 0.0 && per == 1.0)) {
    return pmt_scalar(rate, nper, pv, fv, type);
  }
  if (!(rate > -1.0)) {
    return std::numeric_limits<double>::quiet_NaN();
  }
  const double k = annuity_k(rate, nper, pv, fv);
  return ((k * annuity_v(rate, per - (nper + 1.0))) * -1.0) * payment_factor(rate, type);
}

// Value-returning builtins implemented in `financial_depreciation.cpp`.
Value Sln(const Value* args, std::uint32_t arity, Arena& arena);
Value Syd(const Value* args, std::uint32_t arity, Arena& arena);
Value Ddb(const Value* args, std::uint32_t arity, Arena& arena);
Value Db(const Value* args, std::uint32_t arity, Arena& arena);
Value Vdb(const Value* args, std::uint32_t arity, Arena& arena);
Value Amordegrc(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);
Value Amorlinc(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

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
