// Copyright 2026 libraz. Licensed under the MIT License.
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

#include "eval/coerce.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

// Reads an optional trailing numeric argument at position `index`, falling
// back to `default_value` when `arity <= index`. Non-finite coerced values
// surface as `#NUM!` to match the rest of the math-family argument-handling
// conventions. The default is returned by value rather than by reference so
// callers never need to worry about lifetime of a sentinel.
inline Expected<double, ErrorCode> read_optional_number(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                        double default_value) {
  if (arity <= index) {
    return default_value;
  }
  auto coerced = coerce_to_number(args[index]);
  if (!coerced) {
    return coerced.error();
  }
  const double v = coerced.value();
  if (std::isnan(v) || std::isinf(v)) {
    return ErrorCode::Num;
  }
  return v;
}

// Reads a required numeric argument at position `index`. Parallels
// `read_optional_number` above; kept as a separate helper to keep the
// call sites readable (no "magic default that won't be used").
inline Expected<double, ErrorCode> read_required_number(const Value* args, std::uint32_t index) {
  auto coerced = coerce_to_number(args[index]);
  if (!coerced) {
    return coerced.error();
  }
  const double v = coerced.value();
  if (std::isnan(v) || std::isinf(v)) {
    return ErrorCode::Num;
  }
  return v;
}

// Finalises a scalar financial result: a non-finite value becomes `#NUM!`,
// otherwise wraps the double in a `Value::number`.
inline Value finalize(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
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
inline double ipmt_scalar(double rate, double per, double nper, double pv, double fv, double type) noexcept {
  if (per < 1.0 || per > nper) {
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
Value Amordegrc(const Value* args, std::uint32_t arity, Arena& arena);
Value Amorlinc(const Value* args, std::uint32_t arity, Arena& arena);

// Value-returning builtins implemented in `financial_accrual.cpp`.
Value Accrint(const Value* args, std::uint32_t arity, Arena& arena);
Value Accrintm(const Value* args, std::uint32_t arity, Arena& arena);

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
Value Disc(const Value* args, std::uint32_t arity, Arena& arena);
Value Intrate(const Value* args, std::uint32_t arity, Arena& arena);
Value Received(const Value* args, std::uint32_t arity, Arena& arena);
Value TBillPrice(const Value* args, std::uint32_t arity, Arena& arena);
Value TBillYield(const Value* args, std::uint32_t arity, Arena& arena);
Value TBillEq(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_HELPERS_H_
