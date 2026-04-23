// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the eager accrued-interest built-ins: ACCRINT and
// ACCRINTM. Registered from `financial.cpp` via
// `register_financial_builtins`.
//
// Both functions follow Excel's documented simple-interest formula
//
//     accrued_interest = par * rate * YEARFRAC(start, settlement, basis)
//
// where `start` is either the issue date (ACCRINT with calc_method=TRUE,
// ACCRINTM) or the first-interest date (ACCRINT with calc_method=FALSE
// and settlement > first_interest).
//
// The `first_interest` and `frequency` arguments of ACCRINT are not used
// in the simple formula but are validated to match Excel's behaviour.

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
// Unlike the DISC/INTRATE helper in `financial_rates.cpp`, a zero-length
// span is allowed and yields 0.0 — ACCRINT/ACCRINTM do not divide by the
// year-fraction, so a zero result simply means "no accrued interest".
Expected<double, ErrorCode> yearfrac_accrual(double start, double end, int basis) {
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

// Reads a required date argument, truncating toward zero. Negative
// serials are rejected as `#NUM!` to match the rest of the date-aware
// builtins.
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

// Reads a Bool / numeric "calc_method" tail argument. TRUE/non-zero
// selects the issue-to-settlement formula; FALSE/zero selects the
// first-interest-to-settlement branch. Missing -> TRUE (Excel default).
Expected<bool, ErrorCode> read_calc_method(const Value* args, std::uint32_t arity, std::uint32_t index) {
  if (arity <= index) {
    return true;
  }
  const Value& v = args[index];
  if (v.is_boolean()) {
    return v.as_boolean();
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  return coerced.value() != 0.0;
}

}  // namespace

// --- ACCRINT(issue, first_interest, settlement, rate, par, frequency,
//             [basis=0], [calc_method=TRUE]) --------------------------------
//
// Accrued interest for a security that pays periodic interest. The
// simple formula used by Excel's calc_method=TRUE branch is:
//
//   ACCRINT = par * rate * YEARFRAC(issue, settlement, basis)
//
// With calc_method=FALSE, the start date becomes `first_interest` when
// `settlement > first_interest` (otherwise it stays `issue`).
//
// `first_interest` and `frequency` are otherwise unused in the simple
// formula; Excel still validates them — frequency must be {1,2,4} and
// the date arguments must be non-negative serials.
//
// Domain:
//   - issue >= settlement            ->  #NUM!
//   - rate <= 0                      ->  #NUM!
//   - par <= 0                       ->  #NUM!
//   - frequency not in {1, 2, 4}     ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}   ->  #NUM!
Value Accrint(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto issue = read_date(args, 0);
  if (!issue) {
    return Value::error(issue.error());
  }
  auto first_interest = read_date(args, 1);
  if (!first_interest) {
    return Value::error(first_interest.error());
  }
  auto settlement = read_date(args, 2);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto rate = read_required_number(args, 3);
  if (!rate) {
    return Value::error(rate.error());
  }
  auto par = read_required_number(args, 4);
  if (!par) {
    return Value::error(par.error());
  }
  auto frequency_e = read_required_number(args, 5);
  if (!frequency_e) {
    return Value::error(frequency_e.error());
  }
  auto basis = read_basis(args, arity, 6);
  if (!basis) {
    return Value::error(basis.error());
  }
  auto calc_method = read_calc_method(args, arity, 7);
  if (!calc_method) {
    return Value::error(calc_method.error());
  }
  if (issue.value() >= settlement.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (rate.value() <= 0.0 || par.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const int freq = static_cast<int>(std::trunc(frequency_e.value()));
  if (freq != 1 && freq != 2 && freq != 4) {
    return Value::error(ErrorCode::Num);
  }
  // calc_method=TRUE  -> accrue from `issue` to `settlement` (the default).
  // calc_method=FALSE -> accrue from `first_interest` to `settlement` when
  //                     `settlement > first_interest`, otherwise from
  //                     `issue` to `settlement`.
  double start = issue.value();
  if (!calc_method.value() && settlement.value() > first_interest.value()) {
    start = first_interest.value();
  }
  auto yf = yearfrac_accrual(start, settlement.value(), basis.value());
  if (!yf) {
    return Value::error(yf.error());
  }
  return finalize(par.value() * rate.value() * yf.value());
}

// --- ACCRINTM(issue, settlement, rate, par, [basis=0]) -----------------
//
// Accrued interest at maturity for a security that pays interest only at
// maturity:
//
//   ACCRINTM = par * rate * YEARFRAC(issue, settlement, basis)
//
// Domain:
//   - issue >= settlement            ->  #NUM!
//   - rate <= 0                      ->  #NUM!
//   - par <= 0                       ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}   ->  #NUM!
Value Accrintm(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto issue = read_date(args, 0);
  if (!issue) {
    return Value::error(issue.error());
  }
  auto settlement = read_date(args, 1);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto rate = read_required_number(args, 2);
  if (!rate) {
    return Value::error(rate.error());
  }
  auto par = read_required_number(args, 3);
  if (!par) {
    return Value::error(par.error());
  }
  auto basis = read_basis(args, arity, 4);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (issue.value() >= settlement.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (rate.value() <= 0.0 || par.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = yearfrac_accrual(issue.value(), settlement.value(), basis.value());
  if (!yf) {
    return Value::error(yf.error());
  }
  return finalize(par.value() * rate.value() * yf.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
