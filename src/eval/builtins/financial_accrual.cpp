// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

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
  auto issue = read_financial_date(args, 0);
  if (!issue) {
    return Value::error(issue.error());
  }
  auto first_interest = read_financial_date(args, 1);
  if (!first_interest) {
    return Value::error(first_interest.error());
  }
  auto settlement = read_financial_date(args, 2);
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
  auto frequency = read_coupon_frequency(args, 5);
  if (!frequency) {
    return Value::error(frequency.error());
  }
  auto basis = read_day_count_basis(args, arity, 6);
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
  // Mac Excel 365 always accrues from `issue` to `settlement`, ignoring
  // both `first_interest` and `calc_method`. The MS docs say
  // calc_method=FALSE should switch to first_interest, but the actual
  // Mac Excel build (16.108.1, ja-JP) does not — for 1-bit parity we
  // mirror the observable behaviour rather than the docs. The
  // calc_method argument is still validated for type correctness above.
  (void)calc_method;
  (void)first_interest;
  const double start = issue.value();
  auto yf = yearfrac_for_basis(start, settlement.value(), basis.value());
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
  auto issue = read_financial_date(args, 0);
  if (!issue) {
    return Value::error(issue.error());
  }
  auto settlement = read_financial_date(args, 1);
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
  auto basis = read_day_count_basis(args, arity, 4);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (issue.value() >= settlement.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (rate.value() <= 0.0 || par.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = yearfrac_for_basis(issue.value(), settlement.value(), basis.value());
  if (!yf) {
    return Value::error(yf.error());
  }
  return finalize(par.value() * rate.value() * yf.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
