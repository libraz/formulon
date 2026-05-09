// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the closed-form bond-pricing built-ins:
// PRICEDISC, PRICEMAT, YIELDDISC, YIELDMAT — and the STOCKHISTORY stub.
// Registered from `financial.cpp` via `register_financial_builtins`.
//
// All four closed-form helpers share the following conventions with the
// sibling `financial_rates.cpp` TU:
//   * Date arguments are Excel serial numbers (doubles); the integer part
//     is taken with `std::trunc` before use.
//   * Date ordering is validated per the Microsoft documentation:
//     PRICEDISC / YIELDDISC require `settlement < maturity`; PRICEMAT /
//     YIELDMAT additionally require `issue < settlement`.
//   * `basis` (where present) must be in {0, 1, 2, 3, 4} after truncation.
//
// STOCKHISTORY is an intentional stub: Formulon is a pure calculation
// engine and does not perform network or market-data I/O. The stub body
// follows the WEBSERVICE/PY pattern from `web.cpp` — it relies on the
// dispatcher's eager-argument evaluation to propagate any error inside
// an argument before its fixed `#VALUE!` return fires.

#include "eval/builtins/financial_bond_simple.h"

#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

// --- PRICEDISC(settlement, maturity, discount, redemption, [basis=0]) --
//
// Price per $100 face value of a discounted security:
//
//   PRICEDISC = redemption - discount * redemption * YEARFRAC(settlement, maturity, basis)
//
// Domain per Microsoft docs:
//   - settlement >= maturity         ->  #NUM!
//   - discount <= 0                  ->  #NUM!
//   - redemption <= 0                ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}   ->  #NUM!
Value PriceDisc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto discount = read_required_number(args, 2);
  if (!discount) {
    return Value::error(discount.error());
  }
  auto redemption = read_required_number(args, 3);
  if (!redemption) {
    return Value::error(redemption.error());
  }
  auto basis = read_day_count_basis(args, arity, 4);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (settlement.value() >= maturity.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (discount.value() <= 0.0 || redemption.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = yearfrac_for_basis(settlement.value(), maturity.value(), basis.value());
  if (!yf) {
    return Value::error(yf.error());
  }
  // settlement < maturity guarantees a strictly positive yearfrac for
  // every supported basis; there is no divide here, so a degenerate
  // yearfrac is not a concern.
  const double result = redemption.value() - discount.value() * redemption.value() * yf.value();
  return finalize(result);
}

// --- PRICEMAT(settlement, maturity, issue, rate, yld, [basis=0]) -------
//
// Price per $100 face value of a security that pays interest at maturity.
// The `100` in the formula is the implicit par value.
//
//   A   = YEARFRAC(issue, settlement, basis)
//   DSM = YEARFRAC(settlement, maturity, basis)
//   DIM = YEARFRAC(issue, maturity, basis)
//   PRICEMAT = (100 + DIM * rate * 100) / (1 + DSM * yld) - A * rate * 100
//
// Domain per Microsoft docs:
//   - issue >= settlement                 ->  #NUM!
//   - settlement >= maturity              ->  #NUM!
//   - rate < 0                            ->  #NUM!
//   - yld  < 0                            ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}        ->  #NUM!
//   - 1 + DSM * yld == 0 (degenerate)     ->  #NUM!
Value PriceMat(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto issue = read_financial_date(args, 2);
  if (!issue) {
    return Value::error(issue.error());
  }
  auto rate = read_required_number(args, 3);
  if (!rate) {
    return Value::error(rate.error());
  }
  auto yld = read_required_number(args, 4);
  if (!yld) {
    return Value::error(yld.error());
  }
  auto basis = read_day_count_basis(args, arity, 5);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (issue.value() >= settlement.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (settlement.value() >= maturity.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (rate.value() < 0.0 || yld.value() < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto a_yf = yearfrac_for_basis(issue.value(), settlement.value(), basis.value());
  if (!a_yf) {
    return Value::error(a_yf.error());
  }
  auto dsm_yf = yearfrac_for_basis(settlement.value(), maturity.value(), basis.value());
  if (!dsm_yf) {
    return Value::error(dsm_yf.error());
  }
  auto dim_yf = yearfrac_for_basis(issue.value(), maturity.value(), basis.value());
  if (!dim_yf) {
    return Value::error(dim_yf.error());
  }
  const double denom = 1.0 + dsm_yf.value() * yld.value();
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (100.0 + dim_yf.value() * rate.value() * 100.0) / denom - a_yf.value() * rate.value() * 100.0;
  return finalize(result);
}

// --- YIELDDISC(settlement, maturity, pr, redemption, [basis=0]) --------
//
// Annual yield for a discounted security:
//
//   YIELDDISC = ((redemption - pr) / pr) / YEARFRAC(settlement, maturity, basis)
//
// Domain per Microsoft docs:
//   - settlement >= maturity          ->  #NUM!
//   - pr <= 0                         ->  #NUM!
//   - redemption <= 0                 ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}    ->  #NUM!
Value YieldDisc(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto pr = read_required_number(args, 2);
  if (!pr) {
    return Value::error(pr.error());
  }
  auto redemption = read_required_number(args, 3);
  if (!redemption) {
    return Value::error(redemption.error());
  }
  auto basis = read_day_count_basis(args, arity, 4);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (settlement.value() >= maturity.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (pr.value() <= 0.0 || redemption.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = yearfrac_for_basis(settlement.value(), maturity.value(), basis.value());
  if (!yf) {
    return Value::error(yf.error());
  }
  if (yf.value() == 0.0) {
    // settlement < maturity is enforced above, but a degenerate basis
    // could still yield 0 (not with basis 0..4, but guard anyway).
    return Value::error(ErrorCode::Num);
  }
  const double result = ((redemption.value() - pr.value()) / pr.value()) / yf.value();
  return finalize(result);
}

// --- YIELDMAT(settlement, maturity, issue, rate, pr, [basis=0]) --------
//
// Annual yield of a security that pays interest at maturity:
//
//   A   = YEARFRAC(issue, settlement, basis)
//   DSM = YEARFRAC(settlement, maturity, basis)
//   DIM = YEARFRAC(issue, maturity, basis)
//   YIELDMAT = ((1 + DIM * rate) / (pr/100 + A * rate) - 1) / DSM
//
// Domain per Microsoft docs:
//   - issue >= settlement              ->  #NUM!
//   - settlement >= maturity           ->  #NUM!
//   - rate < 0                         ->  #NUM!
//   - pr <= 0                          ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}     ->  #NUM!
//   - pr/100 + A * rate == 0           ->  #NUM!
//   - DSM == 0                         ->  #NUM!
Value YieldMat(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return Value::error(settlement.error());
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return Value::error(maturity.error());
  }
  auto issue = read_financial_date(args, 2);
  if (!issue) {
    return Value::error(issue.error());
  }
  auto rate = read_required_number(args, 3);
  if (!rate) {
    return Value::error(rate.error());
  }
  auto pr = read_required_number(args, 4);
  if (!pr) {
    return Value::error(pr.error());
  }
  auto basis = read_day_count_basis(args, arity, 5);
  if (!basis) {
    return Value::error(basis.error());
  }
  if (issue.value() >= settlement.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (settlement.value() >= maturity.value()) {
    return Value::error(ErrorCode::Num);
  }
  if (rate.value() < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (pr.value() <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto a_yf = yearfrac_for_basis(issue.value(), settlement.value(), basis.value());
  if (!a_yf) {
    return Value::error(a_yf.error());
  }
  auto dsm_yf = yearfrac_for_basis(settlement.value(), maturity.value(), basis.value());
  if (!dsm_yf) {
    return Value::error(dsm_yf.error());
  }
  auto dim_yf = yearfrac_for_basis(issue.value(), maturity.value(), basis.value());
  if (!dim_yf) {
    return Value::error(dim_yf.error());
  }
  if (dsm_yf.value() == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double denom = pr.value() / 100.0 + a_yf.value() * rate.value();
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = ((1.0 + dim_yf.value() * rate.value()) / denom - 1.0) / dsm_yf.value();
  return finalize(result);
}

// --- STOCKHISTORY(stock, start_date, [end_date], [interval], [headers], [properties...]) ---
//
// Always returns `#VALUE!`. Excel's STOCKHISTORY retrieves historical
// market data from a cloud data service; Formulon is a pure calculation
// engine with no network or market-data integration, so there is no
// realistic value to return. This matches the pattern used by
// WEBSERVICE / PY in `web.cpp`: the stub fires *after* the dispatcher has
// eagerly evaluated every argument, so an error inside any argument (e.g.
// `1/0`) still short-circuits via the default `propagate_errors = true`
// dispatch flag before we are called.
Value StockHistory(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Value);
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
