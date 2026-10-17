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

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

namespace {

// The `(settlement, maturity, issue, rate, amount, [basis])` arguments of
// PRICEMAT (amount = yld) and YIELDMAT (amount = pr).
struct MaturityInterestArgs {
  double settlement;
  double maturity;
  double issue;
  double rate;
  double amount;
  int basis;
};

// Reads `MaturityInterestArgs`; `#NUM!` unless issue < settlement < maturity
// and rate >= 0. The amount's sign rule is the caller's.
Expected<MaturityInterestArgs, ErrorCode> read_maturity_interest_args(const Value* args, std::uint32_t arity) {
  double v[5];
  if (auto read = read_required_numbers(args, "dddnn", v); !read) {
    return std::move(read.error());
  }
  auto basis = read_day_count_basis(args, arity, 5);
  if (!basis) {
    return std::move(basis.error());
  }
  if (v[2] >= v[0]) {
    return ErrorCode::Num;
  }
  if (v[0] >= v[1]) {
    return ErrorCode::Num;
  }
  if (v[3] < 0.0) {
    return ErrorCode::Num;
  }
  return MaturityInterestArgs{v[0], v[1], v[2], v[3], v[4], basis.value()};
}

// The A / DSM / DIM year fractions of the PRICEMAT / YIELDMAT closed forms.
struct MaturityYearFracs {
  double a;
  double dsm;
  double dim;
};

Expected<MaturityYearFracs, ErrorCode> maturity_year_fracs(const MaturityInterestArgs& in, bool date1904) {
  if (in.basis == 1) {
    // One year length, that of the issue-to-settlement span, divides all three
    // serial differences.
    auto year_span = day_count_span(in.issue, in.settlement, in.basis, date1904);
    if (!year_span) {
      return std::move(year_span.error());
    }
    const double year = year_span.value().year;
    const double issue = std::trunc(in.issue);
    const double settlement = std::trunc(in.settlement);
    const double maturity = std::trunc(in.maturity);
    return MaturityYearFracs{(settlement - issue) / year, (maturity - settlement) / year, (maturity - issue) / year};
  }
  auto a_yf = yearfrac_for_basis(in.issue, in.settlement, in.basis, date1904);
  if (!a_yf) {
    return std::move(a_yf.error());
  }
  auto dsm_yf = yearfrac_for_basis(in.settlement, in.maturity, in.basis, date1904);
  if (!dsm_yf) {
    return std::move(dsm_yf.error());
  }
  auto dim_yf = yearfrac_for_basis(in.issue, in.maturity, in.basis, date1904);
  if (!dim_yf) {
    return std::move(dim_yf.error());
  }
  return MaturityYearFracs{a_yf.value(), dsm_yf.value(), dim_yf.value()};
}

}  // namespace

// --- PRICEDISC(settlement, maturity, discount, redemption, [basis=0]) --
//
// Price per $100 face value of a discounted security:
//
//   PRICEDISC = redemption - ((discount * redemption) * DSM) / B
//
// Domain per Microsoft docs:
//   - settlement >= maturity         ->  #NUM!
//   - discount <= 0                  ->  #NUM!
//   - redemption <= 0                ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}   ->  #NUM!
Value PriceDisc(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [settlement, maturity, discount, redemption, basis] = parsed.value();
  auto span = day_count_span(settlement, maturity, basis, date1904);
  if (!span) {
    return Value::error(span.error());
  }
  const double result = -(((discount * redemption) * span.value().days) / span.value().year) + redemption;
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
// Under basis 1 the three terms are serial differences divided by one year
// length B, that of the issue-to-settlement span (`maturity_year_fracs`).
//
// Domain per Microsoft docs:
//   - issue >= settlement                 ->  #NUM!
//   - settlement >= maturity              ->  #NUM!
//   - rate < 0                            ->  #NUM!
//   - yld  < 0                            ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}        ->  #NUM!
//   - 1 + DSM * yld == 0 (degenerate)     ->  #NUM!
Value PriceMat(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_maturity_interest_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const double rate = parsed.value().rate;
  const double yld = parsed.value().amount;
  if (yld < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto yf = maturity_year_fracs(parsed.value(), date1904);
  if (!yf) {
    return Value::error(yf.error());
  }
  const auto [a, dsm, dim] = yf.value();
  const double denom = 1.0 + dsm * yld;
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (100.0 + dim * rate * 100.0) / denom - a * rate * 100.0;
  return finalize(result);
}

// --- YIELDDISC(settlement, maturity, pr, redemption, [basis=0]) --------
//
// Annual yield for a discounted security:
//
//   YIELDDISC = (B / DSM) * ((redemption - pr) / pr)
//
// Domain per Microsoft docs:
//   - settlement >= maturity          ->  #NUM!
//   - pr <= 0                         ->  #NUM!
//   - redemption <= 0                 ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}    ->  #NUM!
Value YieldDisc(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_security_rate_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const auto [settlement, maturity, pr, redemption, basis] = parsed.value();
  auto span = day_count_span(settlement, maturity, basis, date1904);
  if (!span) {
    return Value::error(span.error());
  }
  if (span.value().days <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = (span.value().year / span.value().days) * ((redemption - pr) / pr);
  return finalize(result);
}

// --- YIELDMAT(settlement, maturity, issue, rate, pr, [basis=0]) --------
//
// Annual yield of a security that pays interest at maturity:
//
//   A   = YEARFRAC(issue, settlement, basis)
//   DSM = YEARFRAC(settlement, maturity, basis)
//   DIM = YEARFRAC(issue, maturity, basis)
//   YIELDMAT = ((1 + DIM/B*rate) - (pr/100 + A/B*rate)) / (pr/100 + A/B*rate) * (B/DSM)
//
// with A / DSM / DIM as day counts and B the year length of the basis (under
// basis 1, the issue-to-settlement year length for all three terms).
//
// Domain per Microsoft docs:
//   - issue >= settlement              ->  #NUM!
//   - settlement >= maturity           ->  #NUM!
//   - rate < 0                         ->  #NUM!
//   - pr <= 0                          ->  #NUM!
//   - basis not in {0, 1, 2, 3, 4}     ->  #NUM!
//   - pr/100 + A * rate == 0           ->  #NUM!
//   - DSM == 0                         ->  #NUM!
Value YieldMat(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto parsed = read_maturity_interest_args(args, arity);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const double rate = parsed.value().rate;
  const double pr = parsed.value().amount;
  if (pr <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const MaturityInterestArgs& in = parsed.value();
  auto a_span = day_count_span(in.issue, in.settlement, in.basis, date1904);
  auto dsm_span = day_count_span(in.settlement, in.maturity, in.basis, date1904);
  auto dim_span = day_count_span(in.issue, in.maturity, in.basis, date1904);
  if (!a_span || !dsm_span || !dim_span) {
    return Value::error(ErrorCode::Num);
  }
  const double a_days = a_span.value().days;
  const double a_year = a_span.value().year;
  const double dsm_days = dsm_span.value().days;
  // Under basis 1 one year length, that of the issue-to-settlement span, serves all three terms.
  const double dim_year = in.basis == 1 ? a_year : dim_span.value().year;
  const double dsm_year = in.basis == 1 ? a_year : dsm_span.value().year;
  if (dsm_days == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double accrued = pr / 100.0 + a_days / a_year * rate;
  if (accrued == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = ((1.0 + dim_span.value().days / dim_year * rate) - accrued) / accrued * (dsm_year / dsm_days);
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
