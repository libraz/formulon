// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the eager financial built-ins: PV, FV, PMT, NPER,
// NPV, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC. The depreciation family
// (SLN / SYD / DDB / DB) lives in `financial_depreciation.cpp`; the
// DOLLARDE / DOLLARFR / EFFECT / NOMINAL / FVSCHEDULE / PDURATION /
// RRI / ISPMT group lives in `financial_misc.cpp`. All three
// translation units share the `read_required_number` / `finalize` /
// `pmt_scalar` / ... helpers defined inline in
// `builtins/financial_helpers.h`.
//
// Each function follows Excel's time-value-of-money sign convention
// (cash out = negative, cash in = positive) and defaults to
// end-of-period payments (`type = 0`). See
// `backup/plans/02-calc-engine.md` for the Excel-compatibility
// reference.
//
// IRR is intentionally absent — it needs the un-flattened AST of its
// values argument to walk range/Ref/ArrayLiteral shapes, so it lives on
// the lazy-dispatch seam in `eval/financial_lazy.cpp`.

#include "eval/builtins/financial.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_bond_simple.h"
#include "eval/builtins/financial_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using financial_detail::finalize;
using financial_detail::fv_scalar;
using financial_detail::ipmt_scalar;
using financial_detail::normalize_type;
using financial_detail::pmt_scalar;
using financial_detail::read_optional_number;
using financial_detail::read_required_number;

// --- PV(rate, nper, pmt, [fv=0], [type=0]) ------------------------------
//
// Matches Excel's documented formula:
//
//   rate == 0  ->  PV = -(pmt * nper + fv)
//   rate != 0  ->  f  = (1 + rate)^nper
//                  PV = -((fv + pmt * (1 + rate*type) * (f - 1) / rate) / f)
//
// Sign convention follows Excel: cash-out values are negative, cash-in
// values are positive. Non-finite intermediate or final values surface as
// `#NUM!`.
Value Pv(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto rate = read_required_number(args, 0);
  if (!rate) {
    return Value::error(rate.error());
  }
  auto nper = read_required_number(args, 1);
  if (!nper) {
    return Value::error(nper.error());
  }
  auto pmt = read_required_number(args, 2);
  if (!pmt) {
    return Value::error(pmt.error());
  }
  auto fv = read_optional_number(args, arity, 3, 0.0);
  if (!fv) {
    return Value::error(fv.error());
  }
  auto type = read_optional_number(args, arity, 4, 0.0);
  if (!type) {
    return Value::error(type.error());
  }
  const double r = rate.value();
  const double n = nper.value();
  const double p = pmt.value();
  const double f = fv.value();
  const double t = normalize_type(type.value());
  if (r == 0.0) {
    return finalize(-(p * n + f));
  }
  const double pow_term = std::pow(1.0 + r, n);
  if (pow_term == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = -((f + p * (1.0 + r * t) * (pow_term - 1.0) / r) / pow_term);
  return finalize(result);
}

// --- FV(rate, nper, pmt, [pv=0], [type=0]) ------------------------------
//
//   rate == 0  ->  FV = -(pv + pmt * nper)
//   rate != 0  ->  f  = (1 + rate)^nper
//                  FV = -(pv * f + pmt * (1 + rate*type) * (f - 1) / rate)
Value Fv(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto rate = read_required_number(args, 0);
  if (!rate) {
    return Value::error(rate.error());
  }
  auto nper = read_required_number(args, 1);
  if (!nper) {
    return Value::error(nper.error());
  }
  auto pmt = read_required_number(args, 2);
  if (!pmt) {
    return Value::error(pmt.error());
  }
  auto pv = read_optional_number(args, arity, 3, 0.0);
  if (!pv) {
    return Value::error(pv.error());
  }
  auto type = read_optional_number(args, arity, 4, 0.0);
  if (!type) {
    return Value::error(type.error());
  }
  const double r = rate.value();
  const double n = nper.value();
  const double p = pmt.value();
  const double v = pv.value();
  const double t = normalize_type(type.value());
  if (r == 0.0) {
    return finalize(-(v + p * n));
  }
  const double pow_term = std::pow(1.0 + r, n);
  const double result = -(v * pow_term + p * (1.0 + r * t) * (pow_term - 1.0) / r);
  return finalize(result);
}

// --- PMT(rate, nper, pv, [fv=0], [type=0]) ------------------------------
//
//   rate == 0   ->  PMT = -(pv + fv) / nper   (nper == 0 -> #NUM!)
//   rate != 0   ->  f   = (1 + rate)^nper
//                   PMT = -(pv * f + fv) * rate / ((1 + rate*type) * (f - 1))
//
// Zero denominators (nper == 0 when rate == 0, or f == 1 when rate != 0
// with nper == 0) surface as `#NUM!`.
Value Pmt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto rate = read_required_number(args, 0);
  if (!rate) {
    return Value::error(rate.error());
  }
  auto nper = read_required_number(args, 1);
  if (!nper) {
    return Value::error(nper.error());
  }
  auto pv = read_required_number(args, 2);
  if (!pv) {
    return Value::error(pv.error());
  }
  auto fv = read_optional_number(args, arity, 3, 0.0);
  if (!fv) {
    return Value::error(fv.error());
  }
  auto type = read_optional_number(args, arity, 4, 0.0);
  if (!type) {
    return Value::error(type.error());
  }
  const double r = rate.value();
  const double n = nper.value();
  const double v = pv.value();
  const double f = fv.value();
  const double t = normalize_type(type.value());
  if (r == 0.0) {
    if (n == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    return finalize(-(v + f) / n);
  }
  const double pow_term = std::pow(1.0 + r, n);
  const double denom = (1.0 + r * t) * (pow_term - 1.0);
  if (denom == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = -(v * pow_term + f) * r / denom;
  return finalize(result);
}

// --- NPER(rate, pmt, pv, [fv=0], [type=0]) ------------------------------
//
//   rate == 0  ->  pmt == 0 -> #NUM!, otherwise NPER = -(pv + fv) / pmt
//
//   rate != 0  ->  type == 0:  num = pmt - fv * rate
//                               den = pmt + pv * rate
//                  type == 1:  num = pmt * (1 + rate) - fv * rate
//                               den = pmt * (1 + rate) + pv * rate
//                  NPER = log(num / den) / log(1 + rate)
//                  The quotient num/den must be strictly positive (log
//                  domain); a non-positive quotient surfaces as `#NUM!`.
//
// Note: Excel 365 Mac does *not* clamp NPER to the non-negative half line —
// a sign-mismatched loan such as `=NPER(0.05/12, 500, 25000)` (both pmt
// and pv positive) produces the raw negative algebraic answer (~-45.51),
// not `#NUM!`. The oracle confirmed this. We therefore return whatever
// the log formula yields as long as the input domain is valid.
Value Nper(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto rate = read_required_number(args, 0);
  if (!rate) {
    return Value::error(rate.error());
  }
  auto pmt = read_required_number(args, 1);
  if (!pmt) {
    return Value::error(pmt.error());
  }
  auto pv = read_required_number(args, 2);
  if (!pv) {
    return Value::error(pv.error());
  }
  auto fv = read_optional_number(args, arity, 3, 0.0);
  if (!fv) {
    return Value::error(fv.error());
  }
  auto type = read_optional_number(args, arity, 4, 0.0);
  if (!type) {
    return Value::error(type.error());
  }
  const double r = rate.value();
  const double p = pmt.value();
  const double v = pv.value();
  const double f = fv.value();
  const double t = normalize_type(type.value());
  if (r == 0.0) {
    if (p == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    return finalize(-(v + f) / p);
  }
  if (r <= -1.0) {
    // log(1 + rate) is undefined when rate <= -1; Excel returns #NUM!.
    return Value::error(ErrorCode::Num);
  }
  const double pmt_scaled = t == 0.0 ? p : p * (1.0 + r);
  const double num = pmt_scaled - f * r;
  const double den = pmt_scaled + v * r;
  if (den == 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double ratio = num / den;
  if (ratio <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = std::log(ratio) / std::log(1.0 + r);
  return finalize(result);
}

// --- NPV(rate, value1, value2, ...) -------------------------------------
//
// Discounts each value by its 1-based positional index `i` (NPV's rate is
// applied from period 1, unlike IRR which uses period 0..n-1):
//
//   NPV = sum_{i=1..n} value_i / (1 + rate)^i
//
// When `accepts_ranges = true` with `range_filter_numeric_only = true`,
// the dispatcher flattens each RangeOp into consecutive numeric cells in
// row-major order; the positional counter runs across the concatenated
// flat sequence (scalars in argument order, ranges row-major, continuing
// the index). Text / Bool / Blank cells inside a range are filtered out
// by the dispatcher before reaching this impl. Direct scalar arguments
// still coerce through `coerce_to_number`, so a direct bool argument
// becomes 1.0 / 0.0 and a direct blank becomes 0.0 — matching Excel.
Value Npv(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  // arity guaranteed >= 2 by min_arity in register_financial_builtins.
  auto rate_e = read_required_number(args, 0);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  const double rate = rate_e.value();
  double total = 0.0;
  double discount = 1.0 + rate;
  // Walk value1..valueN. NPV skips every non-numeric value regardless of
  // how it arrived: range-sourced bool/text/blank cells were already dropped
  // by the dispatcher's numeric-only filter, and the same rule applies to
  // direct scalar arguments (TRUE, FALSE, blanks, and text literals do not
  // contribute a cash flow and do not advance the period). `period_discount`
  // only steps forward when a numeric cash flow is actually consumed.
  double period_discount = discount;
  for (std::uint32_t i = 1; i < arity; ++i) {
    const Value& v = args[i];
    if (v.is_error()) {
      return v;
    }
    if (!v.is_number()) {
      continue;
    }
    if (period_discount == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    total += v.as_number() / period_discount;
    period_discount *= discount;
  }
  return finalize(total);
}

// --- RATE(nper, pmt, pv, [fv=0], [type=0], [guess=0.1]) -----------------
//
// Solves for the per-period rate r that zeroes the annuity equation:
//
//   f(r) = pv*(1+r)^nper
//        + pmt * (1 + r*type) * ((1+r)^nper - 1) / r
//        + fv
//
// Newton-Raphson iterates up to 100 times to converge to |delta| < 1e-10
// on the rate, matching IRR's tolerance (see `financial_lazy.cpp`).
//
// Special cases:
//   - `nper < 1`                       -> #NUM! (undefined number of periods)
//   - `pv + pmt*nper + fv == 0`        -> r = 0 exactly (zero-rate shortcut)
//   - rate iterate drops below -1      -> #NUM! (log-domain boundary)
//   - iterate falls within 1e-15 of 0  -> #NUM! (derivative blow-up; caller
//                                        can retry with a different guess)
//   - 100 iterations without convergence -> #NUM!
//
// The damped Newton step halves the update when |f/df| exceeds 1.0 to
// keep oscillations on bad initial guesses from walking past the root;
// this matches the heuristic used by most open-source RATE
// implementations and by Excel in practice.
Value Rate(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto nper_e = read_required_number(args, 0);
  if (!nper_e) {
    return Value::error(nper_e.error());
  }
  auto pmt_e = read_required_number(args, 1);
  if (!pmt_e) {
    return Value::error(pmt_e.error());
  }
  auto pv_e = read_required_number(args, 2);
  if (!pv_e) {
    return Value::error(pv_e.error());
  }
  auto fv_e = read_optional_number(args, arity, 3, 0.0);
  if (!fv_e) {
    return Value::error(fv_e.error());
  }
  auto type_e = read_optional_number(args, arity, 4, 0.0);
  if (!type_e) {
    return Value::error(type_e.error());
  }
  auto guess_e = read_optional_number(args, arity, 5, 0.1);
  if (!guess_e) {
    return Value::error(guess_e.error());
  }

  const double nper = nper_e.value();
  const double pmt = pmt_e.value();
  const double pv = pv_e.value();
  const double fv = fv_e.value();
  const double type = normalize_type(type_e.value());
  double rate = guess_e.value();

  if (nper < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  // Zero-rate degenerate case: if the linear residual is zero the answer
  // is r = 0 exactly. Avoids the division-by-r blow-up in NR.
  if (pv + pmt * nper + fv == 0.0) {
    return Value::number(0.0);
  }

  constexpr int kMaxIter = 100;
  constexpr double kTolerance = 1.0e-10;
  constexpr double kRateFloor = 1.0e-15;
  for (int iter = 0; iter < kMaxIter; ++iter) {
    if (rate <= -1.0) {
      return Value::error(ErrorCode::Num);
    }
    if (std::fabs(rate) < kRateFloor) {
      // The analytic f/df limit at r=0 exists, but evaluating the
      // generic form divides by r. Nudging away from 0 by kRateFloor
      // recovers a well-defined derivative; if we're truly at the root
      // the zero-rate shortcut above has already handled it.
      rate = rate < 0.0 ? -kRateFloor : kRateFloor;
    }
    const double pow_term = std::pow(1.0 + rate, nper);
    const double pow_term_m1 = pow_term - 1.0;
    const double f = pv * pow_term + pmt * (1.0 + rate * type) * pow_term_m1 / rate + fv;
    // df/dr = d/dr [pv*(1+r)^n]
    //       + d/dr [pmt*(1+r*type) * ((1+r)^n - 1) / r]
    //       = pv*n*(1+r)^(n-1)
    //       + pmt*type * ((1+r)^n - 1) / r
    //       + pmt*(1+r*type) * (n*r*(1+r)^(n-1) - ((1+r)^n - 1)) / r^2
    const double pow_nm1 = std::pow(1.0 + rate, nper - 1.0);
    const double df = pv * nper * pow_nm1 + pmt * type * pow_term_m1 / rate +
                      pmt * (1.0 + rate * type) * (nper * rate * pow_nm1 - pow_term_m1) / (rate * rate);
    if (df == 0.0 || std::isnan(df) || std::isinf(df) || std::isnan(f) || std::isinf(f)) {
      return Value::error(ErrorCode::Num);
    }
    double step = f / df;
    // Damp large steps — cheap safeguard against overshoot on poor
    // initial guesses. Empirically stabilises the oracle cases without
    // hurting convergence speed on well-conditioned inputs.
    if (std::fabs(step) > 1.0) {
      step *= 0.5;
    }
    const double new_rate = rate - step;
    if (std::isnan(new_rate) || std::isinf(new_rate)) {
      return Value::error(ErrorCode::Num);
    }
    if (std::fabs(new_rate - rate) < kTolerance) {
      return finalize(new_rate);
    }
    rate = new_rate;
  }
  return Value::error(ErrorCode::Num);
}

// --- IPMT(rate, per, nper, pv, [fv=0], [type=0]) ------------------------
//
// Interest portion of the period-`per` payment on a standard annuity
// schedule. `per` is 1-based and must fall within `[1, nper]`; otherwise
// `#NUM!`. See `ipmt_scalar` above for the formula details.
Value Ipmt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto rate_e = read_required_number(args, 0);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  auto per_e = read_required_number(args, 1);
  if (!per_e) {
    return Value::error(per_e.error());
  }
  auto nper_e = read_required_number(args, 2);
  if (!nper_e) {
    return Value::error(nper_e.error());
  }
  auto pv_e = read_required_number(args, 3);
  if (!pv_e) {
    return Value::error(pv_e.error());
  }
  auto fv_e = read_optional_number(args, arity, 4, 0.0);
  if (!fv_e) {
    return Value::error(fv_e.error());
  }
  auto type_e = read_optional_number(args, arity, 5, 0.0);
  if (!type_e) {
    return Value::error(type_e.error());
  }
  const double result = ipmt_scalar(rate_e.value(), per_e.value(), nper_e.value(), pv_e.value(), fv_e.value(),
                                    normalize_type(type_e.value()));
  return finalize(result);
}

// --- PPMT(rate, per, nper, pv, [fv=0], [type=0]) ------------------------
//
// Principal portion of the period-`per` payment = PMT - IPMT. PPMT shares
// IPMT's domain: `per < 1` and integer `per > nper` surface as `#NUM!`,
// but a fractional `per > nper` evaluates the closed form (matches Mac
// Excel 365 / IronCalc oracle). The integer-per > nper check is applied
// inside `ipmt_scalar`, so PPMT itself only needs to reject `per < 1`.
Value Ppmt(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto rate_e = read_required_number(args, 0);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  auto per_e = read_required_number(args, 1);
  if (!per_e) {
    return Value::error(per_e.error());
  }
  auto nper_e = read_required_number(args, 2);
  if (!nper_e) {
    return Value::error(nper_e.error());
  }
  auto pv_e = read_required_number(args, 3);
  if (!pv_e) {
    return Value::error(pv_e.error());
  }
  auto fv_e = read_optional_number(args, arity, 4, 0.0);
  if (!fv_e) {
    return Value::error(fv_e.error());
  }
  auto type_e = read_optional_number(args, arity, 5, 0.0);
  if (!type_e) {
    return Value::error(type_e.error());
  }
  const double rate = rate_e.value();
  const double per = per_e.value();
  const double nper = nper_e.value();
  const double pv = pv_e.value();
  const double fv = fv_e.value();
  const double type = normalize_type(type_e.value());
  if (per < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double pmt = pmt_scalar(rate, nper, pv, fv, type);
  if (std::isnan(pmt)) {
    return Value::error(ErrorCode::Num);
  }
  const double interest = ipmt_scalar(rate, per, nper, pv, fv, type);
  if (std::isnan(interest)) {
    return Value::error(ErrorCode::Num);
  }
  return finalize(pmt - interest);
}

// --- CUMIPMT(rate, nper, pv, start, end, type) --------------------------
//
// Sum of interest paid from period `start` to `end` inclusive. Unlike the
// PMT family, `type` is REQUIRED here (6-arity, not 5+optional) —
// matching Excel's signature.
//
// Domain checks per Microsoft docs:
//   - rate  > 0
//   - nper  > 0
//   - pv    > 0
//   - 1 <= start <= end <= nper
// Violating any check returns `#NUM!`.
Value Cumipmt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto rate_e = read_required_number(args, 0);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  auto nper_e = read_required_number(args, 1);
  if (!nper_e) {
    return Value::error(nper_e.error());
  }
  auto pv_e = read_required_number(args, 2);
  if (!pv_e) {
    return Value::error(pv_e.error());
  }
  auto start_e = read_required_number(args, 3);
  if (!start_e) {
    return Value::error(start_e.error());
  }
  auto end_e = read_required_number(args, 4);
  if (!end_e) {
    return Value::error(end_e.error());
  }
  auto type_e = read_required_number(args, 5);
  if (!type_e) {
    return Value::error(type_e.error());
  }
  const double rate = rate_e.value();
  const double nper = nper_e.value();
  const double pv = pv_e.value();
  const double start = start_e.value();
  const double end = end_e.value();
  // CUMIPMT validates `type` strictly: only the literal values 0 or 1 are
  // accepted. Fractional or other numeric values yield `#NUM!`. This is
  // stricter than the PMT / PV / FV family where any non-zero value folds
  // to 1 through `normalize_type`.
  if (type_e.value() != 0.0 && type_e.value() != 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double type = normalize_type(type_e.value());
  if (rate <= 0.0 || nper <= 0.0 || pv <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (start < 1.0 || end < 1.0 || start > end || start > nper || end > nper) {
    return Value::error(ErrorCode::Num);
  }
  double total = 0.0;
  // Normalise fractional period indices: Mac Excel 365 rounds `start` UP
  // (ceil) and `end` DOWN (floor) before iterating, so a fractional
  // `start_period = 1.2` skips period 1. This diverges from Microsoft's
  // "truncated to integer" documentation but matches the oracle across the
  // CUMIPMT / CUMPRINC fixture rows.
  const auto start_i = static_cast<std::int64_t>(std::ceil(start));
  const auto end_i = static_cast<std::int64_t>(std::floor(end));
  if (start_i > end_i) {
    return Value::error(ErrorCode::Num);
  }
  for (std::int64_t p = start_i; p <= end_i; ++p) {
    const double interest = ipmt_scalar(rate, static_cast<double>(p), nper, pv, 0.0, type);
    if (std::isnan(interest) || std::isinf(interest)) {
      return Value::error(ErrorCode::Num);
    }
    total += interest;
  }
  return finalize(total);
}

// --- CUMPRINC(rate, nper, pv, start, end, type) -------------------------
//
// Sum of principal paid from period `start` to `end` inclusive. Same
// domain contract as CUMIPMT.
Value Cumprinc(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto rate_e = read_required_number(args, 0);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  auto nper_e = read_required_number(args, 1);
  if (!nper_e) {
    return Value::error(nper_e.error());
  }
  auto pv_e = read_required_number(args, 2);
  if (!pv_e) {
    return Value::error(pv_e.error());
  }
  auto start_e = read_required_number(args, 3);
  if (!start_e) {
    return Value::error(start_e.error());
  }
  auto end_e = read_required_number(args, 4);
  if (!end_e) {
    return Value::error(end_e.error());
  }
  auto type_e = read_required_number(args, 5);
  if (!type_e) {
    return Value::error(type_e.error());
  }
  const double rate = rate_e.value();
  const double nper = nper_e.value();
  const double pv = pv_e.value();
  const double start = start_e.value();
  const double end = end_e.value();
  // See CUMIPMT: CUMPRINC mirrors the strict 0-or-1 `type` validation.
  if (type_e.value() != 0.0 && type_e.value() != 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double type = normalize_type(type_e.value());
  if (rate <= 0.0 || nper <= 0.0 || pv <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (start < 1.0 || end < 1.0 || start > end || start > nper || end > nper) {
    return Value::error(ErrorCode::Num);
  }
  const double pmt = pmt_scalar(rate, nper, pv, 0.0, type);
  if (std::isnan(pmt)) {
    return Value::error(ErrorCode::Num);
  }
  double total = 0.0;
  // See CUMIPMT above: ceil(start), floor(end) matches Mac Excel 365's
  // fractional-period behaviour.
  const auto start_i = static_cast<std::int64_t>(std::ceil(start));
  const auto end_i = static_cast<std::int64_t>(std::floor(end));
  if (start_i > end_i) {
    return Value::error(ErrorCode::Num);
  }
  for (std::int64_t p = start_i; p <= end_i; ++p) {
    const double interest = ipmt_scalar(rate, static_cast<double>(p), nper, pv, 0.0, type);
    if (std::isnan(interest) || std::isinf(interest)) {
      return Value::error(ErrorCode::Num);
    }
    total += pmt - interest;
  }
  return finalize(total);
}

}  // namespace

void register_financial_builtins(FunctionRegistry& registry) {
  // PV / FV / PMT / NPER all share the (rate, nper, third, [fourth], [type])
  // arity pattern: 3 required + up to 2 optional = min 3, max 5.
  registry.register_function(FunctionDef{"PV", 3u, 5u, &Pv});
  registry.register_function(FunctionDef{"FV", 3u, 5u, &Fv});
  registry.register_function(FunctionDef{"PMT", 3u, 5u, &Pmt});
  registry.register_function(FunctionDef{"NPER", 3u, 5u, &Nper});

  // NPV: rate + at least one value, unbounded variadic. Range-aware with
  // numeric-only filtering so bool/text/blank cells inside a range are
  // silently skipped (matching Excel's behaviour for SUM / AVERAGE).
  {
    FunctionDef def{"NPV", 2u, kVariadic, &Npv};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }

  // RATE: 3 required (nper, pmt, pv) + up to 3 optional (fv, type, guess).
  registry.register_function(FunctionDef{"RATE", 3u, 6u, &Rate});

  // IPMT / PPMT: (rate, per, nper, pv, [fv], [type]) = min 4, max 6.
  registry.register_function(FunctionDef{"IPMT", 4u, 6u, &Ipmt});
  registry.register_function(FunctionDef{"PPMT", 4u, 6u, &Ppmt});

  // CUMIPMT / CUMPRINC: `type` is REQUIRED (Excel's signature), so both
  // take exactly 6 args — no optional tail.
  registry.register_function(FunctionDef{"CUMIPMT", 6u, 6u, &Cumipmt});
  registry.register_function(FunctionDef{"CUMPRINC", 6u, 6u, &Cumprinc});

  // Depreciation family — all eager, no range support. SLN/SYD have
  // fixed arity; DDB/DB take an optional trailing (factor / month) arg.
  // VDB accepts an optional factor and no_switch tail; AMORDEGRC /
  // AMORLINC accept an optional basis tail. Implementations live in
  // `financial_depreciation.cpp`.
  registry.register_function(FunctionDef{"SLN", 3u, 3u, &financial_detail::Sln});
  registry.register_function(FunctionDef{"SYD", 4u, 4u, &financial_detail::Syd});
  registry.register_function(FunctionDef{"DDB", 4u, 5u, &financial_detail::Ddb});
  registry.register_function(FunctionDef{"DB", 4u, 5u, &financial_detail::Db});
  registry.register_function(FunctionDef{"VDB", 5u, 7u, &financial_detail::Vdb});
  registry.register_function(FunctionDef{"AMORDEGRC", 6u, 7u, &financial_detail::Amordegrc});
  registry.register_function(FunctionDef{"AMORLINC", 6u, 7u, &financial_detail::Amorlinc});

  // Accrued interest family. Implementations live in
  // `financial_accrual.cpp`.
  //   ACCRINT:  6 required + optional (basis, calc_method), max 8.
  //   ACCRINTM: 4 required + optional basis,              max 5.
  registry.register_function(FunctionDef{"ACCRINT", 6u, 8u, &financial_detail::Accrint});
  registry.register_function(FunctionDef{"ACCRINTM", 4u, 5u, &financial_detail::Accrintm});

  // Fractional-dollar quote conversion. Pure scalar pair, no range arg.
  // Implementations live in `financial_misc.cpp`.
  registry.register_function(FunctionDef{"DOLLARDE", 2u, 2u, &financial_detail::DollarDe});
  registry.register_function(FunctionDef{"DOLLARFR", 2u, 2u, &financial_detail::DollarFr});

  // Rate / period conversions. All scalar, fixed arity.
  // Implementations live in `financial_misc.cpp`.
  registry.register_function(FunctionDef{"EFFECT", 2u, 2u, &financial_detail::Effect});
  registry.register_function(FunctionDef{"NOMINAL", 2u, 2u, &financial_detail::Nominal});
  registry.register_function(FunctionDef{"PDURATION", 3u, 3u, &financial_detail::PDuration});
  registry.register_function(FunctionDef{"RRI", 3u, 3u, &financial_detail::Rri});
  registry.register_function(FunctionDef{"ISPMT", 4u, 4u, &financial_detail::IsPmt});

  // FVSCHEDULE: principal + variadic schedule of rates. Range-aware with
  // numeric-only filtering so Bool / Text / Blank cells inside a schedule
  // range are silently skipped (matching Excel's behaviour for range
  // inputs, mirroring NPV). Implementation lives in `financial_misc.cpp`.
  {
    FunctionDef def{"FVSCHEDULE", 2u, kVariadic, &financial_detail::FvSchedule};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }

  // Security-rate / T-Bill family. All eager scalar, no range support.
  // Implementations live in `financial_rates.cpp`.
  //   DISC / INTRATE / RECEIVED: 4 required + optional basis (min 4, max 5).
  //   TBILLPRICE / TBILLYIELD / TBILLEQ: exactly 3 args (fixed basis).
  registry.register_function(FunctionDef{"DISC", 4u, 5u, &financial_detail::Disc});
  registry.register_function(FunctionDef{"INTRATE", 4u, 5u, &financial_detail::Intrate});
  registry.register_function(FunctionDef{"RECEIVED", 4u, 5u, &financial_detail::Received});
  registry.register_function(FunctionDef{"TBILLPRICE", 3u, 3u, &financial_detail::TBillPrice});
  registry.register_function(FunctionDef{"TBILLYIELD", 3u, 3u, &financial_detail::TBillYield});
  registry.register_function(FunctionDef{"TBILLEQ", 3u, 3u, &financial_detail::TBillEq});

  // Closed-form bond pricing / yield helpers. All eager scalar, no range
  // support. Implementations live in `financial_bond_simple.cpp`.
  //   PRICEDISC / YIELDDISC: 4 required + optional basis (min 4, max 5).
  //   PRICEMAT / YIELDMAT:   5 required + optional basis (min 5, max 6).
  registry.register_function(FunctionDef{"PRICEDISC", 4u, 5u, &financial_detail::PriceDisc});
  registry.register_function(FunctionDef{"PRICEMAT", 5u, 6u, &financial_detail::PriceMat});
  registry.register_function(FunctionDef{"YIELDDISC", 4u, 5u, &financial_detail::YieldDisc});
  registry.register_function(FunctionDef{"YIELDMAT", 5u, 6u, &financial_detail::YieldMat});

  // STOCKHISTORY: stub returning #VALUE!. Formulon is a pure calc engine
  // and does not perform network / market-data I/O. Accepts any tail of
  // properties (2..kVariadic). See `financial_bond_simple.cpp` for the
  // rationale; mirrors the WEBSERVICE / PY pattern in `web.cpp`.
  registry.register_function(FunctionDef{"STOCKHISTORY", 2u, kVariadic, &financial_detail::StockHistory});
}

}  // namespace eval
}  // namespace formulon
