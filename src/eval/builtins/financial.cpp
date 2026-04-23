// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the eager financial built-ins: PV, FV, PMT, NPER,
// NPV, RATE, IPMT, PPMT, CUMIPMT, CUMPRINC, and the depreciation family
// SLN / SYD / DDB / DB. Each follows Excel's time-value-of-money sign
// convention (cash out = negative, cash in = positive) and defaults to
// end-of-period payments (`type = 0`). See
// `backup/plans/02-calc-engine.md` for the Excel-compatibility
// reference.
//
// IRR is intentionally absent — it needs the un-flattened AST of its
// values argument to walk range/Ref/ArrayLiteral shapes, so it lives on
// the lazy-dispatch seam in `eval/financial_lazy.cpp`.

#include "eval/builtins/financial.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <limits>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Reads an optional trailing numeric argument at position `index`, falling
// back to `default_value` when `arity <= index`. Non-finite coerced values
// surface as `#NUM!` to match the rest of the math-family argument-handling
// conventions. The default is returned by value rather than by reference so
// callers never need to worry about lifetime of a sentinel.
Expected<double, ErrorCode> read_optional_number(const Value* args, std::uint32_t arity, std::uint32_t index,
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
Expected<double, ErrorCode> read_required_number(const Value* args, std::uint32_t index) {
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
Value finalize(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// Normalises the `type` argument (end-vs-begin of period). Excel accepts
// any numeric value; zero means end-of-period, any non-zero value means
// begin-of-period. We mirror that here so the impl formulas can use `type`
// as a 0-or-1 multiplier without a secondary branch.
double normalize_type(double t) noexcept {
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
double pmt_scalar(double rate, double nper, double pv, double fv, double type) noexcept {
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
double fv_scalar(double rate, double nper, double pmt, double pv, double type) noexcept {
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
double ipmt_scalar(double rate, double per, double nper, double pv, double fv, double type) noexcept {
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
  // Walk value1..valueN. `position` starts at 1 (period 1) and
  // increments for every value that made it through range filtering.
  double period_discount = discount;
  for (std::uint32_t i = 1; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (period_discount == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    total += coerced.value() / period_discount;
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
// Principal portion of the period-`per` payment = PMT - IPMT. Both
// components share the same `per` domain check (per in [1, nper]).
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
  if (per < 1.0 || per > nper) {
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
  const double type = normalize_type(type_e.value());
  if (rate <= 0.0 || nper <= 0.0 || pv <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  if (start < 1.0 || end < 1.0 || start > end || start > nper || end > nper) {
    return Value::error(ErrorCode::Num);
  }
  double total = 0.0;
  // Truncate start/end to integers; Excel treats fractional period
  // indices as their floor in the range-iteration family.
  const auto start_i = static_cast<std::int64_t>(std::floor(start));
  const auto end_i = static_cast<std::int64_t>(std::floor(end));
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
  const auto start_i = static_cast<std::int64_t>(std::floor(start));
  const auto end_i = static_cast<std::int64_t>(std::floor(end));
  for (std::int64_t p = start_i; p <= end_i; ++p) {
    const double interest = ipmt_scalar(rate, static_cast<double>(p), nper, pv, 0.0, type);
    if (std::isnan(interest) || std::isinf(interest)) {
      return Value::error(ErrorCode::Num);
    }
    total += pmt - interest;
  }
  return finalize(total);
}

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
//   - period <  1  or  period > life      ->  #NUM!
//
// `period` and `life` are used as-is (non-integer values are accepted);
// Excel does not floor them before the arithmetic.
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
  if (life <= 0.0 || period < 1.0 || period > life) {
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
//   rate   = round((1 - (salvage/cost)^(1/life)) * 1000) / 1000
//   dep_1  = cost * rate * (month / 12)
//   dep_i  = (cost - total_prior_dep) * rate               (2 <= i <= life)
//   dep_L+1= (cost - total_prior_dep) * rate * (12 - month) / 12
//                                                           (only when
//                                                            month < 12)
//
// Domain:
//   - cost   <= 0    ->  #NUM!
//   - salvage<  0    ->  #NUM!
//   - life  <= 0     ->  #NUM!
//   - period<  1     ->  #NUM!
//   - month <  1  or month > 12  ->  #NUM!
//   - period == life+1 is only valid when month < 12 (the partial first
//     year leaves a partial last year); otherwise #NUM!.
//   - period >  life+1 -> #NUM! regardless of month.
//
// Edge cases:
//   - `salvage == 0`: `pow(0, 1/life) = 0`, so rate = round(1) = 1.000
//     and the entire cost depreciates in the first period (prorated by
//     `month / 12`). Subsequent periods charge nothing.
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
  const double month = month_e.value();
  if (cost <= 0.0 || salvage < 0.0 || life <= 0.0 || period < 1.0 || month < 1.0 || month > 12.0) {
    return Value::error(ErrorCode::Num);
  }
  // Only valid period values: [1, life] for any month; life+1 only when
  // month < 12 (partial-first-year case); anything beyond is #NUM!.
  if (period > life + 1.0) {
    return Value::error(ErrorCode::Num);
  }
  if (period > life && month >= 12.0) {
    return Value::error(ErrorCode::Num);
  }
  // Rate: (1 - (salvage/cost)^(1/life)), rounded to 3 decimals. When
  // salvage == 0, pow(0, 1/life) == 0 -> rate = 1.000 exactly.
  double rate_raw = 1.0 - std::pow(salvage / cost, 1.0 / life);
  const double rate = std::floor(rate_raw * 1000.0 + 0.5) / 1000.0;
  // First period: prorated by (month/12).
  const double dep_1 = cost * rate * month / 12.0;
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
  // period == life + 1 (partial last year). Guarded above so month < 12.
  const double dep_last = (cost - total) * rate * (12.0 - month) / 12.0;
  return finalize(dep_last);
}

// --- DOLLARDE / DOLLARFR shared denominator scaling --------------------
//
// Excel's fractional-dollar pair converts between "integer.fraction" quotes
// (price) and their decimal equivalents using a denominator-aware scale:
//
//   scale = 10 ^ ceil(log10(denom))
//
// The denominator is truncated to an integer first. denom < 0 -> #NUM!;
// denom in [0, 1) -> #DIV/0! (Excel's documented error for a fraction
// with zero denominator). denom == 1 is degenerate but well-defined: any
// integer quote is already "decimal". The caller passes the truncated
// integer denominator; this helper just computes the scale.
Expected<double, ErrorCode> dollar_scale_from_denom(double denom_raw) {
  if (std::isnan(denom_raw) || std::isinf(denom_raw)) {
    return ErrorCode::Num;
  }
  if (denom_raw < 0.0) {
    return ErrorCode::Num;
  }
  const double denom = std::trunc(denom_raw);
  if (denom < 1.0) {
    // denom in [0, 1) after truncation == 0: divide-by-zero domain.
    return ErrorCode::Div0;
  }
  // ceil(log10(denom)) counts the decimal digits of denom. For denom == 1,
  // log10(1) == 0 and we want scale == 1 (no shift).
  const double logd = std::log10(denom);
  const double digits = std::ceil(logd);
  return std::pow(10.0, digits);
}

// --- DOLLARDE(fractional_price, fraction_denom) ------------------------
//
// Fractional-dollar quote -> decimal price. Given `price = int.frac` where
// `frac` is interpreted as a numerator over `denom`:
//
//   result = trunc(price) + frac * scale / denom
//           where scale = 10 ^ ceil(log10(trunc(denom)))
//
// `frac` is the post-decimal portion of `price` (price - trunc(price)).
// Example: DOLLARDE(1.1, 16) treats 1.1 as "1 + 10/16 = 1.625" because
// scale = 100 and frac = 0.1 -> 0.1 * 100 / 16 = 0.625.
Value DollarDe(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto price_e = read_required_number(args, 0);
  if (!price_e) {
    return Value::error(price_e.error());
  }
  auto denom_e = read_required_number(args, 1);
  if (!denom_e) {
    return Value::error(denom_e.error());
  }
  auto scale = dollar_scale_from_denom(denom_e.value());
  if (!scale) {
    return Value::error(scale.error());
  }
  const double price = price_e.value();
  const double denom = std::trunc(denom_e.value());
  const double integer_part = std::trunc(price);
  const double fractional_part = price - integer_part;
  const double decimal = integer_part + fractional_part * scale.value() / denom;
  return finalize(decimal);
}

// --- DOLLARFR(decimal_price, fraction_denom) ---------------------------
//
// Decimal price -> fractional-dollar quote. Inverse of DOLLARDE:
//
//   result = trunc(price) + frac * denom / scale
//           where frac = price - trunc(price) and scale as above.
Value DollarFr(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto price_e = read_required_number(args, 0);
  if (!price_e) {
    return Value::error(price_e.error());
  }
  auto denom_e = read_required_number(args, 1);
  if (!denom_e) {
    return Value::error(denom_e.error());
  }
  auto scale = dollar_scale_from_denom(denom_e.value());
  if (!scale) {
    return Value::error(scale.error());
  }
  const double price = price_e.value();
  const double denom = std::trunc(denom_e.value());
  const double integer_part = std::trunc(price);
  const double fractional_part = price - integer_part;
  const double fractional = integer_part + fractional_part * denom / scale.value();
  return finalize(fractional);
}

// --- EFFECT(nominal_rate, npery) ---------------------------------------
//
// Annual effective rate from a nominal rate compounded `npery` times per
// year:
//
//   EFFECT = (1 + nominal/npery)^npery - 1
//
// Domain per Microsoft docs:
//   - nominal_rate <= 0  ->  #NUM!
//   - npery        <  1  ->  #NUM!  (after truncation to integer)
//
// `npery` is truncated to an integer before the arithmetic.
Value Effect(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto nom_e = read_required_number(args, 0);
  if (!nom_e) {
    return Value::error(nom_e.error());
  }
  auto npery_e = read_required_number(args, 1);
  if (!npery_e) {
    return Value::error(npery_e.error());
  }
  const double nominal = nom_e.value();
  const double npery = std::trunc(npery_e.value());
  if (nominal <= 0.0 || npery < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = std::pow(1.0 + nominal / npery, npery) - 1.0;
  return finalize(result);
}

// --- NOMINAL(effect_rate, npery) ---------------------------------------
//
// Inverse of EFFECT: the nominal annual rate that, compounded `npery`
// times, yields `effect_rate`:
//
//   NOMINAL = npery * ((1 + effect)^(1/npery) - 1)
//
// Same domain constraints as EFFECT.
Value Nominal(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto eff_e = read_required_number(args, 0);
  if (!eff_e) {
    return Value::error(eff_e.error());
  }
  auto npery_e = read_required_number(args, 1);
  if (!npery_e) {
    return Value::error(npery_e.error());
  }
  const double effect = eff_e.value();
  const double npery = std::trunc(npery_e.value());
  if (effect <= 0.0 || npery < 1.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = npery * (std::pow(1.0 + effect, 1.0 / npery) - 1.0);
  return finalize(result);
}

// --- FVSCHEDULE(principal, schedule) -----------------------------------
//
// Future value of `principal` after applying a sequence of per-period
// rates drawn from `schedule`:
//
//   FV = principal * Π (1 + r_i)     for each r_i in schedule (row-major).
//
// `schedule` is a variadic range-aware tail. When a RangeOp argument is
// supplied, the dispatcher flattens it with `range_filter_numeric_only`,
// so Text / Bool / Blank cells inside the rectangle are dropped silently
// (matching Excel's behaviour for FVSCHEDULE range inputs). Direct scalar
// arguments coerce through `coerce_to_number` - a direct bool becomes
// 1.0/0.0 and a direct blank becomes 0.0 (i.e. applies a 1+0 factor).
//
// An empty schedule (no rates at all) is guarded by the registry's
// min_arity=2, so here we always process at least one rate.
Value FvSchedule(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto principal_e = read_required_number(args, 0);
  if (!principal_e) {
    return Value::error(principal_e.error());
  }
  double result = principal_e.value();
  for (std::uint32_t i = 1; i < arity; ++i) {
    auto rate = coerce_to_number(args[i]);
    if (!rate) {
      return Value::error(rate.error());
    }
    result *= 1.0 + rate.value();
    if (std::isnan(result) || std::isinf(result)) {
      return Value::error(ErrorCode::Num);
    }
  }
  return finalize(result);
}

// --- PDURATION(rate, pv, fv) -------------------------------------------
//
// Number of periods required for an investment at `rate` to grow from `pv`
// to `fv`:
//
//   PDURATION = log(fv / pv) / log(1 + rate)
//
// Domain per Microsoft docs:
//   - rate <= 0, pv <= 0, fv <= 0  ->  #NUM!
Value PDuration(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto rate_e = read_required_number(args, 0);
  if (!rate_e) {
    return Value::error(rate_e.error());
  }
  auto pv_e = read_required_number(args, 1);
  if (!pv_e) {
    return Value::error(pv_e.error());
  }
  auto fv_e = read_required_number(args, 2);
  if (!fv_e) {
    return Value::error(fv_e.error());
  }
  const double rate = rate_e.value();
  const double pv = pv_e.value();
  const double fv = fv_e.value();
  if (rate <= 0.0 || pv <= 0.0 || fv <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = std::log(fv / pv) / std::log(1.0 + rate);
  return finalize(result);
}

// --- RRI(nper, pv, fv) -------------------------------------------------
//
// Equivalent interest rate for the growth of `pv` -> `fv` over `nper`
// periods:
//
//   RRI = (fv / pv)^(1 / nper) - 1
//
// Domain per Microsoft docs:
//   - nper <= 0, pv <= 0, fv <= 0  ->  #NUM!
Value Rri(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto nper_e = read_required_number(args, 0);
  if (!nper_e) {
    return Value::error(nper_e.error());
  }
  auto pv_e = read_required_number(args, 1);
  if (!pv_e) {
    return Value::error(pv_e.error());
  }
  auto fv_e = read_required_number(args, 2);
  if (!fv_e) {
    return Value::error(fv_e.error());
  }
  const double nper = nper_e.value();
  const double pv = pv_e.value();
  const double fv = fv_e.value();
  if (nper <= 0.0 || pv <= 0.0 || fv <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double result = std::pow(fv / pv, 1.0 / nper) - 1.0;
  return finalize(result);
}

// --- ISPMT(rate, per, nper, pv) ----------------------------------------
//
// Lotus 1-2-3 simple-interest schedule (not an amortising loan): interest
// for period `per` when the principal is paid down linearly across
// `nper` periods.
//
//   ISPMT = -pv * rate * (1 - per / nper)
//
// Domain:
//   - nper == 0  ->  #DIV/0!
//
// Unlike the IPMT / PPMT family ISPMT imposes no per-domain check; `per`
// is used as-is (Excel lets negative or out-of-range `per` produce the
// raw algebraic result).
Value IsPmt(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
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
  const double rate = rate_e.value();
  const double per = per_e.value();
  const double nper = nper_e.value();
  const double pv = pv_e.value();
  if (nper == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double result = -pv * rate * (1.0 - per / nper);
  return finalize(result);
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
  registry.register_function(FunctionDef{"SLN", 3u, 3u, &Sln});
  registry.register_function(FunctionDef{"SYD", 4u, 4u, &Syd});
  registry.register_function(FunctionDef{"DDB", 4u, 5u, &Ddb});
  registry.register_function(FunctionDef{"DB", 4u, 5u, &Db});

  // Fractional-dollar quote conversion. Pure scalar pair, no range arg.
  registry.register_function(FunctionDef{"DOLLARDE", 2u, 2u, &DollarDe});
  registry.register_function(FunctionDef{"DOLLARFR", 2u, 2u, &DollarFr});

  // Rate / period conversions. All scalar, fixed arity.
  registry.register_function(FunctionDef{"EFFECT", 2u, 2u, &Effect});
  registry.register_function(FunctionDef{"NOMINAL", 2u, 2u, &Nominal});
  registry.register_function(FunctionDef{"PDURATION", 3u, 3u, &PDuration});
  registry.register_function(FunctionDef{"RRI", 3u, 3u, &Rri});
  registry.register_function(FunctionDef{"ISPMT", 4u, 4u, &IsPmt});

  // FVSCHEDULE: principal + variadic schedule of rates. Range-aware with
  // numeric-only filtering so Bool / Text / Blank cells inside a schedule
  // range are silently skipped (matching Excel's behaviour for range
  // inputs, mirroring NPV).
  {
    FunctionDef def{"FVSCHEDULE", 2u, kVariadic, &FvSchedule};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
}

}  // namespace eval
}  // namespace formulon
