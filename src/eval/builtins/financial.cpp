// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
// end-of-period payments (`type = 0`).
//
// IRR is intentionally absent — it needs the un-flattened AST of its
// values argument to walk range/Ref/ArrayLiteral shapes, so it lives on
// the lazy-dispatch seam in `eval/financial_lazy.cpp`.

#include "eval/builtins/financial.h"

#include <cmath>
#include <cstdint>
#include <utility>

#include "eval/builtins/financial_bond_simple.h"
#include "eval/builtins/financial_duration.h"
#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddfprice.h"
#include "eval/builtins/financial_oddfyield.h"
#include "eval/builtins/financial_oddlprice.h"
#include "eval/builtins/financial_oddlyield.h"
#include "eval/builtins/financial_price.h"
#include "eval/builtins/financial_yield.h"
#include "eval/builtins/registration_helpers.h"
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

struct TvmArgs {
  double first;
  double second;
  double third;
  double fourth;
  double type;
};

struct RateArgs {
  double nper;
  double pmt;
  double pv;
  double fv;
  double type;
  double guess;
};

struct PaymentArgs {
  double rate;
  double per;
  double nper;
  double pv;
  double fv;
  double type;
};

struct CumPaymentArgs {
  double rate;
  double nper;
  double pv;
  double start;
  double end;
  double type;
};

Expected<TvmArgs, ErrorCode> read_tvm_args(const Value* args, std::uint32_t arity, double fourth_default) {
  auto first = read_required_number(args, 0);
  if (!first) {
    return first.error();
  }
  auto second = read_required_number(args, 1);
  if (!second) {
    return second.error();
  }
  auto third = read_required_number(args, 2);
  if (!third) {
    return third.error();
  }
  auto fourth = read_optional_number(args, arity, 3, fourth_default);
  if (!fourth) {
    return fourth.error();
  }
  auto type = read_optional_number(args, arity, 4, 0.0);
  if (!type) {
    return type.error();
  }
  return TvmArgs{first.value(), second.value(), third.value(), fourth.value(), normalize_type(type.value())};
}

Expected<RateArgs, ErrorCode> read_rate_args(const Value* args, std::uint32_t arity) {
  auto tvm = read_tvm_args(args, arity, 0.0);
  if (!tvm) {
    return tvm.error();
  }
  auto guess = read_optional_number(args, arity, 5, 0.1);
  if (!guess) {
    return guess.error();
  }
  return RateArgs{tvm.value().first,  tvm.value().second, tvm.value().third,
                  tvm.value().fourth, tvm.value().type,   guess.value()};
}

Expected<PaymentArgs, ErrorCode> read_payment_args(const Value* args, std::uint32_t arity) {
  auto rate = read_required_number(args, 0);
  if (!rate) {
    return rate.error();
  }
  auto per = read_required_number(args, 1);
  if (!per) {
    return per.error();
  }
  auto nper = read_required_number(args, 2);
  if (!nper) {
    return nper.error();
  }
  auto pv = read_required_number(args, 3);
  if (!pv) {
    return pv.error();
  }
  auto fv = read_optional_number(args, arity, 4, 0.0);
  if (!fv) {
    return fv.error();
  }
  auto type = read_optional_number(args, arity, 5, 0.0);
  if (!type) {
    return type.error();
  }
  return PaymentArgs{rate.value(), per.value(), nper.value(), pv.value(), fv.value(), normalize_type(type.value())};
}

Expected<CumPaymentArgs, ErrorCode> read_cum_payment_args(const Value* args) {
  for (std::uint32_t i = 0; i < 6; ++i) {
    if (args[i].kind() == ValueKind::Bool) {
      return ErrorCode::Value;
    }
  }
  auto rate = read_required_number(args, 0);
  if (!rate) {
    return rate.error();
  }
  auto nper = read_required_number(args, 1);
  if (!nper) {
    return nper.error();
  }
  auto pv = read_required_number(args, 2);
  if (!pv) {
    return pv.error();
  }
  auto start = read_required_number(args, 3);
  if (!start) {
    return start.error();
  }
  auto end = read_required_number(args, 4);
  if (!end) {
    return end.error();
  }
  auto type = read_required_number(args, 5);
  if (!type) {
    return type.error();
  }
  if (type.value() != 0.0 && type.value() != 1.0) {
    return ErrorCode::Num;
  }
  return CumPaymentArgs{rate.value(),  nper.value(), pv.value(),
                        start.value(), end.value(),  normalize_type(type.value())};
}

Expected<std::pair<std::int64_t, std::int64_t>, ErrorCode> cum_period_bounds(double start, double end, double nper) {
  if (start < 1.0 || end < 1.0 || start > end || start > nper || end > nper) {
    return ErrorCode::Num;
  }
  // Mac Excel 365 rounds `start` UP (ceil) and `end` DOWN (floor) before
  // iterating; see CUMIPMT / CUMPRINC oracle fixtures.
  const auto start_i = static_cast<std::int64_t>(std::ceil(start));
  const auto end_i = static_cast<std::int64_t>(std::floor(end));
  if (start_i > end_i) {
    return ErrorCode::Num;
  }
  return std::pair<std::int64_t, std::int64_t>{start_i, end_i};
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
  auto input = read_tvm_args(args, arity, 0.0);
  if (!input) {
    return Value::error(input.error());
  }
  const double r = input.value().first;
  const double n = input.value().second;
  const double p = input.value().third;
  const double f = input.value().fourth;
  const double t = input.value().type;
  if (r == -1.0) {
    // (1+r)^n is 0 for n>0 (yields division by 0 inside the closed form)
    // or infinite for n<=0; Excel collapses both to #DIV/0!.
    return Value::error(ErrorCode::Div0);
  }
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
  auto input = read_tvm_args(args, arity, 0.0);
  if (!input) {
    return Value::error(input.error());
  }
  const double r = input.value().first;
  const double n = input.value().second;
  const double p = input.value().third;
  const double v = input.value().fourth;
  const double t = input.value().type;
  if (r == -1.0) {
    // (1+r)^n is 0 for n>0 (zero payment discount factor) or infinite for
    // n<=0; Excel collapses both to #DIV/0!.
    return Value::error(ErrorCode::Div0);
  }
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
  auto input = read_tvm_args(args, arity, 0.0);
  if (!input) {
    return Value::error(input.error());
  }
  const double r = input.value().first;
  const double n = input.value().second;
  const double v = input.value().third;
  const double f = input.value().fourth;
  const double t = input.value().type;
  if (r <= -1.0) {
    // Mac Excel 365 rejects rate <= -1 for PMT outright with #NUM!, even
    // when the closed form would evaluate to a finite value (e.g. rate=-3
    // with integer nper). This differs from FV/PV which do compute at
    // rate < -1 when nper is an integer (verified via IronCalc xlsx
    // oracle).
    return Value::error(ErrorCode::Num);
  }
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
  auto input = read_tvm_args(args, arity, 0.0);
  if (!input) {
    return Value::error(input.error());
  }
  const double r = input.value().first;
  const double p = input.value().second;
  const double v = input.value().third;
  const double f = input.value().fourth;
  const double t = input.value().type;
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
  // Walk value1..valueN. Range-sourced bool/text/blank cells were already
  // dropped by the dispatcher's numeric-only filter; a direct scalar
  // logical / numeric-text argument is coerced and contributes at its
  // period (see the per-iteration comment). `period_discount` only steps
  // forward when a cash flow is actually consumed.
  double period_discount = discount;
  for (std::uint32_t i = 1; i < arity; ++i) {
    const Value& v = args[i];
    if (v.is_error()) {
      return v;
    }
    // The dispatcher already dropped range-sourced Bool / Text / Blank
    // cells (range_filter_numeric_only), so any non-Number reaching this
    // impl is a DIRECT scalar argument. Excel counts a directly-passed
    // logical / numeric-text argument (TRUE -> 1, "5" -> 5) at its period;
    // only non-numeric text is ignored. Coerce here so the period counter
    // advances for those, matching Excel.
    double cash = 0.0;
    if (v.is_number()) {
      cash = v.as_number();
    } else {
      auto coerced = coerce_to_number(v);
      if (!coerced) {
        continue;  // non-numeric text: ignored, period unchanged.
      }
      cash = coerced.value();
    }
    if (period_discount == 0.0) {
      return Value::error(ErrorCode::Num);
    }
    total += cash / period_discount;
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
  auto input = read_rate_args(args, arity);
  if (!input) {
    return Value::error(input.error());
  }
  const double nper = input.value().nper;
  const double pmt = input.value().pmt;
  const double pv = input.value().pv;
  const double fv = input.value().fv;
  const double type = input.value().type;
  double rate = input.value().guess;

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
  // Residual gate on the accepted root, mirroring IRR/XIRR: step
  // convergence alone can settle the iterate at a point where the TVM
  // residual is still large, so a converged rate is verified against the
  // equation before it is published.
  constexpr double kResidualTolerance = 1.0e-6;
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
    // Damped Newton step: halve the update when |f/df| exceeds 1.0. A raw
    // Newton step can overshoot badly from a poor initial guess (the TVM
    // function is steep near rate = -1 and near a high-nper root), letting
    // the iterate jump the wrong side of a singularity and either diverge
    // or settle on a different root than Excel reports. Halving the large
    // first steps keeps the iterate in the basin of the principal root
    // without slowing convergence on well-conditioned inputs. The damping
    // factor is heuristic; it is monitored against the oracle corpus.
    if (std::fabs(step) > 1.0) {
      step *= 0.5;
    }
    const double new_rate = rate - step;
    if (std::isnan(new_rate) || std::isinf(new_rate)) {
      return Value::error(ErrorCode::Num);
    }
    if (std::fabs(new_rate - rate) < kTolerance) {
      // Degenerate fixed point at rate == -1 (or dangerously close): the
      // TVM relation (1+r)^n is zero there, so the algebra that produced
      // this convergence is ill-defined. Excel returns #NUM! instead of
      // publishing a spurious -1.0 root. Without this clamp we happily
      // "converged" on `RATE(10, -1200, 2000, 0, 1)` where the iterate
      // slides toward -1 and never crosses out.
      if (new_rate <= -1.0 + kTolerance) {
        return Value::error(ErrorCode::Num);
      }
      // Residual gate: a converged step must also zero the TVM equation.
      // The damping above can stall the iterate at a non-root on
      // pathological inputs; reject those as non-convergent rather than
      // publishing a spurious rate. The equation's terms scale with
      // pv * (1+r)^nper, so the residual is compared relative to that
      // magnitude (an absolute bound would be unreachable for large pv).
      const double pow_check = std::pow(1.0 + new_rate, nper);
      const double annuity = pmt * (1.0 + new_rate * type) * (pow_check - 1.0) / new_rate;
      const double residual = pv * pow_check + annuity + fv;
      const double scale = std::fabs(pv * pow_check) + std::fabs(annuity) + std::fabs(fv) + 1.0;
      if (!std::isfinite(residual) || std::fabs(residual) > kResidualTolerance * scale) {
        return Value::error(ErrorCode::Num);
      }
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
  auto input = read_payment_args(args, arity);
  if (!input) {
    return Value::error(input.error());
  }
  const double result = ipmt_scalar(input.value().rate, input.value().per, input.value().nper, input.value().pv,
                                    input.value().fv, input.value().type);
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
  auto input = read_payment_args(args, arity);
  if (!input) {
    return Value::error(input.error());
  }
  const double rate = input.value().rate;
  const double per = input.value().per;
  const double nper = input.value().nper;
  const double pv = input.value().pv;
  const double fv = input.value().fv;
  const double type = input.value().type;
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
  auto input = read_cum_payment_args(args);
  if (!input) {
    return Value::error(input.error());
  }
  const double rate = input.value().rate;
  const double nper = input.value().nper;
  const double pv = input.value().pv;
  if (rate <= 0.0 || nper <= 0.0 || pv <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto bounds = cum_period_bounds(input.value().start, input.value().end, nper);
  if (!bounds) {
    return Value::error(bounds.error());
  }
  const std::int64_t start_i = bounds.value().first;
  const std::int64_t end_i = bounds.value().second;
  double total = 0.0;
  for (std::int64_t p = start_i; p <= end_i; ++p) {
    const double interest = ipmt_scalar(rate, static_cast<double>(p), nper, pv, 0.0, input.value().type);
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
  auto input = read_cum_payment_args(args);
  if (!input) {
    return Value::error(input.error());
  }
  const double rate = input.value().rate;
  const double nper = input.value().nper;
  const double pv = input.value().pv;
  if (rate <= 0.0 || nper <= 0.0 || pv <= 0.0) {
    return Value::error(ErrorCode::Num);
  }
  auto bounds = cum_period_bounds(input.value().start, input.value().end, nper);
  if (!bounds) {
    return Value::error(bounds.error());
  }
  const std::int64_t start_i = bounds.value().first;
  const std::int64_t end_i = bounds.value().second;
  const double pmt = pmt_scalar(rate, nper, pv, 0.0, input.value().type);
  if (std::isnan(pmt)) {
    return Value::error(ErrorCode::Num);
  }
  double total = 0.0;
  for (std::int64_t p = start_i; p <= end_i; ++p) {
    const double interest = ipmt_scalar(rate, static_cast<double>(p), nper, pv, 0.0, input.value().type);
    if (std::isnan(interest) || std::isinf(interest)) {
      return Value::error(ErrorCode::Num);
    }
    total += pmt - interest;
  }
  return finalize(total);
}

}  // namespace

void register_financial_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"PV", 3u, 5u, &Pv},
      {"FV", 3u, 5u, &Fv},
      {"PMT", 3u, 5u, &Pmt},
      {"NPER", 3u, 5u, &Nper},
      {"NPV", 2u, kVariadic, &Npv, true, true, true},
      {"RATE", 3u, 6u, &Rate},
      {"IPMT", 4u, 6u, &Ipmt},
      {"PPMT", 4u, 6u, &Ppmt},
      {"CUMIPMT", 6u, 6u, &Cumipmt},
      {"CUMPRINC", 6u, 6u, &Cumprinc},
      {"SLN", 3u, 3u, &financial_detail::Sln},
      {"SYD", 4u, 4u, &financial_detail::Syd},
      {"DDB", 4u, 5u, &financial_detail::Ddb},
      {"DB", 4u, 5u, &financial_detail::Db},
      {"VDB", 5u, 7u, &financial_detail::Vdb},
      {"AMORDEGRC", 6u, 7u, &financial_detail::Amordegrc},
      {"AMORLINC", 6u, 7u, &financial_detail::Amorlinc},
      {"ACCRINT", 6u, 8u, &financial_detail::Accrint},
      {"ACCRINTM", 4u, 5u, &financial_detail::Accrintm},
      {"DOLLARDE", 2u, 2u, &financial_detail::DollarDe},
      {"DOLLARFR", 2u, 2u, &financial_detail::DollarFr},
      {"EFFECT", 2u, 2u, &financial_detail::Effect},
      {"NOMINAL", 2u, 2u, &financial_detail::Nominal},
      {"PDURATION", 3u, 3u, &financial_detail::PDuration},
      {"RRI", 3u, 3u, &financial_detail::Rri},
      {"ISPMT", 4u, 4u, &financial_detail::IsPmt},
      {"FVSCHEDULE", 2u, kVariadic, &financial_detail::FvSchedule, true, true, true},
      {"DISC", 4u, 5u, &financial_detail::Disc},
      {"INTRATE", 4u, 5u, &financial_detail::Intrate},
      {"RECEIVED", 4u, 5u, &financial_detail::Received},
      {"TBILLPRICE", 3u, 3u, &financial_detail::TBillPrice},
      {"TBILLYIELD", 3u, 3u, &financial_detail::TBillYield},
      {"TBILLEQ", 3u, 3u, &financial_detail::TBillEq},
      {"PRICEDISC", 4u, 5u, &financial_detail::PriceDisc},
      {"PRICEMAT", 5u, 6u, &financial_detail::PriceMat},
      {"YIELDDISC", 4u, 5u, &financial_detail::YieldDisc},
      {"YIELDMAT", 5u, 6u, &financial_detail::YieldMat},
      {"DURATION", 5u, 6u, &financial_detail::Duration},
      {"MDURATION", 5u, 6u, &financial_detail::MDuration},
      {"PRICE", 6u, 7u, &financial_detail::Price},
      {"YIELD", 6u, 7u, &financial_detail::Yield},
      {"ODDLPRICE", 7u, 8u, &financial_detail::OddlPrice},
      {"ODDLYIELD", 7u, 8u, &financial_detail::OddlYield},
      {"ODDFPRICE", 8u, 9u, &financial_detail::OddfPrice},
      {"ODDFYIELD", 8u, 9u, &financial_detail::OddfYield},
      {"STOCKHISTORY", 2u, kVariadic, &financial_detail::StockHistory},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
