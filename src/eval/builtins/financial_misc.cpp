// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the remaining eager financial built-ins:
// DOLLARDE / DOLLARFR, EFFECT / NOMINAL, FVSCHEDULE, PDURATION, RRI, ISPMT.
// Registered from `financial.cpp` via `register_financial_builtins`.

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

}  // namespace

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

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
