// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared scalar-numeric helpers used across the built-in catalog. These
// were previously duplicated (4x for `kPi`, 5x for the `non-finite ->
// #NUM!` wrap, and 3x for `read_(required|optional)_number` /
// `NumberPair` / `NumberTriple`); centralising them keeps the math /
// stats / financial / distributions / complex / dynamic-array TUs in
// lock-step on coercion semantics and lets the compiler share one
// definition across translation units.
//
// All helpers are inline so the header is self-contained; no
// numeric_helpers.cpp is needed.

#ifndef FORMULON_EVAL_BUILTINS_NUMERIC_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_NUMERIC_HELPERS_H_

#include <cmath>
#include <cstdint>

#include "eval/coerce.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace builtins_detail {

/// Mathematical constant pi to ~15 significant digits. Matches the value
/// `std::acos(-1.0)` produces on any IEEE-754 system, which keeps
/// `RADIANS(180) == kPi` exact and the normal-PDF normalisation
/// (`1 / sqrt(2 * kPi)`) byte-for-byte identical across math, stats,
/// and distributions TUs.
inline constexpr double kPi = 3.14159265358979323846;

/// Wraps a scalar result in a `Value::number`, surfacing `#NUM!` for
/// non-finite (`NaN` / `Inf`) values. This is the universal "finalise a
/// computed double" convention used by EXP, the trig family, the
/// stats descriptive functions, the probability distributions, and the
/// time-value-of-money financials. Centralising it ensures every
/// numeric impl reports overflow / domain-violation via the same code
/// path.
inline Value to_finite_value(double r) noexcept {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

/// (a, b) pair returned by `read_number_pair`. Used by distribution
/// builtins (chi-squared / F-dist / fisher / phi) and the financial
/// TVM family for compact two-argument argument extraction.
struct NumberPair {
  double first;
  double second;
};

/// (a, b, c) triple returned by `read_number_triple`. Used by the
/// three-argument distribution builtins (normal / beta / gamma /
/// weibull / lognormal) and several financial / stats helpers.
struct NumberTriple {
  double first;
  double second;
  double third;
};

/// Reads a required numeric argument from `args[index]`. When
/// `check_finite` is `true` (the default), a `NaN` / `Inf` coerced
/// value surfaces as `#NUM!` so impl bodies never have to re-check.
/// Callers that want to handle non-finite values themselves (e.g. the
/// `stats_detail::read_number_arg` path) pass `check_finite = false`.
inline Expected<double, ErrorCode> read_required_number(const Value* args, std::uint32_t index,
                                                        bool check_finite = true) {
  auto coerced = coerce_to_number(args[index]);
  if (!coerced) {
    return coerced.error();
  }
  const double v = coerced.value();
  if (check_finite && (std::isnan(v) || std::isinf(v))) {
    return ErrorCode::Num;
  }
  return v;
}

/// Reads an optional trailing numeric argument at position `index`,
/// returning `default_value` when `arity <= index`. Otherwise behaves
/// like `read_required_number`. The default is returned by value so
/// callers never have to worry about the lifetime of a sentinel
/// reference.
inline Expected<double, ErrorCode> read_optional_number(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                        double default_value, bool check_finite = true) {
  if (arity <= index) {
    return default_value;
  }
  return read_required_number(args, index, check_finite);
}

/// Reads two required numeric arguments and returns them as a
/// `NumberPair`. Propagates the left-most coercion / non-finite error.
inline Expected<NumberPair, ErrorCode> read_number_pair(const Value* args, std::uint32_t first_index,
                                                        std::uint32_t second_index, bool check_finite = true) {
  auto first = read_required_number(args, first_index, check_finite);
  if (!first) {
    return first.error();
  }
  auto second = read_required_number(args, second_index, check_finite);
  if (!second) {
    return second.error();
  }
  return NumberPair{first.value(), second.value()};
}

/// Reads three required numeric arguments and returns them as a
/// `NumberTriple`. Propagates the left-most coercion / non-finite
/// error.
inline Expected<NumberTriple, ErrorCode> read_number_triple(const Value* args, std::uint32_t first_index,
                                                            std::uint32_t second_index, std::uint32_t third_index,
                                                            bool check_finite = true) {
  auto first = read_required_number(args, first_index, check_finite);
  if (!first) {
    return first.error();
  }
  auto second = read_required_number(args, second_index, check_finite);
  if (!second) {
    return second.error();
  }
  auto third = read_required_number(args, third_index, check_finite);
  if (!third) {
    return third.error();
  }
  return NumberTriple{first.value(), second.value(), third.value()};
}

}  // namespace builtins_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_NUMERIC_HELPERS_H_
