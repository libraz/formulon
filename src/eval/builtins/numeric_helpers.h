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

/// Snaps a quotient to its nearest integer only when it is within a few ULPs.
/// This preserves decimal exact multiples such as `7.1 / 0.1`, which binary
/// floating point can represent just below 71, without collapsing genuinely
/// non-integer quotients near zero.
inline double snap_to_integer(double quotient) noexcept {
  const double nearest = std::round(quotient);
  if (std::fabs(quotient - nearest) < 2e-15 * std::fabs(quotient)) {
    return nearest;
  }
  return quotient;
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

/// Reads a required numeric argument from `args[index]`; a non-finite value
/// surfaces as `#NUM!` from `coerce_to_number`.
inline Expected<double, ErrorCode> read_required_number(const Value* args, std::uint32_t index) {
  return coerce_to_number(args[index]);
}

/// Reads an optional trailing numeric argument at position `index`,
/// returning `default_value` when `arity <= index`. Otherwise behaves
/// like `read_required_number`. The default is returned by value so
/// callers never have to worry about the lifetime of a sentinel
/// reference.
inline Expected<double, ErrorCode> read_optional_number(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                        double default_value) {
  if (arity <= index) {
    return default_value;
  }
  return read_required_number(args, index);
}

/// Reads two required numeric arguments and returns them as a
/// `NumberPair`. Propagates the left-most coercion / non-finite error.
inline Expected<NumberPair, ErrorCode> read_number_pair(const Value* args, std::uint32_t first_index,
                                                        std::uint32_t second_index) {
  auto first = read_required_number(args, first_index);
  if (!first) {
    return std::move(first.error());
  }
  auto second = read_required_number(args, second_index);
  if (!second) {
    return std::move(second.error());
  }
  return NumberPair{first.value(), second.value()};
}

/// Reads three required numeric arguments and returns them as a
/// `NumberTriple`. Propagates the left-most coercion / non-finite
/// error.
inline Expected<NumberTriple, ErrorCode> read_number_triple(const Value* args, std::uint32_t first_index,
                                                            std::uint32_t second_index, std::uint32_t third_index) {
  auto first_second = read_number_pair(args, first_index, second_index);
  if (!first_second) {
    return std::move(first_second.error());
  }
  auto third = read_required_number(args, third_index);
  if (!third) {
    return std::move(third.error());
  }
  return NumberTriple{first_second.value().first, first_second.value().second, third.value()};
}

}  // namespace builtins_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_NUMERIC_HELPERS_H_
