// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's aggregate built-in functions:
// SUM, MIN, MAX, AVERAGE, PRODUCT, COUNT, COUNTA, COUNTBLANK, CONCAT,
// CONCATENATE, and LEN. Each impl follows the same recipe as the rest of
// the builtin catalog: coerce arguments via `eval/coerce.h`, propagate the
// left-most coercion error, and return a `Value`.

#include "eval/builtins_aggregate.h"

#include <cmath>
#include <cstdint>
#include <string>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "eval/utf8_length.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// SUM(value, ...) --------------------------------------------------------
// Excel's SUM coerces each argument to a number; non-coercible text yields
// #VALUE! and any error among the inputs propagates left-to-right.
Value Sum(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    total += coerced.value();
  }
  if (std::isnan(total) || std::isinf(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

// CONCAT(value, ...) / CONCATENATE(value, ...) ---------------------------
// Both spellings share an implementation. Each argument is rendered via
// `coerce_to_text`; left-most error wins. The joined result is interned in
// the call's arena so the returned Value remains readable for the caller.
Value Concat(const Value* args, std::uint32_t arity, Arena& arena) {
  std::string joined;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_text(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    joined.append(coerced.value());
  }
  const std::string_view interned = arena.intern(joined);
  return Value::text(interned);
}

// LEN(text) --------------------------------------------------------------
// Excel reports length in UTF-16 code units, which differs from byte length
// for any non-ASCII codepoint. We coerce the argument to text, then count
// units via the standalone helper.
Value Len(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_text(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  return Value::number(static_cast<double>(utf16_units_in(coerced.value())));
}

// --- Aggregates ---------------------------------------------------------

// MIN(value, ...) - smallest of the coerced numbers. The Excel "skip text
// in cell-references" rule does NOT apply here: a literal non-numeric
// argument coerces via `coerce_to_number` and surfaces `#VALUE!` on
// failure. The caller's pre-evaluation has already short-circuited any
// argument that was itself an error.
Value Min(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  // arity >= 1 by registry contract (min_arity = 1).
  auto first = coerce_to_number(args[0]);
  if (!first) {
    return Value::error(first.error());
  }
  double best = first.value();
  for (std::uint32_t i = 1; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value() < best) {
      best = coerced.value();
    }
  }
  if (std::isnan(best) || std::isinf(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

// MAX(value, ...) - symmetric to MIN.
Value Max(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto first = coerce_to_number(args[0]);
  if (!first) {
    return Value::error(first.error());
  }
  double best = first.value();
  for (std::uint32_t i = 1; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value() > best) {
      best = coerced.value();
    }
  }
  if (std::isnan(best) || std::isinf(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

// AVERAGE(value, ...) - arithmetic mean. With `min_arity = 1` enforced
// at the registry, the divisor is always at least 1, so there is no
// divide-by-zero edge case.
Value Average(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    total += coerced.value();
  }
  const double r = total / static_cast<double>(arity);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// PRODUCT(value, ...) - product of all args. Overflow to Inf -> `#NUM!`.
Value Product(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 1.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    total *= coerced.value();
  }
  if (std::isnan(total) || std::isinf(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

// --- Counting aggregators -----------------------------------------------
//
// COUNT / COUNTA / COUNTBLANK. Unlike SUM / AVERAGE / MIN / MAX / PRODUCT
// these three are registered with `propagate_errors = false`: Excel's COUNT
// family is specified in terms of which cells to "count" rather than which
// values to coerce, so an error inside a range must not short-circuit the
// whole call. That opt-out means the impls see Error-typed values in their
// args array directly and must skip them explicitly.
//
// Accepted divergence (range-vs-direct parity):
//   =COUNT(1, 1/0) in Excel is #DIV/0! because the error propagates as a
//   direct argument; =COUNT(A1:A2) where one cell holds #DIV/0! silently
//   skips the error and counts only the numerics. We do not distinguish
//   the two call shapes - any error anywhere is skipped - so direct-arg
//   callers get the range-shape behaviour. The same simplification already
//   applies to text / bool values appearing inside SUM-family ranges (see
//   FunctionDef::accepts_ranges comment in function_registry.h). A true
//   range-context concept is deferred.

// COUNT(value, ...) - count of Number values. Booleans, text (even
// numeric-looking text like "5"), blanks, and errors are all skipped.
// Direct-arg booleans are also skipped (Excel's COUNT never counts
// booleans even when they appear outside a range).
Value Count(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (args[i].is_number()) {
      total += 1.0;
    }
  }
  return Value::number(total);
}

// COUNTA(value, ...) - count of non-Blank values. Numbers, booleans, text
// (including the empty string produced by a formula returning ""), and
// errors are all counted. Only the Blank scalar is skipped.
Value CountA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (!args[i].is_blank()) {
      total += 1.0;
    }
  }
  return Value::number(total);
}

// COUNTBLANK(value, ...) - count of Blank scalars and Text values whose
// contents are exactly "". Numbers (including 0), booleans (including
// FALSE), non-empty text, and errors are all skipped. The public Excel 365
// signature accepts a single range; we accept variadic for symmetry with
// the sibling aggregators - a single A1:B2 ref still expands to many
// scalar args via the dispatcher.
Value CountBlank(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    const Value& v = args[i];
    if (v.is_blank() || (v.is_text() && v.as_text().empty())) {
      total += 1.0;
    }
  }
  return Value::number(total);
}

}  // namespace

void register_aggregate_builtins(FunctionRegistry& registry) {
  {
    // SUM is range-aware: `=SUM(A1:A100)` expands the rectangle into scalar
    // cell values before this impl runs.
    FunctionDef def{"SUM", 1u, kVariadic, &Sum};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  registry.register_function(FunctionDef{"CONCAT", 1u, kVariadic, &Concat});
  // CONCATENATE is the legacy spelling kept by Excel for compatibility; it
  // shares the implementation with CONCAT.
  registry.register_function(FunctionDef{"CONCATENATE", 1u, kVariadic, &Concat});
  registry.register_function(FunctionDef{"LEN", 1u, 1u, &Len});

  // Aggregates (min_arity = 1, variadic). Each is range-aware: a RangeOp
  // argument is flattened into scalar cell values by the dispatcher before
  // the impl runs.
  {
    FunctionDef def{"MIN", 1u, kVariadic, &Min};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"MAX", 1u, kVariadic, &Max};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"AVERAGE", 1u, kVariadic, &Average};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"PRODUCT", 1u, kVariadic, &Product};
    def.accepts_ranges = true;
    registry.register_function(def);
  }

  // Counting aggregators. All three are range-aware and opt out of the
  // dispatcher's left-most-error rule so the impl itself decides which
  // values to count (see the block comment above `Count` in this file for
  // the range-vs-direct divergence).
  {
    FunctionDef def{"COUNT", 1u, kVariadic, &Count, /*propagate_errors=*/false};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"COUNTA", 1u, kVariadic, &CountA, /*propagate_errors=*/false};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"COUNTBLANK", 1u, kVariadic, &CountBlank, /*propagate_errors=*/false};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
}

}  // namespace eval
}  // namespace formulon
