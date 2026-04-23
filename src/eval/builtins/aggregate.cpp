// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's aggregate built-in functions:
// SUM, MIN, MAX, AVERAGE, PRODUCT, COUNT, COUNTA, COUNTBLANK, CONCAT,
// CONCATENATE, and LEN. Each impl follows the same recipe as the rest of
// the builtin catalog: coerce arguments via `eval/coerce.h`, propagate the
// left-most coercion error, and return a `Value`.

#include "eval/builtins/aggregate.h"

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

// MIN(value, ...) - smallest of the coerced numbers. A literal non-numeric
// argument coerces via `coerce_to_number` and surfaces `#VALUE!` on failure.
// The caller's pre-evaluation has already short-circuited any argument that
// was itself an error. When every argument is filtered out by the
// range-vs-direct provenance rule (e.g. `=MIN(A1:A3)` over an empty / all-
// text range), Excel returns 0 rather than an error.
Value Min(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  if (arity == 0) {
    return Value::number(0.0);
  }
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

// MAX(value, ...) - symmetric to MIN. Empty post-filter arity also returns 0.
Value Max(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  if (arity == 0) {
    return Value::number(0.0);
  }
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

// AVERAGE(value, ...) - arithmetic mean. The registry enforces min_arity=1
// on the pre-expansion argument count, but the provenance filter may drop
// every range-sourced value before this impl runs; in that case Excel
// reports #DIV/0!.
Value Average(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  if (arity == 0) {
    return Value::error(ErrorCode::Div0);
  }
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
// When every argument was filtered out by the range-vs-direct provenance
// rule (e.g. `=PRODUCT(A1:A3)` over an empty / all-text range), Excel
// returns 0 rather than the mathematical identity 1.
Value Product(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  if (arity == 0) {
    return Value::number(0.0);
  }
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
// COUNTA / COUNTBLANK. Both are registered with `propagate_errors = false`:
// Excel's COUNT family is specified in terms of which cells to "count"
// rather than which values to coerce, so an error inside a range must not
// short-circuit the whole call. That opt-out means the impls see
// Error-typed values in their args array directly and must skip them
// explicitly.
//
// COUNT itself is routed through the lazy dispatch table (see
// `eval_count_lazy` in `special_forms_lazy.cpp`) because Excel counts Bool
// values differently depending on whether they are direct arguments or
// sourced from a range: `=COUNT(1, TRUE, 3)` is 3, but `=COUNT(A1:A3)`
// where A2 holds TRUE is 2. That provenance distinction requires per-arg
// AST inspection that the eager dispatcher's flattened values vector has
// already erased.

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
    // cell values before this impl runs. Range-sourced Bool / Text / Blank
    // cells are silently skipped to match Excel's provenance rule; direct
    // arguments continue to coerce normally.
    FunctionDef def{"SUM", 1u, kVariadic, &Sum};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  registry.register_function(FunctionDef{"CONCAT", 1u, kVariadic, &Concat});
  // CONCATENATE is the legacy spelling kept by Excel for compatibility; it
  // shares the implementation with CONCAT.
  registry.register_function(FunctionDef{"CONCATENATE", 1u, kVariadic, &Concat});
  registry.register_function(FunctionDef{"LEN", 1u, 1u, &Len});

  // Aggregates (min_arity = 1, variadic). Each is range-aware: a RangeOp
  // argument is flattened into scalar cell values by the dispatcher before
  // the impl runs. The `range_filter_numeric_only` flag mirrors Excel's
  // provenance rule: Bool / Text / Blank cells sourced from a range are
  // dropped silently, while direct scalar arguments continue through
  // normal coercion (so =SUM(10,TRUE,30) is 41 but =SUM(A1:A3) with a
  // TRUE cell is 40).
  {
    FunctionDef def{"MIN", 1u, kVariadic, &Min};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"MAX", 1u, kVariadic, &Max};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"AVERAGE", 1u, kVariadic, &Average};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"PRODUCT", 1u, kVariadic, &Product};
    def.accepts_ranges = true;
    def.range_filter_numeric_only = true;
    registry.register_function(def);
  }

  // Counting aggregators. Both are range-aware and opt out of the
  // dispatcher's left-most-error rule so the impl itself decides which
  // values to count. COUNT is registered as a lazy impl in
  // `tree_walker.cpp`, not here, because it needs per-arg AST shape to
  // apply Excel's direct-vs-range provenance rule for Bool values.
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
