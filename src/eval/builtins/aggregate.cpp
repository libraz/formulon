//
// Implementation of Formulon's aggregate built-in functions:
// SUM, SUMSQ, MIN, MAX, AVERAGE, PRODUCT, COUNT, COUNTA, COUNTBLANK, CONCAT,
// CONCATENATE, LEN, and PERCENTOF. Most impls follow the same recipe as the
// rest of the builtin catalog: coerce arguments via `eval/coerce.h`,
// propagate the left-most coercion error, and return a `Value`. PERCENTOF
// is the exception — it must compute per-argument totals, so it lives at
// the bottom of this file as a lazy impl (registered via the central
// dispatch table in `tree_walker.cpp` rather than this file's registrar).

#include "eval/builtins/aggregate.h"

#include <cmath>
#include <cstdint>
#include <functional>
#include <string>
#include <vector>

#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/utf8_length.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
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

// Shared kernel for MIN / MAX. A literal non-numeric argument coerces via
// `coerce_to_number` and surfaces `#VALUE!` on failure. The caller's
// pre-evaluation has already short-circuited any argument that was itself
// an error. When every argument is filtered out by the range-vs-direct
// provenance rule (e.g. `=MIN(A1:A3)` over an empty / all-text range),
// Excel returns 0 rather than an error.
//
// `Cmp` is invoked as `Cmp{}(candidate, current_best)` and must return
// true when `candidate` should replace `current_best`. Passing
// `std::less<>` therefore selects the minimum; `std::greater<>` selects
// the maximum.
template <typename Cmp>
Value extreme(const Value* args, std::uint32_t arity) {
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
    if (Cmp{}(coerced.value(), best)) {
      best = coerced.value();
    }
  }
  if (std::isnan(best) || std::isinf(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

// MIN(value, ...) - smallest of the coerced numbers.
Value Min(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return extreme<std::less<double>>(args, arity);
}

// MAX(value, ...) - symmetric to MIN. Empty post-filter arity also returns 0.
Value Max(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return extreme<std::greater<double>>(args, arity);
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

// SUMSQ(value, ...) - sum of squares. Follows the same provenance rule as
// SUM: direct scalar args coerce through `coerce_to_number` (so TRUE -> 1,
// "5" -> 5, "abc" -> #VALUE!), while Bool / Text / Blank cells sourced
// from a range are dropped before the impl runs (the dispatcher's
// `range_filter_numeric_only` flag). Any error in the argument list
// propagates left-to-right.
Value SumSq(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    const double x = coerced.value();
    total += x * x;
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
// errors are all counted. Plain blanks and raw-reference blanks remain
// skipped; a blank explicitly owned by a derived value array counts as an
// occupied array cell.
Value CountA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (!args[i].is_blank() || args[i].blank_counts_for_counta()) {
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

// SUM-style per-argument total used by PERCENTOF. Mirrors the eager SUM
// dispatch path's two provenance rules:
//
//   * For range-shaped args (`Ref` / `RangeOp` / `SpillRef`, plus the
//     reference-producing calls `OFFSET` / `CHOOSE` / `IF` / `ROW` /
//     `COLUMN` that `resolve_range_arg` already expands) and for inline
//     `{a;b;c}` array literals, Bool / Text / Blank cells are silently
//     dropped. This is the same filter that `range_filter_numeric_only`
//     applies in `tree_walker.cpp`.
//   * For any other (scalar) arg, `coerce_to_number` runs: TRUE -> 1,
//     FALSE -> 0, blank -> 0, "5" -> 5, "abc" -> #VALUE!.
//
// On success writes the computed total to `*out_total`. On any error
// (range expansion failure, error cell inside a range, error scalar,
// non-coercible scalar text) writes the propagating Value to `*out_err`
// and returns `false`. Errors discovered while walking a range follow
// canonical row-major scan order, matching SUM and SUMPRODUCT.
//
// This is *not* a refactor of SUM: SUM rides the eager dispatcher because
// it produces a single total over its concatenated arg vector. PERCENTOF
// must compute per-argument totals (numerator vs. denominator), which the
// eager path collapses into one flat values vector — that boundary is the
// architectural reason this helper exists alongside SUM rather than being
// extracted from it.
bool sum_arg_for_percentof(const parser::AstNode& raw_arg, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, double* out_total, Value* out_err) {
  // LET-binding passthrough: `=LET(r, A1:A3, PERCENTOF(r, A1:A10))` parses
  // `r` as a `NameRef`; we want to see the bound RangeOp / ArrayLiteral /
  // OFFSET-call shape so the kind dispatch below treats it the same way as
  // the literal argument. Single-cell Refs and pure scalar bindings stay as
  // NameRef so the scalar-coerce branch handles them via `eval_node`.
  const parser::AstNode* effective = &raw_arg;
  if (raw_arg.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(raw_arg, ctx.name_env());
    if (&resolved != &raw_arg && is_range_shaped_ast(resolved)) {
      effective = &resolved;
    }
  }
  const parser::AstNode& arg_node = *effective;
  const parser::NodeKind k = arg_node.kind();

  // Range-shaped args (and the reference-producing calls that expand to
  // rectangles) go through `resolve_range_arg`; cells receive the
  // numeric-only provenance filter.
  const bool is_range_call =
      k == parser::NodeKind::Call && (strings::case_insensitive_eq(arg_node.as_call_name(), "OFFSET") ||
                                      strings::case_insensitive_eq(arg_node.as_call_name(), "CHOOSE") ||
                                      strings::case_insensitive_eq(arg_node.as_call_name(), "IF") ||
                                      strings::case_insensitive_eq(arg_node.as_call_name(), "ROW") ||
                                      strings::case_insensitive_eq(arg_node.as_call_name(), "COLUMN"));
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp || k == parser::NodeKind::SpillRef ||
      is_range_call) {
    auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
    if (!resolved) {
      *out_err = Value::error(resolved.error());
      return false;
    }
    double total = 0.0;
    for (const Value& v : resolved.value().cells) {
      if (v.is_error()) {
        *out_err = v;
        return false;
      }
      // Range-sourced provenance: only Number cells contribute.
      if (!v.is_number()) {
        continue;
      }
      total += v.as_number();
    }
    if (std::isnan(total) || std::isinf(total)) {
      *out_err = Value::error(ErrorCode::Num);
      return false;
    }
    *out_total = total;
    return true;
  }

  // Inline array literal: walk in row-major order with the same range-like
  // filter (Bool / Text / Blank are dropped, errors propagate).
  if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = arg_node.as_array_rows();
    const std::uint32_t cols = arg_node.as_array_cols();
    double total = 0.0;
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        const Value v = eval_node(arg_node.as_array_element(r, c), arena, registry, ctx);
        if (v.is_error()) {
          *out_err = v;
          return false;
        }
        if (!v.is_number()) {
          continue;
        }
        total += v.as_number();
      }
    }
    if (std::isnan(total) || std::isinf(total)) {
      *out_err = Value::error(ErrorCode::Num);
      return false;
    }
    *out_total = total;
    return true;
  }

  // Scalar argument: evaluate and coerce strictly (matching SUM's direct-
  // scalar branch). TRUE -> 1, "5" -> 5, "abc" -> #VALUE!, blank -> 0.
  const Value v = eval_node(arg_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return false;
  }
  const double total = coerced.value();
  if (std::isnan(total) || std::isinf(total)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  *out_total = total;
  return true;
}

}  // namespace

// PERCENTOF(data_subset, data_all)
//
// Excel 365 (Sep 2024 GA) ratio aggregator. Returns
// `SUM(data_subset) / SUM(data_all)` with SUM-style coercion:
//
//   * Range-sourced cells: Bool / Text / Blank skipped silently.
//   * Inline `{...}` array literals: same filter as ranges.
//   * Direct scalar args: full `coerce_to_number` rules apply, so a direct
//     TRUE coerces to 1, "5" to 5, and "abc" to #VALUE!.
//
// Error precedence (left-to-right):
//   1. `data_subset` evaluation / coercion error -> that error.
//   2. `data_all` evaluation / coercion error    -> that error.
//   3. `SUM(data_all) == 0`                      -> #DIV/0!.
//
// WHY a lazy impl rather than the eager `accepts_ranges` registration the
// task brief proposes: the eager dispatcher in `tree_walker.cpp`
// concatenates every flattened argument into a single `values` vector
// before invoking the impl, so an eager PERCENTOF would have no way to
// distinguish where `data_subset` ends and `data_all` begins for a call
// like `=PERCENTOF(A1:A2, B1:B3)`. We need per-argument totals, hence
// lazy dispatch — the same architectural reason SUMPRODUCT and the *IFS
// family are lazy.
Value eval_percentof_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }
  double numerator = 0.0;
  Value err = Value::blank();
  if (!sum_arg_for_percentof(call.as_call_arg(0), arena, registry, ctx, &numerator, &err)) {
    return err;
  }
  double denominator = 0.0;
  if (!sum_arg_for_percentof(call.as_call_arg(1), arena, registry, ctx, &denominator, &err)) {
    return err;
  }
  if (denominator == 0.0) {
    return Value::error(ErrorCode::Div0);
  }
  const double r = numerator / denominator;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

void register_aggregate_builtins(FunctionRegistry& registry) {
  // SUM / MIN / MAX / AVERAGE / PRODUCT / SUMSQ are range-aware: a RangeOp
  // argument is flattened into scalar cell values by the dispatcher before
  // the impl runs. The `range_filter_numeric_only` flag mirrors Excel's
  // provenance rule: Bool / Text / Blank cells sourced from a range are
  // dropped silently, while direct scalar arguments continue through normal
  // coercion.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"SUM", 1u, kVariadic, &Sum, true, true, true},
      // CONCAT accepts ranges: cells are coerced to text in row-major order.
      {"CONCAT", 1u, kVariadic, &Concat, true, true},
      // CONCATENATE keeps legacy implicit-intersection semantics, so it is
      // intentionally not range-aware.
      {"CONCATENATE", 1u, kVariadic, &Concat},
      {"LEN", 1u, 1u, &Len},
      {"MIN", 1u, kVariadic, &Min, true, true, true},
      {"MAX", 1u, kVariadic, &Max, true, true, true},
      {"AVERAGE", 1u, kVariadic, &Average, true, true, true},
      {"PRODUCT", 1u, kVariadic, &Product, true, true, true},
      {"SUMSQ", 1u, kVariadic, &SumSq, true, true, true},
      {"COUNTA", 1u, kVariadic, &CountA, false, true},
      {"COUNTBLANK", 1u, kVariadic, &CountBlank, false, true},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));

  // SUMPRODUCT is routed through the lazy dispatch table (see
  // `eval_sumproduct_lazy` in `shape_ops_lazy.cpp`) because it must
  // preserve each argument's (rows, cols) shape to shape-check parallel
  // rectangles, and must walk inline `{...}` array literals element by
  // element. Pre-evaluating every arg via the eager path would erase
  // both pieces of information.
  //
  // PERCENTOF (Excel 365, Sep 2024 GA) is also lazy: see
  // `eval_percentof_lazy` near the top of this file. It needs per-argument
  // SUM totals (numerator vs. denominator), which the eager dispatcher
  // collapses into a single concatenated `args[]` vector.

  // COUNT is registered as a lazy impl in `tree_walker.cpp`, not here,
  // because it needs per-arg AST shape to apply Excel's direct-vs-range
  // provenance rule for Bool values.
}

}  // namespace eval
}  // namespace formulon
