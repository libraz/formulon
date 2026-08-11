//
// Range-argument resolution helper shared by the conditional aggregators
// (`COUNTIF` / `SUMIF` / `AVERAGEIF` / `COUNTIFS` / ...) and the lookup
// family (`VLOOKUP` / `HLOOKUP` / `INDEX` / `MATCH` / `XLOOKUP`). These
// builtins all need to turn a lazily-parsed AST argument — which may be
// either a `RangeOp` (`A1:B2`) or a bare `Ref` — into a flat row-major
// vector of cell `Value`s plus an optional (rows, cols) shape.
//
// Hoisting this helper out of `tree_walker.cpp` lets the conditional and
// lookup impls be split into their own translation units without each TU
// pulling in the rest of the evaluator's internals.

#ifndef FORMULON_EVAL_RANGE_ARGS_H_
#define FORMULON_EVAL_RANGE_ARGS_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

class Arena;

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;
struct FunctionDef;

/// Applies the FunctionDef's provenance-aware range filter to one cell.
bool append_range_sourced_value(const FunctionDef& def, const Value& value, std::vector<Value>* values, Value* out_err);

/// Applies the same filter in source order to a contiguous range.
bool append_range_sourced_values(const FunctionDef& def, const Value* cells, std::size_t count,
                                 std::vector<Value>* values, Value* out_err);

/// Applies the same filter while compacting retained values into an arena
/// buffer. `out_cells` must have room for `count` values; the retained count
/// may be zero after a valid non-empty input.
bool filter_range_sourced_values(const FunctionDef& def, const Value* cells, std::size_t count, Value* out_cells,
                                 std::size_t* out_count, Value* out_err);

/// Resolution result for a range-shaped argument: a flat row-major vector
/// of cell `Value`s plus the rectangle's shape. A 1-cell `Ref` produces
/// `rows = cols = 1`; a `RangeOp` produces the computed shape; a flattened
/// `Value::Array` from a dynamic-array producer carries the array's own
/// dimensions.
struct RangeResult {
  std::vector<Value> cells;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
};

/// Resolves `arg_node` to a `RangeResult`, expanding `RangeOp` / `Ref` /
/// `SpillRef` / range-shaped Calls (`OFFSET` / `CHOOSE` / `IF` / `ROW` /
/// `COLUMN`) and unwrapping dynamic-array producers (`MUNIT`,
/// `SEQUENCE`, lambda invocations, …) into row-major cells.
///
/// Excel-level errors (`#REF!` for a missing sheet, `#VALUE!` for an
/// unsupported argument shape, propagated cell errors, …) are returned
/// as the `Unexpected<ErrorCode>` payload. The Excel `ErrorCode` is the
/// correct error vocabulary here because `resolve_range_arg` is invoked
/// during formula evaluation; engine-level `Error` codes (zip-bomb,
/// arena exhaustion, …) cannot originate here.
///
/// This is the only public surface; the previous `bool + out_param`
/// transitional overload was deleted once the lazy-family fan-out
/// finished migrating.
Expected<RangeResult, ErrorCode> resolve_range_arg(const parser::AstNode& arg_node, Arena& arena,
                                                   const FunctionRegistry& registry, const EvalContext& ctx);

/// Materialises an AST argument as an `ArrayValue` via the standard
/// `eval_node_as_array` seam. Scalar inputs wrap to a 1x1 array. On
/// failure writes the propagating error into `*out_err` and returns
/// false; otherwise `*out` borrows from the caller arena.
///
/// Errors come in two flavours:
///   * a pre-evaluated error subtree propagates verbatim;
///   * a non-array, non-error result (e.g. an unexpected scalar) yields
///     `#VALUE!`.
///
/// Callers that need a different error vocabulary (e.g. `#N/A` for the
/// regression / hypothesis families) should layer their own remapping
/// on top - this helper is the lowest-common-denominator path used by
/// SUMPRODUCT-style families.
bool resolve_array_value(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, const ArrayValue** out, Value* out_err);

/// Resolution variant for the regression and hypothesis-test families.
/// Accepts `Ref` / `RangeOp` (delegated to `resolve_range_arg`),
/// `ArrayLiteral` (each element is evaluated), and any other subtree
/// (a pre-evaluated error propagates verbatim; otherwise the call is
/// rejected with `#N/A`).
///
/// Excel uses `#N/A` rather than `#VALUE!` as the shape-error vocabulary
/// for these families, so any `#VALUE!` from `resolve_range_arg`'s
/// shape rejection is remapped to `#N/A`. `#REF!` passes through.
Expected<RangeResult, ErrorCode> resolve_array_arg_na(const parser::AstNode& arg_node, Arena& arena,
                                                      const FunctionRegistry& registry, const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RANGE_ARGS_H_
