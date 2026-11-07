//
// The array-context evaluation seam: `eval_node_as_array` and the cellwise
// binary broadcaster behind it. Shared by the lazy families that need a
// shape-preserving `Value::Array` from an AST node.

#ifndef FORMULON_EVAL_ARRAY_CONTEXT_H_
#define FORMULON_EVAL_ARRAY_CONTEXT_H_

#include <cstdint>
#include <vector>

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Evaluates `node` in array context, always returning a `Value::Array`
/// (or a scalar `Value::error(...)` on failure).
///
/// Dispatch by AST kind:
///   * `Ref` / `RangeOp` / `Call(OFFSET|CHOOSE|IF|ROW|COLUMN)` -> expand
///     via `resolve_range_arg`; wrap the resulting flat cells + shape in
///     a freshly arena-allocated `ArrayValue`.
///   * `ArrayLiteral` -> walk via the existing `resolve_error_free_range`
///     pattern; wrap into an `ArrayValue`.
///   * `BinaryOp` -> recurse into `eval_binop_array_ctx`.
///   * `UnaryOp` -> recurse into the operand via `eval_node_as_array`,
///     then apply the unary op cellwise.
///   * Anything else (Number / Text / Bool literal, function `Call`
///     other than the reference-shaped ones above) -> evaluate via plain
///     `eval_node`; if scalar, wrap into a 1x1 `ArrayValue`; if Array,
///     forward (this can only happen if a future array-aware function
///     starts producing Arrays). Errors propagate as scalar `Value::error`.
///
/// Lifetime: the returned `ArrayValue` and its cells live in `arena`;
/// the caller must keep `arena` alive for the lifetime of the value.
Value eval_node_as_array(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

/// Evaluates a `BinaryOp` AST in array context, broadcasting cellwise.
///
/// Recursively evaluates both operands via `eval_node_as_array`. Shape
/// rules:
///   * 1x1 broadcasts against any shape (the 1x1 cell is repeated).
///   * Non-1x1 shapes must match exactly; mismatch yields a scalar
///     `#VALUE!` (NOT an `ArrayValue` of `#VALUE!` cells - Mac Excel
///     short-circuits the whole expression).
///   * Errors in either operand short-circuit and propagate as scalar.
///
/// For each output cell, dispatches to the appropriate scalar helper
/// from `eval/scalar_ops.h`:
///   * arithmetic ops (Add/Sub/Mul/Div/Pow): coerce both cells to
///     numbers via `coerce_to_number`, then `apply_arithmetic`. Cell-
///     level coercion errors become per-cell `Value::error`, which
///     populates the output array (arithmetic over an Array of error
///     cells preserves the errors verbatim).
///   * concat (`&`): `apply_concat`.
///   * comparison ops: `apply_comparison`.
Value eval_binop_array_ctx(const parser::AstNode& binary_op_node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

namespace detail {

/// Builds an arena-allocated `ArrayValue` from a flat row-major cell vector and
/// shape; `nullptr` when the shape is rejected or the arena is exhausted.
const ArrayValue* make_array_value(Arena& arena, std::uint32_t rows, std::uint32_t cols,
                                   const std::vector<Value>& cells);

/// Wraps `arr` as a `Value`, or `#NUM!` when it is null.
Value array_value_or_num_error(const ArrayValue* arr);

/// True for the reference-shaped `Call` names `resolve_range_arg` expands into
/// a flat rectangle.
bool is_range_producing_call(const parser::AstNode& call_node);

}  // namespace detail

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_ARRAY_CONTEXT_H_
