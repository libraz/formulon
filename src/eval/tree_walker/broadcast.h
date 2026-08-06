//
// Private seam between the tree-walker's recursive node visitor
// (`tree_walker/walker.cpp`) and the top-level array-broadcasting
// helpers (`tree_walker/broadcast.cpp`). The two translation units were
// split out of the original monolithic `tree_walker.cpp`; this header
// publishes the three entry points walker.cpp needs from broadcast.cpp.
//
// `apply_binop_per_cell` is the scalar fast path used both by walker.cpp
// (when neither operand is an Array) and by `broadcast_binop` (for each
// cell of the broadcast output).
//
// `broadcast_binop` / `broadcast_unary` handle the Array-aware case for
// BinaryOp / UnaryOp AST nodes: they accept already-evaluated `Value`s
// (no second AST walk) and shape the output to match `eval_binop_array_ctx`
// in `shape_ops_lazy.cpp`. The caller (walker.cpp) decides whether the
// resulting `Value::Array` spills (via `EvalContext::dispatch_array_result`)
// or feeds into a downstream operator that consumes it as an Array.
//
// This header is internal to the tree-walker family and is not part of
// the public evaluator surface — production callers reach the evaluator
// through `eval/tree_walker.h`.

#ifndef FORMULON_EVAL_TREE_WALKER_BROADCAST_H_
#define FORMULON_EVAL_TREE_WALKER_BROADCAST_H_

#include <cstdint>

#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

// A non-owning shape + cell view over a `Value`, used by the broadcast
// helpers. For an Array it aliases the existing cells buffer (no copy);
// for a scalar it aliases a caller-supplied 1-element backing slot.
struct ArrayView {
  std::uint32_t rows;
  std::uint32_t cols;
  const Value* cells;
};

// Resolves `v` to an `ArrayView`. For an Array the view aliases the
// existing cells buffer; for a scalar the caller-supplied 1-element
// backing slot `scalar_slot` is populated and aliased. Lifetime: the view
// is valid as long as either the source Array or `scalar_slot` outlives it.
ArrayView as_array_view(const Value& v, Value* scalar_slot);

// Fetches the operand cell contributing to output position `(r, c)` under
// Excel broadcasting, or `nullptr` when the operand cannot supply that
// position (a non-1 axis shorter than the output -> `#N/A`). A size-1 axis
// always reads index 0 (broadcast stretch).
const Value* broadcast_cell(const ArrayView& v, std::uint32_t r, std::uint32_t c);

// Scalar per-cell binary-operator application. Errors propagate
// left-to-right; numeric arithmetic coerces both operands through
// `coerce_to_number`; comparisons / concat dispatch through the
// `scalar_ops.h` helpers.
Value apply_binop_per_cell(parser::BinOp op, const Value& lhs, const Value& rhs, Arena& arena);

// Shape-aware binary-operator broadcast. When either operand is an
// `ArrayValue`, walks the result rectangle (with 1x1 scalar broadcast on
// either side) and produces a `Value::Array` whose cells are the
// per-cell `apply_binop_per_cell` results. Mismatched non-1x1 shapes
// surface scalar `#VALUE!` (Excel's whole-expression short-circuit).
Value broadcast_binop(parser::BinOp op, const Value& lhs, const Value& rhs, Arena& arena);

// Shape-aware unary-operator broadcast. For a scalar operand, returns
// the plain `apply_unary` result. For an Array operand, returns a
// `Value::Array` of the same shape with `apply_unary` applied per cell
// (errors are written through verbatim).
Value broadcast_unary(parser::UnaryOp op, const Value& operand, Arena& arena);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_BROADCAST_H_
