//
// Private seam between the tree-walker's recursive node visitor
// (`tree_walker/walker.cpp`) and the top-level array-broadcasting
// helpers (`tree_walker/broadcast.cpp`). The two translation units were
// split out of the original monolithic `tree_walker.cpp`; this header
// publishes the three entry points walker.cpp needs from broadcast.cpp, plus
// the scalar-function broadcast that dispatch.cpp calls.
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
#include <vector>

#include "eval/tail_array.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

struct FunctionDef;

// A non-owning shape + cell view over a `Value`, used by the broadcast
// helpers. For an Array it aliases the existing cells buffer (no copy);
// for a scalar it aliases a caller-supplied 1-element backing slot.
// For a `TailArray`, `rows` x `cols` is the declared size, `cells` is the dense
// head (`head` along `axis`) and `tail` is the repeated row / column.
struct ArrayView {
  std::uint32_t rows;
  std::uint32_t cols;
  const Value* cells;
  std::uint32_t head = 0;
  const Value* tail = nullptr;
  TailAxis axis = TailAxis::kRows;
};

// Resolves `v` to an `ArrayView`. For an Array the view aliases the
// existing cells buffer; for a scalar the caller-supplied 1-element
// backing slot `scalar_slot` is populated and aliased. Lifetime: the view
// is valid as long as either the source Array or `scalar_slot` outlives it.
ArrayView as_array_view(const Value& v, Value* scalar_slot);

// `Shaped` overload: a `TailArray` yields a tail-aware view, anything else
// behaves like the `Value` overload on `s.value`.
ArrayView as_array_view(const Shaped& s, Value* scalar_slot);

// Fetches the operand cell contributing to output position `(r, c)` under
// Excel broadcasting, or `nullptr` when the operand cannot supply that
// position (a non-1 axis shorter than the output -> `#N/A`). A size-1 axis
// always reads index 0 (broadcast stretch).
const Value* broadcast_cell(const ArrayView& v, std::uint32_t r, std::uint32_t c);

// Widens `*rows` x `*cols` to cover `v`: each axis takes the larger extent.
void widen_shape(const ArrayView& v, std::uint32_t* rows, std::uint32_t* cols);

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

// `Shaped` overload of `broadcast_binop`. Without a `TailArray` operand it
// delegates to the `Value` version. With `TailArray` operands on one axis the
// result is a `TailArray` (never `from_reference`): head cells are computed
// per cell and the tail row / column once. Operands on different axes are
// densified first (`#NUM!` past the derived-array cap).
Shaped broadcast_binop(parser::BinOp op, const Shaped& lhs, const Shaped& rhs, Arena& arena);

// `Shaped` overload of `broadcast_unary`; a `TailArray` operand yields a
// `TailArray` of the same declared size.
Shaped broadcast_unary(parser::UnaryOp op, const Shaped& operand, Arena& arena);

// Evaluates a scalar (non-range-aware) function element-wise across the
// broadcast rectangle of its already-collected arguments and returns a
// spilled `Value::Array`. `out_rows` / `out_cols` are the max extents
// over the array arguments; scalars broadcast to every cell. A 1x1
// result unwraps to a plain scalar so a single-cell range argument does
// not spill. Per-cell error propagation mirrors the scalar dispatch
// path: with `propagate_errors`, the first error argument short-circuits
// that cell before the impl runs.
Value broadcast_scalar_call(const FunctionDef& def, const std::vector<Value>& args, std::uint32_t out_rows,
                            std::uint32_t out_cols, Arena& arena);

// Evaluates a `last_arg_scalar` function once per element of `last`, the
// array its last argument evaluated to, with `values` holding the other
// arguments. A 1x1 result unwraps to a scalar.
Value lift_last_arg_call(const FunctionDef& def, std::vector<Value> values, const ArrayValue& last, Arena& arena);

// `broadcast_scalar_call` for arguments of which `tails[i]` (when non-null) is
// a whole column / row standing in for `args[i]`. Arguments on one axis give a
// `TailArray`: the function is applied to the head cells and once to the tail
// row / column. Whole columns mixed with whole rows are expanded and take the
// dense route.
Shaped broadcast_scalar_call_shaped(const FunctionDef& def, const std::vector<Value>& args,
                                    const std::vector<const TailArray*>& tails, Arena& arena);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_BROADCAST_H_
