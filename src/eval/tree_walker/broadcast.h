// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

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
