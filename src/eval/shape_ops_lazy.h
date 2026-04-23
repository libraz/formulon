// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for shape / geometry-inspection builtins: `ROWS`, `COLUMNS`,
// `ROW`, `COLUMN`, and `SUMPRODUCT`. These all require per-argument AST
// shape to distinguish a `Ref` (single cell), a `RangeOp` (rectangle),
// and an `ArrayLiteral` (inline `{...}` literal) — information the eager
// dispatcher would already have discarded by flattening each arg to a
// `Value`.
//
// Publishes externs consumed by the central `kLazyDispatch` table in
// `tree_walker.cpp`; see `eval/lazy_impls.h` for the shared
// `LazyImpl` signature and the `eval_node` entry point these impls
// recurse through.

#ifndef FORMULON_EVAL_SHAPE_OPS_LAZY_H_
#define FORMULON_EVAL_SHAPE_OPS_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `ROWS(array)` - returns the row count of a reference, range, or array
/// literal. A bare scalar (or any non-reference expression that evaluates
/// without error) is treated as a 1-row value. Errors propagate.
Value eval_rows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `COLUMNS(array)` - symmetric to `ROWS`, returning the column count.
Value eval_columns_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// `ROW(reference)` - returns the 1-based row index of a cell reference
/// (or the first row of a RangeOp). An array literal yields `#VALUE!`
/// because it is not a reference. Errors propagate. The zero-argument
/// form (`=ROW()` returning the caller cell's row) is not supported.
Value eval_row_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

/// `COLUMN(reference)` - symmetric to `ROW`, returning the 1-based column
/// index. Same shape constraints, same restriction on zero-arity.
Value eval_column_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `SUMPRODUCT(arr1, arr2, ...)` - element-wise product accumulated into a
/// sum across N parallel rectangles. All arguments must have identical
/// `(rows, cols)` shape or the call returns `#VALUE!`. Range and array-
/// literal args are accepted; scalar args degenerate to 1x1. Non-numeric
/// cells (Bool, Text, Blank) contribute 0; errors propagate in row-major
/// scan order across args.
Value eval_sumproduct_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SHAPE_OPS_LAZY_H_
