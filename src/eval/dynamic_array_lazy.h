// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for dynamic-array spilling builtins that need per-argument AST
// shape inspection: `FILTER`, future `SORT` / `UNIQUE` / `SORTBY` / `RANDARRAY`.
// These functions either produce an `ArrayValue` whose footprint depends on
// the input array's 2D shape, or accept a range argument they must keep as a
// 2D rectangle (the eager dispatcher would flatten range cells into a 1D
// vector and lose the shape).
//
// Sibling of `eval/shape_ops_lazy.h` (which hosts the SUMPRODUCT-side helpers
// and TRANSPOSE) and `eval/builtins/dynamic_array.cpp` (the eager-arg
// SEQUENCE impl). Externs registered in the central `kLazyDispatch` table in
// `tree_walker.cpp`; see `eval/lazy_impls.h` for the shared signature.

#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_LAZY_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `FILTER(array, include, [if_empty])` — returns a subset of `array`
/// determined by a parallel boolean mask `include`.
///
/// Shape rules:
///   * `include` must be a 1D array matching one of `array`'s axes:
///     either `(array.rows, 1)` to filter rows, or `(1, array.cols)` to
///     filter columns. Anything else surfaces `#VALUE!`.
///   * Each `include` cell coerces to bool via `coerce_to_bool`. Coercion
///     errors short-circuit and propagate (the entire result is the error).
///   * Cells in `include` evaluating to TRUE keep the corresponding
///     row / column in `array`; FALSE drops it.
///
/// Empty result handling:
///   * If no rows / cols are kept and `if_empty` is provided, `if_empty`
///     is returned scalar (Excel does not spill it; matches Mac).
///   * If no rows / cols are kept and `if_empty` is omitted, `#CALC!`.
///
/// Errors in `array` cells are preserved verbatim in the output; FILTER
/// does not coerce or evaluate cell contents (it only routes them).
/// Errors in `include` cells (other than coercion failures, which were
/// short-circuited above) are not currently distinguished from FALSE —
/// Mac Excel actually propagates per-cell errors here, but the conservative
/// "any error short-circuits" path matches the dominant case and avoids a
/// surprise spill of mixed bool / error cells.
Value eval_filter_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_LAZY_H_
