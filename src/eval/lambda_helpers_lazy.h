//
// Lazy impls for Excel 365's six LAMBDA-helper functions: BYROW, BYCOL,
// MAP, REDUCE, SCAN, MAKEARRAY. Each consumes a `Lambda` value as one of
// its arguments and applies it across a range / array, threading the
// captured environment through every per-cell call.
//
// Why lazy: the lambda argument is a free-form `Lambda` AST node that the
// eager dispatcher would already have evaluated to a `Value::Lambda` (a
// closure) — the helpers don't need anything more from the AST level. But
// the array argument(s) need to be evaluated in array context so 2D shape
// is preserved (the same reason TRIMRANGE / FILTER / SORT live in the
// lazy tier), and every per-cell call has to bind freshly evaluated
// arguments into the lambda's captured environment without re-running the
// caller's argument expressions. Splitting the helpers off into their own
// TU keeps the central dispatch table thin and lets the per-cell
// invocation pipeline live in one place.
//
// All six impls require the lambda argument to evaluate to a Lambda and
// to declare the expected parameter count for that helper:
//   * BYROW / BYCOL   — 1 param.
//   * MAP             — N params, where N is the number of array args.
//   * REDUCE / SCAN   — 2 params (accumulator, current).
//   * MAKEARRAY       — 2 params (row_index, col_index).
// Wrong type / wrong arity surfaces `#VALUE!`.
//
// Per-cell error handling: any error returned from the lambda is propagated
// as the helper's whole result (Mac Excel: a single error short-circuits
// the entire spill). The helpers that produce a 2D output (BYROW / BYCOL /
// MAP / MAKEARRAY) additionally reject array results from the lambda with
// `#CALC!`, matching Mac Excel's behaviour for cells that try to spill from
// inside another spill.
//
// Publishes externs consumed by the central `kLazyDispatch` table in
// `tree_walker.cpp`; see `eval/lazy_impls.h` for the shared `LazyImpl`
// signature and the `eval_node` entry point these impls recurse through.

#ifndef FORMULON_EVAL_LAMBDA_HELPERS_LAZY_H_
#define FORMULON_EVAL_LAMBDA_HELPERS_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `BYROW(array, function)` — applies the 1-arg `function` to each row of
/// `array`. Each call receives a horizontal `1 x cols` sub-array slice; the
/// output is a vertical `rows x 1` array of per-row results. A lambda that
/// returns an array (multi-cell result) for any row surfaces `#CALC!`. An
/// empty input (`rows == 0` or `cols == 0`) surfaces `#CALC!`. Lambda arity
/// other than 1 surfaces `#VALUE!`.
Value eval_byrow_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `BYCOL(array, function)` — symmetric to `BYROW`. Each call receives a
/// vertical `rows x 1` sub-array; the output is a horizontal `1 x cols`
/// array. Same error semantics as `BYROW`.
Value eval_bycol_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `MAP(array1, [array2, ...], function)` — applies the N-arg `function`
/// elementwise across all `N` input arrays. All arrays must share an
/// identical `(rows, cols)` shape; mismatch surfaces `#N/A`. Lambda arity
/// must equal the number of array arguments; mismatch surfaces `#VALUE!`.
/// A lambda that returns an array for any cell surfaces `#CALC!`.
Value eval_map_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

/// `REDUCE(initial_value, array, function)` — left-fold over `array` in
/// row-major order with the 2-arg `function` (accumulator, current). The
/// accumulator is seeded with `initial_value` and threaded through every
/// call; the helper returns the final accumulator. The accumulator is not
/// constrained to a scalar shape — any `Value` (including arrays) flows
/// through. An empty input array returns `initial_value` unchanged. Lambda
/// arity other than 2 surfaces `#VALUE!`. The first error returned by the
/// lambda short-circuits the fold.
Value eval_reduce_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `SCAN(initial_value, array, function)` — like `REDUCE` but emits every
/// intermediate accumulator. The output has the same `(rows, cols)` shape
/// as `array`; cell `(r, c)` is the accumulator value AFTER processing
/// the corresponding input cell. An empty input surfaces `#CALC!` (no spill
/// shape to emit). Lambda arity other than 2 surfaces `#VALUE!`. The first
/// error returned by the lambda short-circuits the scan.
Value eval_scan_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `MAKEARRAY(rows, cols, function)` — builds a `rows x cols` array where
/// each cell `(r, c)` is the result of calling the 2-arg `function` with
/// the 1-based row and column indices. `rows` and `cols` are coerced to
/// integers; values below 1 surface `#NUM!`, and shapes that exceed
/// Excel's worksheet grid (`rows > 1048576` or `cols > 16384`) also
/// surface `#NUM!`. A lambda that returns an array for any cell surfaces
/// `#CALC!`. Lambda arity other than 2 surfaces `#VALUE!`.
Value eval_makearray_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LAMBDA_HELPERS_LAZY_H_
