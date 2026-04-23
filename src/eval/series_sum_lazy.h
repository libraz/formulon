// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impl for `SERIESSUM(x, n, m, coefficients)`. The final argument
// is an array of coefficients, so this function rides the lazy-dispatch
// seam to preserve each argument's AST shape: the dispatcher's eager
// path would flatten a `RangeOp` or `ArrayLiteral` coefficient list into
// a single `Value` and erase the boundary between the three scalar
// arguments and the array tail.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and
// `eval_node` entry point this impl recurses through.

#ifndef FORMULON_EVAL_SERIES_SUM_LAZY_H_
#define FORMULON_EVAL_SERIES_SUM_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `SERIESSUM(x, n, m, coefficients)` - evaluates the power series
/// `Σᵢ aᵢ · x^(n + (i-1)·m)` where `aᵢ` is the i-th coefficient
/// (1-based). The first three arguments are scalars (`x`, initial power
/// `n`, step `m`); the fourth argument may be a `Ref`, a `RangeOp`, or
/// an inline `ArrayLiteral`. Non-numeric scalar arguments return
/// `#VALUE!`; non-numeric cells inside the coefficient array are
/// skipped (with their corresponding power still advancing, to preserve
/// alignment with the 1-based term index Excel uses). An error in any
/// argument propagates in left-to-right order; an empty coefficient
/// list returns `#VALUE!`.
Value eval_series_sum_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SERIES_SUM_LAZY_H_
