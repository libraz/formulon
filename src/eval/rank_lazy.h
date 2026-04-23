// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the rank / percentile-rank family:
// RANK (legacy), RANK.EQ, RANK.AVG, PERCENTRANK (legacy),
// PERCENTRANK.INC, and PERCENTRANK.EXC.
//
// Each of these six functions shares the shape
// `fn(scalar_or_range, scalar_or_range, [scalar])` in which exactly one
// argument is an array / Ref / RangeOp and the others are scalars. The
// eager dispatch path flattens any `accepts_ranges` argument into a flat
// `Value[]` before the impl runs, which would erase the boundary between
// the flattened range cells and the trailing scalar. We ride the lazy
// dispatch seam so each impl can resolve the array argument via
// `eval::resolve_range_arg` (or walk an inline `{...}` ArrayLiteral) and
// evaluate the scalar arguments separately.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` entry point these impls recurse through, and
// `eval/regression_lazy.h` for the sibling family this one is modelled
// on.

#ifndef FORMULON_EVAL_RANK_LAZY_H_
#define FORMULON_EVAL_RANK_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `RANK.EQ(number, ref, [order])` - rank of `number` within `ref`.
/// Legacy `RANK` shares this impl; Excel keeps both names. `order = 0`
/// (or omitted) sorts descending (largest gets rank 1); any nonzero
/// `order` sorts ascending. Ties share the best (lowest-numbered) rank.
/// `number` must be numeric (Bool / Text -> `#VALUE!`); `ref` must be a
/// range, Ref, or ArrayLiteral. Errors in `ref` propagate. If `number`
/// is not present in `ref` after filtering to numeric cells, the result
/// is `#N/A`.
Value eval_rank_eq_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// `RANK.AVG(number, ref, [order])` - like RANK.EQ, but ties are
/// averaged. For a tie of size `k` starting at position `p` (1-based),
/// every tied value receives rank `p + (k - 1) / 2`.
Value eval_rank_avg_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

/// `PERCENTRANK.INC(array, x, [significance])` - inclusive percentile
/// rank of `x` among the sorted values of `array`. Legacy `PERCENTRANK`
/// shares this impl. The default `significance` is 3; values < 1 yield
/// `#NUM!`; non-integer `significance` is truncated toward zero. An
/// array with fewer than two numeric cells yields `#N/A`. `x` outside
/// `[min, max]` yields `#N/A`.
Value eval_percentrank_inc_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx);

/// `PERCENTRANK.EXC(array, x, [significance])` - exclusive percentile
/// rank. Uses the `(k + 1) / (N + 1)` formula rather than
/// `k / (N - 1)`. `x` outside `[min, max]` yields `#N/A`. Same
/// significance / truncation rules as PERCENTRANK.INC.
Value eval_percentrank_exc_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RANK_LAZY_H_
