// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the pairwise linear-regression family: CORREL,
// COVARIANCE.P, COVARIANCE.S, SLOPE, INTERCEPT, RSQ, and
// FORECAST.LINEAR (aliased as FORECAST). Each takes two parallel arrays
// (plus a leading x-scalar in FORECAST.LINEAR's case) and computes a
// statistic over matched (x, y) pairs.
//
// These ride the lazy-dispatch seam rather than the eager path because
// every array argument may arrive as a bare `Ref`, a `RangeOp`, or an
// inline `{...}` `ArrayLiteral`, and the two arrays must share an
// identical `(rows, cols)` rectangle — a shape-check that the eager
// dispatcher would erase by flattening each arg to a single `Value`.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` entry point these impls recurse through, and
// `eval/shape_ops_lazy.h` for the sibling family (`SUMPRODUCT` etc.)
// that this one is modelled on.

#ifndef FORMULON_EVAL_REGRESSION_LAZY_H_
#define FORMULON_EVAL_REGRESSION_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `CORREL(known_y, known_x)` - Pearson correlation coefficient over
/// matched (x, y) pairs. Non-numeric cells drop the whole pair; errors
/// in any cell propagate. Shape mismatch returns `#N/A`. A degenerate
/// data set (n < 2, variance(x) == 0, or variance(y) == 0) returns
/// `#DIV/0!`.
Value eval_correl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `COVARIANCE.P(known_y, known_x)` - population covariance,
/// `sum((x - mean_x)(y - mean_y)) / n`. Empty pair set -> `#DIV/0!`.
Value eval_covariance_p_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

/// `COVARIANCE.S(known_y, known_x)` - sample covariance,
/// `sum((x - mean_x)(y - mean_y)) / (n - 1)`. Fewer than 2 pairs ->
/// `#DIV/0!`.
Value eval_covariance_s_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

/// `SLOPE(known_y, known_x)` - linear-regression slope,
/// `sum_xy / sum_xx`. `sum_xx == 0` or n < 2 -> `#DIV/0!`.
Value eval_slope_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `INTERCEPT(known_y, known_x)` - linear-regression intercept,
/// `mean_y - slope * mean_x`. Same degeneracy conditions as SLOPE.
Value eval_intercept_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

/// `RSQ(known_y, known_x)` - square of the Pearson correlation. Same
/// degeneracy conditions as CORREL.
Value eval_rsq_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

/// `FORECAST.LINEAR(x, known_y, known_x)` - linear prediction at `x`.
/// Aliased as `FORECAST` (legacy name). Errors on the scalar `x`
/// argument propagate; degeneracy in the regression surfaces as
/// `#DIV/0!`; shape mismatch surfaces as `#N/A`.
Value eval_forecast_linear_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REGRESSION_LAZY_H_
