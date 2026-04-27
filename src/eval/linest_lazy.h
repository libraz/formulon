// Copyright 2026 libraz. Licensed under the MIT License.
//
// LINEST — multivariate ordinary-least-squares regression.
//
// `LINEST(known_y, [known_x], [const]=TRUE, [stats]=FALSE)` fits the
// linear model `y = b_0 + b_1*x_1 + ... + b_k*x_k` via the normal
// equations `(X^T X) β = X^T y`. The system is solved by Gauss-Jordan
// elimination on the augmented matrix `[A | b | I]` so the same pass
// yields both the coefficient vector β and the inverse `A^-1`, the
// latter being needed for the per-coefficient standard errors when
// `stats=TRUE`. The Gauss-Jordan kernel is the same one used by
// MINVERSE; LINEST does not call MINVERSE directly, but the numerical
// recipe is identical.
//
// Output shapes (all returned as a 2-D `ArrayValue`):
//   stats=FALSE -> 1 x (k+1) row `[b_k, b_{k-1}, ..., b_1, b_0]`
//   stats=TRUE  -> 5 x (k+1) matrix:
//     row 1: coefficients in the same right-to-left order as above;
//     row 2: per-coefficient standard errors;
//     row 3: [r^2, se_y, #N/A, ..., #N/A];
//     row 4: [F, df_resid, #N/A, ..., #N/A];
//     row 5: [ss_reg, ss_resid, #N/A, ..., #N/A].
// When `const=FALSE` the intercept slot in row 1 is `0` and its SE in
// row 2 is `#N/A`, matching Mac Excel's output.
//
// See `eval/regression_lazy.h` for the pairwise regression family
// (CORREL / SLOPE / INTERCEPT / RSQ / STEYX / FORECAST.LINEAR) — those
// are scalar reductions over a single predictor and share the same
// `(known_y, known_x)` argument convention but not the matrix machinery.
// See `eval/matrix_ops_lazy.h` for the underlying linear-algebra
// primitives (Gauss-Jordan with partial pivoting, strict numeric
// matrix coercion).

#ifndef FORMULON_EVAL_LINEST_LAZY_H_
#define FORMULON_EVAL_LINEST_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `LINEST(known_y, [known_x], [const]=TRUE, [stats]=FALSE)` —
/// least-squares regression. Returns a 1 x (k+1) row of coefficients
/// when `stats=FALSE`, or a 5 x (k+1) statistics matrix when
/// `stats=TRUE`. See the file-level comment for the full output layout.
///
/// Errors:
///   - any non-numeric (Text / Blank) cell in `known_y` or `known_x`
///     surfaces as `#VALUE!` (matrix-strict coercion, matches MMULT);
///   - error cells propagate verbatim in left-to-right scan order
///     (`known_y` first);
///   - shape mismatch between `known_y` and `known_x` -> `#REF!`;
///   - singular `X^T X` (collinear predictors or insufficient data)
///     -> `#NUM!`;
///   - `const=TRUE` and `m <= k` (or `const=FALSE` and `m < k`) ->
///     `#NUM!` because the residual degrees of freedom would be
///     non-positive.
Value eval_linest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LINEST_LAZY_H_
