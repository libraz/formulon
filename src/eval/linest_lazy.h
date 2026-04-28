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
/// Rank-deficient `X` (perfectly collinear predictors) yields a
/// partial fit rather than `#NUM!`: coefficients on redundant columns
/// are 0, the standard-error and inverse-A slots for those columns
/// are 0, and surviving columns absorb the fit. When `const=TRUE`,
/// the intercept is processed first so it absorbs `mean(y)` when all
/// predictors are collinear with the constant column — matching Mac
/// Excel 365's behaviour on `LINEST({1;2;3;4}, {1,2;1,2;1,2;1,2})`
/// (`-> [0, 0, 2.5]`).
///
/// Errors:
///   - any non-numeric (Text / Blank) cell in `known_y` or `known_x`
///     surfaces as `#VALUE!` (matrix-strict coercion, matches MMULT);
///   - error cells propagate verbatim in left-to-right scan order
///     (`known_y` first);
///   - shape mismatch between `known_y` and `known_x` -> `#REF!`;
///   - `const=TRUE` and `m <= k` (or `const=FALSE` and `m < k`) ->
///     `#NUM!` because the residual degrees of freedom would be
///     non-positive.
Value eval_linest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `TREND(known_y, [known_x], [new_x], [const]=TRUE)` — predicted
/// `y` values along the same least-squares fit that LINEST would
/// estimate, evaluated at `new_x`. Internally fits the model the same
/// way (normal equations + Gauss-Jordan), then computes
/// `y_hat = beta . [features, 1]` for each row of `new_x`.
///
/// `new_x` is optional. When omitted, it defaults to `known_x`, which
/// makes TREND return the fitted values at the original observations.
/// `new_x` must carry the same predictor count `k` as `known_x` along
/// the predictor axis; the orthogonal axis carries the new
/// observation count `N`. The output is a 1-D array of length `N`
/// whose orientation matches `known_y` — column when `known_y` is a
/// column, row otherwise.
///
/// Error semantics match LINEST: matrix-strict numeric coercion,
/// left-to-right error propagation (`known_y` -> `known_x` ->
/// `new_x`), `#REF!` on shape mismatch, `#NUM!` on under-determined
/// data. Rank-deficient `X` produces fitted values consistent with
/// the partial-fit recipe documented for LINEST (the prediction
/// `y_hat` collapses to `mean(y)` along the suppressed directions).
Value eval_trend_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `LOGEST(known_y, [known_x], [const]=TRUE, [stats]=FALSE)` — fits
/// the exponential model `y = b * m_1^x_1 * m_2^x_2 * ... * m_k^x_k`
/// by running LINEST on `ln(y)` against `known_x` and then
/// exponentiating the row-1 coefficients in the output. Rows 2 through
/// 5 of the `stats=TRUE` block are left on the linear (log) scale —
/// matching Microsoft's documented behaviour: "the additional
/// statistics are calculated using the linear transformation, so the
/// additional statistics from LOGEST are comparable to those from
/// LINEST".
///
/// When `const=FALSE`, the trailing intercept slot is `1` (i.e.
/// `exp(0)`) rather than `0` — `LOGEST` reports the multiplicative
/// identity for the suppressed `b` term.
///
/// Errors:
///   - any `y <= 0` -> `#NUM!` because `ln` is undefined;
///   - all other diagnostics match LINEST.
Value eval_logest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `GROWTH(known_y, [known_x], [new_x], [const]=TRUE)` — predicted
/// `y` values along the LOGEST exponential fit, evaluated at `new_x`.
/// Internally takes `ln(y)`, fits via LINEST, then computes
/// `exp(beta . [features, 1])` for each row of `new_x`. Argument
/// shape rules and the orientation of the output mirror TREND.
///
/// Errors:
///   - any `y <= 0` -> `#NUM!`;
///   - all other diagnostics match TREND / LINEST.
Value eval_growth_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LINEST_LAZY_H_
