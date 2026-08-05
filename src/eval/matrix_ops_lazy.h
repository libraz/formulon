//
// Lazy impls for the linear-algebra builtins: `MMULT`, `MDETERM`, `MINVERSE`.
// Each accepts one or two array arguments and produces either a scalar
// (MDETERM) or an `ArrayValue` (MMULT, MINVERSE). The dispatcher has to be
// lazy because a `Range` / `RangeOp` argument must keep its 2D shape — the
// eager path would flatten the rectangle into a row-major `Value` vector
// and lose the matrix dimensions.
//
// Hand-rolled implementation (per `CLAUDE.md`'s no-Eigen policy):
//   * MMULT  — triple loop, O(rows(A) * cols(A) * cols(B)).
//   * MDETERM / MINVERSE — Gaussian elimination with partial pivoting
//     (numerically stable for the 2x2 .. 50x50 matrices Mac Excel users
//     actually feed in; xlsx itself is capped well below the LU
//     conditioning crossover).
//
// Externs registered in the central `kLazyDispatch` table in
// `tree_walker.cpp`; see `eval/lazy_impls.h` for the shared signature.

#ifndef FORMULON_EVAL_MATRIX_OPS_LAZY_H_
#define FORMULON_EVAL_MATRIX_OPS_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `MMULT(array1, array2)` — matrix product.
///
/// Both arguments are interpreted as numeric rectangles. Rules (matches
/// Mac Excel 365, ja-JP):
///   * `cols(array1)` must equal `rows(array2)`; otherwise `#VALUE!`.
///   * Every cell in either argument must be a `Number` (or numeric Bool).
///     The first non-numeric / blank / text cell -> `#VALUE!`. Error cells
///     propagate verbatim with the leftmost-wins rule (left-to-right,
///     row-major within each argument).
///   * Result shape is `rows(array1) x cols(array2)`. Each output cell is
///     the dot product of the corresponding row of A and column of B.
///   * NaN / Inf in the accumulated sum surfaces `#NUM!` (matches Mac
///     Excel's overflow behaviour for matrix arithmetic).
Value eval_mmult_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `MDETERM(array)` — determinant of a square matrix.
///
/// Rules (matches Mac Excel 365, ja-JP):
///   * `array` must be square (`rows == cols`); otherwise `#VALUE!`.
///   * Every cell must be a `Number` (or numeric Bool). Non-numeric -> `#VALUE!`.
///     Error cells propagate verbatim (leftmost-wins, row-major).
///   * Result is a scalar `Number`. Singular / near-singular matrices
///     return `0` (Mac Excel rounds the determinant to 0 when the LU
///     pivot underflows below ~1e-13 relative scale; we replicate this
///     by treating an exactly-zero pivot as a singular result).
Value eval_mdeterm_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// `MINVERSE(array)` — inverse of a square matrix.
///
/// Rules (matches Mac Excel 365, ja-JP):
///   * `array` must be square; otherwise `#VALUE!`.
///   * Every cell must be a `Number` (or numeric Bool). Non-numeric ->
///     `#VALUE!`. Error cells propagate verbatim.
///   * Singular matrix (zero pivot during Gauss-Jordan) -> `#NUM!`.
///   * Result shape matches input (`n x n`).
Value eval_minverse_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_MATRIX_OPS_LAZY_H_
