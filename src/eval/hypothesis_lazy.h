// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for Excel's hypothesis-test / probability family:
// T.TEST (aliased as TTEST), F.TEST (aliased as FTEST),
// CHISQ.TEST (aliased as CHITEST), Z.TEST (aliased as ZTEST), and PROB.
//
// Each function takes one or two array arguments whose `(rows, cols)`
// shape is observable — CHISQ.TEST and T.TEST's paired mode (`type == 1`)
// and PROB all require array1.shape == array2.shape, while the two-sample
// T.TEST types and F.TEST collect each array's numeric cells
// independently. This is exactly the shape information the eager dispatch
// path would erase by flattening each argument to a single `Value`, so
// the family rides the lazy-dispatch seam — same rationale as the
// regression family in `eval/regression_lazy.h`.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` entry point these impls recurse through. The distribution
// back-ends (regularized incomplete beta, upper-tail gamma, erfc) come
// from `eval/stats/special_functions.h` and `<cmath>`.

#ifndef FORMULON_EVAL_HYPOTHESIS_LAZY_H_
#define FORMULON_EVAL_HYPOTHESIS_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `T.TEST(array1, array2, tails, type)` - Student's t-test p-value.
///
/// `tails` must truncate to 1 or 2; `type` must truncate to 1 (paired),
/// 2 (two-sample equal variance), or 3 (Welch's unequal variance).
/// Otherwise the function returns `#NUM!`. Paired mode requires matching
/// `(rows, cols)`; a shape mismatch surfaces as `#N/A`. Two-sample and
/// Welch modes collect numeric cells from each array independently
/// (non-numeric cells dropped, errors propagate). Degenerate statistics
/// (`s_d == 0`, `s_p == 0`, `df` underflow, or too few samples) surface
/// as `#DIV/0!`. Registered under both `T.TEST` (modern) and `TTEST`
/// (legacy) with identical semantics.
Value eval_t_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `F.TEST(array1, array2)` - two-tailed F-test of equal variances.
///
/// Collects each array's numeric cells independently (no shape check, per
/// Excel). `n1 < 2` or `n2 < 2` -> `#DIV/0!`; either sample variance == 0
/// -> `#DIV/0!`. Computes `F = s1²/s2²` and returns
/// `2 * min(cdf, 1 - cdf)` where `cdf` is the F(df1, df2) CDF at `F`.
/// Registered under both `F.TEST` (modern) and `FTEST` (legacy).
Value eval_f_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `CHISQ.TEST(actual_range, expected_range)` - chi-squared independence
/// p-value.
///
/// Shape MUST match, else `#N/A`. Any error cell propagates; any
/// non-numeric cell on either side -> `#VALUE!` (CHISQ.TEST expects a
/// clean contingency table). Any expected_i <= 0 -> `#NUM!`. For a
/// rectangle `(r, c)` with r > 1 AND c > 1, `df = (r-1)(c-1)`; for a
/// 1-D sequence (r == 1 OR c == 1), `df = n - 1`. `df < 1` -> `#NUM!`.
/// p-value = `q_gamma(df/2, chi²/2)`. Registered under both `CHISQ.TEST`
/// (modern) and `CHITEST` (legacy).
Value eval_chisq_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

/// `Z.TEST(array, x, [sigma])` - one-tailed z-test p-value for the mean.
///
/// `array` is an array argument; `x` is a scalar number. `sigma`, when
/// provided, must be a positive scalar number (else `#NUM!`); when
/// omitted, the sample standard deviation is used (requires n >= 2 and
/// non-zero SD, else `#DIV/0!`). Returns
/// `1 - NORM.S.DIST((mean - x)/(sigma/sqrt(n)), TRUE) =
/// 0.5 * erfc(z / sqrt(2))`. Registered under both `Z.TEST` (modern) and
/// `ZTEST` (legacy).
Value eval_z_test_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `PROB(x_range, prob_range, lower_limit, [upper_limit])` - sum of
/// probabilities in a value range.
///
/// `x_range` and `prob_range` must share a shape (else `#N/A`). Any
/// `prob_i < 0`, `prob_i > 1`, or sum-of-probs != 1 (absolute tolerance
/// 1e-12) -> `#NUM!`. With arity 3, returns the sum of `prob_i` for
/// which `x_i == lower_limit`. With arity 4, returns the sum of `prob_i`
/// for which `lower_limit <= x_i <= upper_limit`; `upper < lower` ->
/// `#NUM!`.
Value eval_prob_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_HYPOTHESIS_LAZY_H_
