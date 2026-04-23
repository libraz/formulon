// Copyright 2026 libraz. Licensed under the MIT License.
//
// Special mathematical functions used by Excel's statistical distribution
// family (CHISQ.*, T.*, F.*, GAMMA.*, BETA.*). The API deliberately returns
// plain IEEE-754 doubles (NaN on domain error) so callers can share the same
// helper from any number of distribution wrappers without re-coercing an
// `Expected<double, ErrorCode>` at each step; the call-site is expected to
// translate NaN into the appropriate Excel-visible error (`#NUM!`) exactly
// where the Excel semantics are known.
//
// Algorithms follow Numerical Recipes §6.2 (regularized incomplete gamma)
// and §6.4 (regularized incomplete beta). Inputs outside the domain of a
// given function return IEEE-754 quiet NaN instead of signalling an error;
// callers check with `std::isnan`. Each function satisfies
// `p_gamma(a, x) + q_gamma(a, x) == 1` and
// `regularized_incomplete_beta(a, b, x) + regularized_incomplete_beta(b, a, 1-x) == 1`
// to within a few ulps when all results are finite.

#ifndef FORMULON_EVAL_STATS_SPECIAL_FUNCTIONS_H_
#define FORMULON_EVAL_STATS_SPECIAL_FUNCTIONS_H_

namespace formulon {
namespace eval {
namespace stats {

/// Regularized lower incomplete gamma function
/// `P(a, x) = γ(a, x) / Γ(a)`.
///
/// Evaluated by the series expansion for `x < a + 1` and via the complement
/// of the continued-fraction form (`1 - q_gamma(a, x)`) otherwise. The two
/// paths agree to within a few ulps at the boundary.
///
/// Returns `NaN` if `a <= 0` or `x < 0`; `0` at `x == 0`.
double p_gamma(double a, double x) noexcept;

/// Regularized upper incomplete gamma function
/// `Q(a, x) = Γ(a, x) / Γ(a) = 1 - P(a, x)`.
///
/// Evaluated directly via Lentz's modified continued-fraction algorithm for
/// `x >= a + 1`, and as `1 - p_gamma` otherwise.
///
/// Returns `NaN` if `a <= 0` or `x < 0`; `1` at `x == 0`.
double q_gamma(double a, double x) noexcept;

/// Regularized incomplete beta function
/// `I_x(a, b) = B(x; a, b) / B(a, b)`, the CDF of the Beta(a, b) law.
///
/// Evaluated via Lentz's modified continued fraction (Numerical Recipes
/// §6.4) on the branch `x < (a + 1) / (a + b + 2)` and via the standard
/// symmetry reflection `1 - I_{1-x}(b, a)` otherwise. The reflection keeps
/// both branches on the fast-convergence side of the split.
///
/// Preconditions: `a > 0`, `b > 0`, `0 <= x <= 1`. Violations return
/// IEEE-754 quiet NaN. At the endpoints returns `0` (x == 0) and `1`
/// (x == 1) exactly.
///
/// The T / F Excel distribution family routes through this helper; the
/// caller translates NaN into `#NUM!`.
double regularized_incomplete_beta(double a, double b, double x) noexcept;

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_SPECIAL_FUNCTIONS_H_
