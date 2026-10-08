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

/// Natural log of the gamma function, `ln(|Gamma(x)|)`.
///
/// A thread-safe wrapper around the platform log-gamma: plain `std::lgamma`
/// additionally writes the sign of `Gamma(x)` to the global (non-atomic)
/// `signgam` on glibc/musl/Apple's libm, which is a data race when two
/// worker threads of a parallel recalc both evaluate a lgamma-based builtin
/// (COMBIN, the Poisson/binomial/hypergeometric/gamma/beta/t/F/negative
/// binomial family). This callable never reads that sign, so it always
/// routes through the reentrant `lgamma_r` where the platform provides one.
double log_gamma(double x) noexcept;

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

/// Inverse of the regularized incomplete gamma functions: the `x` with
/// `P(a, x) == prob` (`upper == false`) or `Q(a, x) == prob` (`upper == true`).
///
/// Solved in `ln x` against `ln prob` on whichever tail is smaller, so a
/// probability far below 1 (the lower tail of a large shape, say) keeps full
/// relative accuracy. Returns `NaN`
/// unless `a > 0` and `0 < prob < 1`, or when the iteration fails to converge.
double gamma_quantile(double a, double prob, bool upper) noexcept;

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
/// Like the gamma pair, this returns either a converged value or NaN on
/// every exit path -- a truncated continued fraction is never returned as
/// a result. The iteration budget scales with `a + b`, so the shapes the
/// Excel distribution family reaches converge rather than truncate.
///
/// Accuracy degrades with the shape parameters even when the recursion
/// converges, because the prefactor is evaluated in log space: the
/// absolute error stays below 1e-6 up to roughly `a + b == 1e9` and grows
/// past it. Callers needing 1-bit agreement beyond that magnitude cannot
/// get it from this helper.
///
/// The T / F Excel distribution family routes through this helper; the
/// caller translates NaN into `#NUM!`.
double regularized_incomplete_beta(double a, double b, double x) noexcept;

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_SPECIAL_FUNCTIONS_H_
