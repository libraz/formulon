//
// Regularized incomplete beta ratio and the beta-family distribution
// kernels built on it, evaluated operation by operation (no fused
// multiply-add, subnormal products flushed) so results reproduce Excel's
// bit for bit. Domain violations and non-convergence yield NaN; callers
// translate it into `#NUM!`.

#ifndef FORMULON_EVAL_STATS_INCOMPLETE_BETA_H_
#define FORMULON_EVAL_STATS_INCOMPLETE_BETA_H_

namespace formulon {
namespace eval {
namespace stats {

/// `log(B(a, b))` for `a, b > 0`.
double beta_log_excel(double a, double b) noexcept;

/// Beta density kernel `x^(a-1) (1-x)^(b-1) / B(a, b)` on `(0, 1)`; 0 at the ends.
double beta_density_excel(double a, double b, double x) noexcept;

/// `I_x(a, b)` (`upper == false`) or `1 - I_x(a, b)` (`upper == true`) for
/// `0 <= x <= 1`. `y` is `1 - x` when the caller holds it more accurately
/// than the subtraction; pass a negative value to derive it.
double beta_ratio_excel(double a, double b, double x, bool upper, double y = -1.0) noexcept;

/// BETA.DIST on the support `[lo, hi]`.
double beta_dist_cdf_excel(double x, double a, double b, double lo, double hi) noexcept;
double beta_dist_pdf_excel(double x, double a, double b, double lo, double hi) noexcept;

/// `P(X <= k)` of a binomial; `k` and `n` are truncated.
double binom_cdf_excel(double k, double n, double p) noexcept;

/// BINOM.DIST.RANGE: `P(s1 <= X <= s2)`; pass `s2 == s1` for the point mass.
/// Requires `0 <= p <= 1` and `0 <= s1 <= s2 <= n`.
double binom_range_excel(double n, double p, double s1, double s2) noexcept;

/// Negative binomial pmf / cdf at `k` failures before `r` successes; both truncated.
double negbinom_pmf_excel(double k, double r, double p) noexcept;
double negbinom_cdf_excel(double k, double r, double p) noexcept;

/// F distribution cdf (`upper == false`) or right tail; degrees of freedom truncated.
double f_dist_cdf_excel(double x, double d1, double d2, bool upper) noexcept;
double f_dist_pdf_excel(double x, double d1, double d2) noexcept;

/// Student t density, left tail (`upper == false`), right tail, and two-tailed
/// probability for `x >= 0`; `df` truncated.
double t_dist_pdf_excel(double x, double df) noexcept;
double t_dist_tail_excel(double x, double df, bool upper) noexcept;
double t_dist_2t_excel(double x, double df) noexcept;

/// Hypergeometric pmf / cdf: `x` successes in a sample of `n` from a
/// population of `n_pop` holding `m` successes; all arguments truncated.
double hypgeom_pmf_excel(double x, double n, double m, double n_pop) noexcept;
double hypgeom_cdf_excel(double x, double n, double m, double n_pop) noexcept;

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_INCOMPLETE_BETA_H_
