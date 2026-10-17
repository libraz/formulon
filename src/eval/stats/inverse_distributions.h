//
// Inverse incomplete gamma and beta cores and the quantile functions built
// on them, evaluated operation by operation (no fused multiply-add,
// subnormal products flushed) so results reproduce Excel's bit for bit.
// Domain violations and non-convergence yield NaN; callers translate it
// into `#NUM!`.

#ifndef FORMULON_EVAL_STATS_INVERSE_DISTRIBUTIONS_H_
#define FORMULON_EVAL_STATS_INVERSE_DISTRIBUTIONS_H_

namespace formulon {
namespace eval {
namespace stats {

/// The `x` with `P(a, x) == p` (`upper == false`) or `Q(a, x) == p`
/// (`upper == true`), by Halley iteration on `gamma_ratio_excel`.
double gamma_inverse_excel(double a, double p, bool upper) noexcept;

/// GAMMA.INV: `b` times the unit-scale quantile.
double gamma_inv_excel(double p, double a, double b) noexcept;

/// CHISQ.INV (`upper == false`) and CHISQ.INV.RT; `df` truncated.
double chisq_inv_excel(double p, double df, bool upper) noexcept;

/// The `x` in `[0, 1]` with `I_x(a, b) == p` (`upper == false`) or
/// `1 - I_x(a, b) == p`, by damped Newton steps on `beta_ratio_excel`. When
/// `complement` is given it receives `1 - x` as the solver holds it.
double beta_inverse_excel(double a, double b, double p, bool upper, double* complement = nullptr) noexcept;

/// BETA.INV on the support `[lo, hi]`.
double beta_inv_excel(double p, double a, double b, double lo, double hi) noexcept;

/// T.INV and T.INV.2T; `df` truncated.
double t_inv_excel(double p, double df) noexcept;
double t_inv_2t_excel(double p, double df) noexcept;

/// F.INV (`upper == false`) and F.INV.RT; degrees of freedom truncated.
double f_inv_excel(double p, double d1, double d2, bool upper) noexcept;

/// BINOM.INV: the smallest `k` with `P(X <= k) >= alpha`, by bisection on
/// the cdf; `n` truncated.
double binom_inv_excel(double n, double p, double alpha) noexcept;

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_INVERSE_DISTRIBUTIONS_H_
