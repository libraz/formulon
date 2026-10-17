//
// Regularized incomplete gamma ratio and the log-space prefactor helpers it
// is built on, evaluated operation by operation (no fused multiply-add,
// subnormal products flushed) so that results reproduce Excel's bit for bit.
// Failure to converge yields NaN; callers translate it into `#NUM!`.

#ifndef FORMULON_EVAL_STATS_INCOMPLETE_GAMMA_H_
#define FORMULON_EVAL_STATS_INCOMPLETE_GAMMA_H_

namespace formulon {
namespace eval {
namespace stats {

/// `log(1 + x) - x`, by series in `x / (x + 2)` on `(-0.99, 1)`.
double log1pmx(double x) noexcept;

/// `log(Gamma(x)) - (x - 0.5) * log(x) + x - log(sqrt(2 pi))` for integer `n`
/// up to 6 from a table, otherwise a series in `1 / n^2` after shifting `n`
/// above 6. Requires `n > 0`.
double stirlerr(double n) noexcept;

/// Deviance term `x * log(x / np) + np - x`, via an atanh series when
/// `|x - np| / (x + np) <= 0.2`.
double bd0(double x, double np) noexcept;

/// `log(Gamma(x))` for `x > 0`: rational minimax on `[0.7, 3]`, shifted
/// products up to 8, Stirling beyond.
double lgamma_excel(double x) noexcept;

/// Poisson density `exp(-lambda) * lambda^k / Gamma(k + 1)` without
/// intermediate overflow (Stirling/deviance form when both arguments exceed 1.1).
double dpois_raw(double lambda, double k) noexcept;

/// Binomial density `C(n, k) p^k (1 - p)^(n - k)` for integer-valued `n`,
/// `k` (truncated), from the Stirling/deviance terms summed with error
/// tracking. Requires `0 < p < 1` unless `k` is 0 or `n`.
double dbinom_raw(double n, double p, double k) noexcept;

/// Regularized incomplete gamma `P(a, x)` (`upper == false`) or `Q(a, x)`
/// (`upper == true`). Requires `a >= 0` and `x >= 0`. NaN when a series or
/// continued fraction does not converge.
double gamma_ratio_excel(double a, double x, bool upper) noexcept;

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_INCOMPLETE_GAMMA_H_
