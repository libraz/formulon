//
// Special mathematical functions used by Excel's statistical distribution
// family. The API returns plain IEEE-754 doubles (NaN on domain error) so
// callers can share the helpers from any number of distribution wrappers
// without re-coercing an `Expected<double, ErrorCode>`; the call-site
// translates NaN into the Excel-visible error where the semantics are known.

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

/// Excel's ERF / ERF.PRECISE of one argument: odd; a rational in `x^2` below
/// 0.5 and `1 - erfc` beyond, evaluated operation by operation as Excel does.
/// NaN propagates.
double erf_excel(double x) noexcept;

/// Excel's ERFC / ERFC.PRECISE: `1 - erf` below 0.5, a rational in `x` times
/// `exp(-x^2)` beyond (0 past 100), with `erfc(-x) = 2 - erfc(x)`. NaN
/// propagates.
double erfc_excel(double x) noexcept;

/// Standard normal CDF as Excel's NORM.S.DIST(z, TRUE) evaluates it:
/// `erfc_excel(z * -0.7071067811865476) * 0.5`.
double normal_cdf(double z) noexcept;

/// Inverse error function for `|y| < 1`: rationals in `y^2` and in
/// `sqrt(-log(1 - |y|))`, with a tail form beyond 0.9375.
double erfinv_excel(double y) noexcept;

/// Standard normal quantile as Excel's NORM.S.INV evaluates it: tail
/// rationals in `sqrt(-ln q)` for `p < 0.075` or `p > 0.9375`, `erfinv(2p - 1)
/// * sqrt(2)` between. Requires `0 < p < 1`.
double normal_quantile(double p) noexcept;

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_SPECIAL_FUNCTIONS_H_
