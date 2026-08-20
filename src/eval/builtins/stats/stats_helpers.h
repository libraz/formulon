//
// Internal header -- do not include outside `src/eval/builtins/stats*`.
//
// Shared numeric-collection helpers and forward declarations of the
// Value-returning distribution builtins that live in the sibling
// `stats_distributions.cpp` translation unit. Keeping the declarations
// here (rather than duplicating extern statements across TUs) lets
// `stats.cpp`'s `register_stats_builtins` take the address of every
// distribution impl without also having to know their bodies.

#ifndef FORMULON_EVAL_BUILTINS_STATS_STATS_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_STATS_STATS_HELPERS_H_

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/builtins/numeric_helpers.h"
#include "eval/coerce.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

// Mathematical constant pi, used to normalise the standard-normal PDF.
// Alias to the shared constant in `eval/builtins/numeric_helpers.h` so
// every TU sees the same bit pattern.
inline constexpr double kStatsPi = builtins_detail::kPi;

// Extracts the numeric values from `args[0..count-1]`. Non-Number values
// (text / bool / blank after range expansion) are silently skipped. Errors
// never reach this helper because the dispatcher short-circuits with
// `propagate_errors = true`.
std::vector<double> collect_numerics(const Value* args, std::uint32_t count);

// "A"-family value collector for AVERAGEA / MAXA / MINA / VARA / VARPA /
// STDEVA / STDEVPA. The dispatcher has already transformed range-sourced
// values (Bool -> 0/1, Text -> 0, Blank dropped) via `range_filter_a_coerce`;
// this helper handles direct scalar arguments. Direct Bool is coerced to
// 0/1, direct Blank to 0, and direct Text is strictly coerced via
// `coerce_to_number` -- numeric-looking text becomes its numeric value and
// non-numeric text surfaces as `#VALUE!`. Errors never reach this helper
// because the dispatcher short-circuits with `propagate_errors = true`.
Expected<std::vector<double>, ErrorCode> collect_a(const Value* args, std::uint32_t count);

// Direct-scalar-aware collector used by SMALL / LARGE. Range-sourced cells
// that are non-Number have already been dropped by the dispatcher's
// `range_filter_numeric_only` filter, so this helper only sees Number kinds
// plus any direct scalar arguments. Direct Number -> kept; direct Bool ->
// 1.0 / 0.0; direct Text -> strict `coerce_to_number` (propagates
// `#VALUE!` on unparseable text). Direct Blank is dropped silently, which
// is where this helper diverges from the "A"-family rule. Errors never
// reach this helper (dispatcher short-circuits via `propagate_errors`).
Expected<std::vector<double>, ErrorCode> collect_small_large(const Value* args, std::uint32_t count);

// (mean, sum_of_squared_deviations) pair returned by `compute_mean_ss`.
struct MeanSS {
  double mean;
  double ss;  // Sum of squared deviations from the mean.
};

// Aliases to the shared POD struct types in `numeric_helpers.h`. The
// stats namespace keeps the type names re-exposed so existing call
// sites (and the dozens of forward-declared distribution impls below)
// continue to read naturally.
using NumberPair = builtins_detail::NumberPair;
using NumberTriple = builtins_detail::NumberTriple;

// Stats-specific argument reader: unlike the math / financial / dist
// counterparts, this one does NOT reject NaN / Inf at the coercion
// step. Several callers (T.TEST tail handling, CONFIDENCE.NORM size
// guard) want the raw double so they can apply their own range checks.
inline Expected<double, ErrorCode> read_number_arg(const Value* args, std::uint32_t index) {
  return builtins_detail::read_required_number(args, index, /*check_finite=*/false);
}

inline Expected<NumberPair, ErrorCode> read_number_pair(const Value* args, std::uint32_t first_index,
                                                        std::uint32_t second_index) {
  return builtins_detail::read_number_pair(args, first_index, second_index, /*check_finite=*/false);
}

inline Expected<NumberTriple, ErrorCode> read_number_triple(const Value* args, std::uint32_t first_index,
                                                            std::uint32_t second_index, std::uint32_t third_index) {
  return builtins_detail::read_number_triple(args, first_index, second_index, third_index, /*check_finite=*/false);
}

inline Expected<bool, ErrorCode> read_bool_arg(const Value* args, std::uint32_t index) {
  auto value = coerce_to_bool(args[index]);
  if (!value) {
    return value.error();
  }
  return value.value();
}

inline Value finite_number_result(double value) {
  return builtins_detail::to_finite_value(value);
}

// Helper: compute `(mean, sum_of_squared_deviations)` over a numeric slice.
// Empty input returns `{0, 0}` which the callers treat as a DIV/0! case.
MeanSS compute_mean_ss(const std::vector<double>& xs);

// Result of `centered_and_scaled`: holds the numeric slice + the
// pre-computed mean / scale used by SKEW / SKEW.P / KURT after the
// sample-size and variance-zero guards have already been applied.
struct CenteredScaled {
  std::vector<double> xs;
  double n;  // xs.size() as double
  double mean;
  double scale;  // sqrt(variance), where variance = ss / (n - sample_offset)
};

// Collects the numeric arguments via `collect_numerics`, applies the
// minimum-sample-size guard (`xs.size() < min_n` -> `#DIV/0!`),
// computes the mean and sum-of-squared-deviations via `compute_mean_ss`,
// derives variance with denominator `(n - sample_offset)` (1.0 for sample
// variance, 0.0 for population variance), guards against zero variance
// (also `#DIV/0!`), and returns the {xs, n, mean, sqrt(variance)} tuple.
//
// Both guard failures return `ErrorCode::Div0`. Callers convert that to
// `Value::error(...)` and return.
Expected<CenteredScaled, ErrorCode> centered_and_scaled(const Value* args, std::uint32_t arity, std::size_t min_n,
                                                        double sample_offset);

// Arithmetic-mean helper used by DEVSQ / AVEDEV / SKEW / KURT.
double mean_of(const std::vector<double>& xs) noexcept;

// Helper: read the trailing scalar `k` argument for LARGE / SMALL /
// PERCENTILE / QUARTILE. Returns the raw coerced double so each caller
// applies its own range / truncation rules.
Expected<double, ErrorCode> read_kth_arg(const Value& v);

// Frequency table used by MODE / MODE.SNGL / MODE.MULT to identify the
// values tied for the maximum count. Order of `values` / `counts` follows
// first-occurrence in the input slice. `best_count` is the maximum across
// `counts` (or 0 if `xs` is empty).
struct ModeFrequencies {
  std::vector<double> values;
  std::vector<std::size_t> counts;
  std::size_t best_count = 0;
};

// Builds the first-occurrence-ordered frequency table for MODE / MODE.SNGL
// / MODE.MULT. O(n^2) by design: the input slices are bounded by Excel's
// 255-arg limit for direct scalars and ~1M for ranges, but the duplicate
// detection has to be order-preserving so a hash map would still need a
// parallel insertion-order vector.
ModeFrequencies build_mode_frequencies(const std::vector<double>& xs);

// PERCENTILE.INC / QUARTILE.INC kernel: linear interpolation across an
// already-sorted slice. Callers must pre-sort `xs` and guarantee
// `!xs.empty()` and `k in [0, 1]`. Returns `#NUM!` only on a non-finite
// blend (`Inf - Inf`); the empty-input and out-of-range guards live at
// the callsites.
Value percentile_inc_sorted(const std::vector<double>& xs, double k);

// PERCENTILE.EXC / QUARTILE.EXC kernel: exclusive interpolation across an
// already-sorted slice. Callers must pre-sort `xs` and guarantee
// `!xs.empty()`. The shared kernel applies the open-interval boundary
// check (`pos < 1 || pos >= n`) and returns `#NUM!` at or beyond the
// boundary, matching Mac Excel 365.
Value percentile_exc_sorted(const std::vector<double>& xs, double k);

// Inverse standard-normal CDF. Uses Peter Acklam's rational
// approximation for the initial guess (good to ~1e-6 in practice across
// the unit interval) and then runs Halley-method refinement steps to
// bring the accuracy up to ~1e-14, comfortably inside Excel's reported
// precision. Callers must guarantee `0 < p < 1`; `p <= 0` or `p >= 1`
// should surface `#NUM!` before calling in.
double InverseStandardNormal(double p);

// Ceiling on how many probability-mass terms a discrete cumulative
// builtin may add up one at a time. Every such builtin steps once per
// unit of its argument, so `=BINOM.DIST(1E12, 1E12, 0.5, TRUE)` would run
// for longer than the process will live: the cost tracks the *value* of
// the input rather than its size. One Excel column's worth of terms is
// the same ceiling the depreciation schedules use, and is already far
// past any summation whose accumulated round-off is worth trusting.
//
// A request that needs more terms than this is answered by a closed form
// in constant time where one is accurate enough to stand behind, and
// refused with `#NUM!` where it is not. The Poisson CDF goes through
// `stats::q_gamma`; the binomial family goes through
// `stats::regularized_incomplete_beta`, subject to
// `beta_shapes_are_resolvable` below.
//
// Term-by-term summation is kept below the ceiling so results in the
// ordinary range stay bit-for-bit what they were, and so a single-point
// `BINOM.DIST.RANGE` keeps agreeing exactly with `BinomPmf`.
inline constexpr double kMaxCumulativeTerms = 1048576.0;

// Largest value of `min(a, b)` for which the closed-form beta CDF is
// accurate enough to answer with, rather than refuse.
//
// `stats::regularized_incomplete_beta` evaluates its prefactor in log
// space, where five large terms cancel; the precision left in the result
// therefore falls away as the shapes grow. The loss is worst when the two
// shapes are balanced, so `min(a, b)` -- not their sum -- is what predicts
// it: a skewed pair such as `(5, 1e12)` stays exact, while a balanced one
// does not.
//
// Measured against the symmetric median, whose exact value is 0.5, the
// absolute error is 6.7e-7 at `min(a, b) == 3e8` and 1.17e-6 at 5e8. The
// oracle suite compares this family at a tolerance of 1e-6, so the bound
// sits at the last measured point comfortably inside it. Raising it means
// re-measuring, not rounding up.
inline constexpr double kMaxBalancedBetaShape = 3.0e8;

// Whether the closed-form beta CDF can stand behind a result for the
// shape pair `(a, b)`. Callers that get `false` refuse with `#NUM!`
// rather than return a number outside the tolerance above.
//
// The helper itself independently returns NaN when its recursion cannot
// converge or its prefactor has lost all meaning, which `#NUM!` also
// surfaces; this bound is the tighter, accuracy-driven half of the same
// question.
inline bool beta_shapes_are_resolvable(double a, double b) noexcept {
  return std::fmin(a, b) <= kMaxBalancedBetaShape;
}

// Probability mass function of Binomial(n, p) at k. Shared between
// BINOM.DIST (in `stats_distributions.cpp`) and BINOM.INV /
// BINOM.DIST.RANGE (in `stats_distributions_misc.cpp`).
double BinomPmf(double k, double n, double prob);

// Cumulative distribution function of Binomial(n, p) at k, in closed form
// via the regularized incomplete beta identity
// `P(X <= k) = I_{1-p}(n - k, k + 1)`. Assumes the caller has validated
// `0 <= k <= n` and `0 <= p <= 1`, and has checked the shapes with
// `beta_shapes_are_resolvable`. Constant time regardless of `n`; returns
// NaN when the underlying recursion cannot produce a value, which callers
// surface as `#NUM!`.
double BinomCdf(double k, double n, double prob) noexcept;

// Cumulative distribution function of Poisson(mean) at k, in closed form
// via `P(X <= k) = Q(k + 1, mean)` (the regularized upper incomplete
// gamma). Assumes `k >= 0` and `mean > 0`; the degenerate `mean == 0`
// point mass is handled by the caller. Constant time regardless of `k`,
// and NaN when the underlying series cannot converge, which callers
// surface as `#NUM!`.
double PoissonCdf(double k, double mean) noexcept;

// Newton-Raphson inverter for Student's t CDF. Shared between T.INV /
// T.INV.2T (in `stats_distributions.cpp`) and CONFIDENCE.T (in
// `stats_distributions_misc.cpp`). Assumes `0 < p < 1` and `df >= 1`.
double TInvCore(double p, double df) noexcept;

// Value-returning order-statistic builtins implemented in
// `stats/stats_order.cpp`.
Value Median(const Value* args, std::uint32_t arity, Arena& arena);
Value Mode(const Value* args, std::uint32_t arity, Arena& arena);
Value ModeMult(const Value* args, std::uint32_t arity, Arena& arena);
Value Large(const Value* args, std::uint32_t arity, Arena& arena);
Value Small(const Value* args, std::uint32_t arity, Arena& arena);
Value PercentileInc(const Value* args, std::uint32_t arity, Arena& arena);
Value PercentileExc(const Value* args, std::uint32_t arity, Arena& arena);
Value QuartileInc(const Value* args, std::uint32_t arity, Arena& arena);
Value QuartileExc(const Value* args, std::uint32_t arity, Arena& arena);
Value TrimMean(const Value* args, std::uint32_t arity, Arena& arena);

// Value-returning higher-moment builtins implemented in
// `stats/stats_moments.cpp`.
Value GeoMean(const Value* args, std::uint32_t arity, Arena& arena);
Value HarMean(const Value* args, std::uint32_t arity, Arena& arena);
Value DevSq(const Value* args, std::uint32_t arity, Arena& arena);
Value AveDev(const Value* args, std::uint32_t arity, Arena& arena);
Value Skew(const Value* args, std::uint32_t arity, Arena& arena);
Value SkewP(const Value* args, std::uint32_t arity, Arena& arena);
Value Kurt(const Value* args, std::uint32_t arity, Arena& arena);
Value Standardize(const Value* args, std::uint32_t arity, Arena& arena);

// Value-returning distribution builtins implemented in
// `stats/stats_distributions.cpp`.
Value NormDist(const Value* args, std::uint32_t arity, Arena& arena);
Value NormSDist(const Value* args, std::uint32_t arity, Arena& arena);
Value NormInv(const Value* args, std::uint32_t arity, Arena& arena);
Value NormSInv(const Value* args, std::uint32_t arity, Arena& arena);
Value BinomDist(const Value* args, std::uint32_t arity, Arena& arena);
Value PoissonDist(const Value* args, std::uint32_t arity, Arena& arena);
Value ChisqDist(const Value* args, std::uint32_t arity, Arena& arena);
Value ChisqDistRt(const Value* args, std::uint32_t arity, Arena& arena);
Value ChisqInv(const Value* args, std::uint32_t arity, Arena& arena);
Value ChisqInvRt(const Value* args, std::uint32_t arity, Arena& arena);
Value ExponDist(const Value* args, std::uint32_t arity, Arena& arena);
Value TDist(const Value* args, std::uint32_t arity, Arena& arena);
Value TDist2T(const Value* args, std::uint32_t arity, Arena& arena);
Value TDistRt(const Value* args, std::uint32_t arity, Arena& arena);
Value TInv(const Value* args, std::uint32_t arity, Arena& arena);
Value TInv2T(const Value* args, std::uint32_t arity, Arena& arena);
Value FDist(const Value* args, std::uint32_t arity, Arena& arena);
Value FDistRt(const Value* args, std::uint32_t arity, Arena& arena);
Value FInv(const Value* args, std::uint32_t arity, Arena& arena);
Value FInvRt(const Value* args, std::uint32_t arity, Arena& arena);
Value NormSDistLegacy(const Value* args, std::uint32_t arity, Arena& arena);
Value TDistLegacy(const Value* args, std::uint32_t arity, Arena& arena);
Value ConfidenceNorm(const Value* args, std::uint32_t arity, Arena& arena);
Value ConfidenceT(const Value* args, std::uint32_t arity, Arena& arena);
Value BinomInv(const Value* args, std::uint32_t arity, Arena& arena);
Value Fisher(const Value* args, std::uint32_t arity, Arena& arena);
Value FisherInv(const Value* args, std::uint32_t arity, Arena& arena);
Value Gauss(const Value* args, std::uint32_t arity, Arena& arena);
Value Phi(const Value* args, std::uint32_t arity, Arena& arena);
Value NegBinomDist(const Value* args, std::uint32_t arity, Arena& arena);
Value NegBinomDistLegacy(const Value* args, std::uint32_t arity, Arena& arena);
Value BinomDistRange(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace stats_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_STATS_STATS_HELPERS_H_
