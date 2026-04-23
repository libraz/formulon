// Copyright 2026 libraz. Licensed under the MIT License.
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

#include <cstdint>
#include <vector>

#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

// Mathematical constant pi, used to normalise the standard-normal PDF.
// Matches `std::acos(-1.0)` on any IEEE-754 system.
inline constexpr double kStatsPi = 3.14159265358979323846;

// Extracts the numeric values from `args[0..count-1]`. Non-Number values
// (text / bool / blank after range expansion) are silently skipped. Errors
// never reach this helper because the dispatcher short-circuits with
// `propagate_errors = true`.
std::vector<double> collect_numerics(const Value* args, std::uint32_t count);

// "A"-family value collector for AVERAGEA / MAXA / MINA / VARA / VARPA /
// STDEVA / STDEVPA. For these functions Excel evaluates text as 0, Bool as
// 0 / 1, and still skips Blank cells. The dispatcher has already performed
// the range-sourced transformation (Bool -> 0/1, Text -> 0, Blank dropped)
// when `range_filter_a_coerce = true`; this helper applies the same rules
// to direct scalar arguments so the full behaviour matches regardless of
// provenance.
std::vector<double> collect_a(const Value* args, std::uint32_t count);

// (mean, sum_of_squared_deviations) pair returned by `compute_mean_ss`.
struct MeanSS {
  double mean;
  double ss;  // Sum of squared deviations from the mean.
};

// Helper: compute `(mean, sum_of_squared_deviations)` over a numeric slice.
// Empty input returns `{0, 0}` which the callers treat as a DIV/0! case.
MeanSS compute_mean_ss(const std::vector<double>& xs);

// Arithmetic-mean helper used by DEVSQ / AVEDEV / SKEW / KURT.
double mean_of(const std::vector<double>& xs) noexcept;

// Helper: read the trailing scalar `k` argument for LARGE / SMALL /
// PERCENTILE / QUARTILE. Returns the raw coerced double so each caller
// applies its own range / truncation rules.
Expected<double, ErrorCode> read_kth_arg(const Value& v);

// Inverse standard-normal CDF. Uses Peter Acklam's rational
// approximation for the initial guess (good to ~1e-6 in practice across
// the unit interval) and then runs Halley-method refinement steps to
// bring the accuracy up to ~1e-14, comfortably inside Excel's reported
// precision. Callers must guarantee `0 < p < 1`; `p <= 0` or `p >= 1`
// should surface `#NUM!` before calling in.
double InverseStandardNormal(double p);

// Probability mass function of Binomial(n, p) at k. Shared between
// BINOM.DIST (in `stats_distributions.cpp`) and BINOM.INV /
// BINOM.DIST.RANGE (in `stats_distributions_misc.cpp`).
double BinomPmf(double k, double n, double prob);

// Newton-Raphson inverter for Student's t CDF. Shared between T.INV /
// T.INV.2T (in `stats_distributions.cpp`) and CONFIDENCE.T (in
// `stats_distributions_misc.cpp`). Assumes `0 < p < 1` and `df >= 1`.
double TInvCore(double p, double df) noexcept;

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
