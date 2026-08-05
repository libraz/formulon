//
// Numeric aggregation kernels shared by SUBTOTAL (`builtins/subtotal.cpp`),
// AGGREGATE (`aggregate_lazy.cpp`), and the QUARTILE / PERCENTILE family
// (`builtins/stats.cpp`).
//
// Each kernel consumes a `std::vector<double>` of *already-filtered* numeric
// values: callers are responsible for dropping non-numeric inputs (Text,
// Blank, Bool when applicable) and for short-circuiting on Error cells.
// What the kernels add is the post-collection arithmetic + Excel-visible
// error-code surfacing (empty-input -> #DIV/0!, non-finite intermediate ->
// #NUM!, k out of range -> #NUM!, etc.).
//
// Algorithm note: `run_variance` and `run_stdev` deliberately use the
// two-pass mean / sum-of-squared-deviations formulation rather than
// Welford's online recurrence. Both SUBTOTAL and AGGREGATE were already
// two-pass before the consolidation, so keeping that algorithm gives
// bit-identical results to the pre-refactor implementations across the
// oracle corpus. The same applies to PERCENTILE.INC / .EXC, which match
// the position formulas Mac Excel 365 reports.

#ifndef FORMULON_EVAL_AGGREGATE_KERNELS_H_
#define FORMULON_EVAL_AGGREGATE_KERNELS_H_

#include <vector>

#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace aggregate_kernels {

/// Sum of all elements. Returns `#NUM!` when the running total goes
/// non-finite (overflow). Empty input returns 0.
Expected<double, ErrorCode> run_sum(const std::vector<double>& xs);

/// Product of all elements. Returns `#NUM!` on non-finite. Empty input
/// returns 0 (matches the Excel convention that aggregator family uses
/// for SUBTOTAL/AGGREGATE; not the mathematical identity 1).
Expected<double, ErrorCode> run_product(const std::vector<double>& xs);

/// Arithmetic mean. Empty input returns `#DIV/0!`. Non-finite mean -> `#NUM!`.
Expected<double, ErrorCode> run_average(const std::vector<double>& xs);

/// Maximum / minimum. Empty input returns 0 (matches Excel's SUBTOTAL /
/// AGGREGATE behaviour, NOT MAX / MIN which short-circuit on empty
/// numerics differently). Non-finite element -> `#NUM!`.
Expected<double, ErrorCode> run_max(const std::vector<double>& xs);
Expected<double, ErrorCode> run_min(const std::vector<double>& xs);

/// Variance. `sample = true` selects the sample variance (denominator
/// n-1, requires at least 2 elements); `sample = false` selects the
/// population variance (denominator n, requires at least 1 element).
/// Insufficient elements -> `#DIV/0!`; non-finite intermediate ->
/// `#NUM!`. Two-pass algorithm: `mean = sum/n`, then
/// `var = sum((x-mean)^2) / denom`.
Expected<double, ErrorCode> run_variance(const std::vector<double>& xs, bool sample);

/// Standard deviation, `sqrt(run_variance(xs, sample))`. Propagates the
/// variance's `#DIV/0!` / `#NUM!`. A negative variance from floating
/// rounding becomes `#NUM!`.
Expected<double, ErrorCode> run_stdev(const std::vector<double>& xs, bool sample);

/// PERCENTILE.INC at fractional rank `k` in [0, 1]. `xs_sorted` must be
/// non-empty and sorted ascending. Out-of-range `k` -> `#NUM!`; non-
/// finite interpolated result -> `#NUM!`. Implements Excel's position
/// formula `pos = 1 + k*(n-1)` (1-based) with linear interpolation
/// between the two neighbouring elements.
Expected<double, ErrorCode> percentile_sorted_inc(const std::vector<double>& xs_sorted, double k);

/// PERCENTILE.EXC at fractional rank `k` in (0, 1). `xs_sorted` must be
/// non-empty and sorted ascending. `k` outside the boundary
/// `[1/(n+1), n/(n+1)]` -> `#NUM!`; non-finite interpolated result ->
/// `#NUM!`. Implements Excel's position formula `pos = k*(n+1)`
/// (1-based) with linear interpolation between the two neighbouring
/// elements.
Expected<double, ErrorCode> percentile_sorted_exc(const std::vector<double>& xs_sorted, double k);

/// MODE.SNGL kernel: the value tied for the highest frequency that
/// appears *first* in the input order. Excel exact-double equality is
/// used for tie-breaking. Returns `#N/A` when the slice is empty or no
/// value repeats. `xs` is consumed in its original (unsorted) order so
/// the first-occurrence rule is observable; callers must NOT pre-sort.
/// Shared between MODE / MODE.SNGL (`stats/stats_order.cpp`) and
/// AGGREGATE function 13 (`aggregate_lazy.cpp`) so the two cannot
/// diverge on the tie-break rule.
Expected<double, ErrorCode> mode_first_occurrence(const std::vector<double>& xs);

}  // namespace aggregate_kernels
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_AGGREGATE_KERNELS_H_
