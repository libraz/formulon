//
// Numeric aggregation kernels shared by SUBTOTAL (`src/eval/builtins/subtotal.cpp`),
// AGGREGATE (`src/eval/aggregate_lazy.cpp`), AVERAGE (`src/eval/builtins/aggregate.cpp`),
// the stats family (`src/eval/builtins/stats.cpp`, `src/eval/builtins/stats/stats_order.cpp`),
// and the pivot aggregator (`src/pivot/aggregator.cpp`).
//
// Each kernel consumes either a `std::vector<double>` of already-filtered
// numeric values or a `NumericInputView`. The view lets a caller retain its
// source values and expose only the numeric cells through a small getter;
// returning `nullopt` skips one input. Callers remain responsible for
// short-circuiting on Error cells before invoking a view kernel. What the
// kernels add is the post-collection arithmetic + Excel-visible error-code
// surfacing (empty-input -> #DIV/0!, non-finite intermediate -> #NUM!, k out
// of range -> #NUM!, etc.).
//
// Algorithm note: `run_variance` and `run_stdev` deliberately use the
// two-pass mean / sum-of-squared-deviations formulation rather than
// Welford's online recurrence. Both SUBTOTAL and AGGREGATE were already
// two-pass before the consolidation, so keeping that algorithm gives
// bit-identical results to the pre-refactor implementations across the
// oracle corpus; `run_average` is the plain sum / count likewise. The same
// applies to PERCENTILE.INC / .EXC, which match the position formulas Mac
// Excel 365 reports.

#ifndef FORMULON_NUMERIC_AGGREGATE_KERNELS_H_
#define FORMULON_NUMERIC_AGGREGATE_KERNELS_H_

#include <cstddef>
#include <optional>
#include <vector>

#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace numeric_aggregate_kernels {

/// Borrowed numeric input. `context` is passed unchanged to `get` for each
/// index in `[0, size)`. A `nullopt` result means that the source element is
/// not part of the numeric aggregate. The view never owns or outlives the
/// pointed-to context; vector overloads below provide the common owning-free
/// adapter for already-filtered slices.
struct NumericInputView {
  const void* context = nullptr;
  std::size_t size = 0;
  std::optional<double> (*get)(const void* context, std::size_t index) = nullptr;
};

/// Sum of all elements. Returns `#NUM!` when the running total goes
/// non-finite (overflow). Empty input returns 0.
Expected<double, ErrorCode> run_sum(const std::vector<double>& xs);

/// Product of all elements. Returns `#NUM!` on non-finite. Empty input
/// returns 0 (matches the Excel convention that aggregator family uses
/// for SUBTOTAL/AGGREGATE; not the mathematical identity 1).
Expected<double, ErrorCode> run_product(const std::vector<double>& xs);

/// Arithmetic mean (plain sum / count). Empty input returns `#DIV/0!`. A
/// non-finite result returns `#NUM!`.
Expected<double, ErrorCode> run_average(const NumericInputView& values);
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
/// `#NUM!`.
Expected<double, ErrorCode> run_variance(const NumericInputView& values, bool sample);
Expected<double, ErrorCode> run_variance(const std::vector<double>& xs, bool sample);

/// Standard deviation: the square root of `run_variance`; a variance that
/// overflows reports `#NUM!`.
Expected<double, ErrorCode> run_stdev(const NumericInputView& values, bool sample);
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

/// MEDIAN kernel: the middle element of `xs` by value, or the mean of the
/// two middle elements when the count is even. `xs` is sorted internally,
/// so callers must not pre-sort and may `std::move` their vector in.
/// Returns `#NUM!` for an empty slice and for a non-finite midpoint.
///
/// Shared between MEDIAN (`src/eval/builtins/stats/stats_order.cpp`) and AGGREGATE function
/// 12 (`src/eval/aggregate_lazy.cpp`). Excel treats the two spellings as
/// interchangeable, so a single definition is what keeps them from
/// reporting different error codes for the same filtered input.
Expected<double, ErrorCode> run_median(std::vector<double> xs);

/// MODE.SNGL kernel: the value tied for the highest frequency that
/// appears *first* in the input order. Excel exact-double equality is
/// used for tie-breaking. Returns `#N/A` when the slice is empty or no
/// value repeats. `xs` is consumed in its original (unsorted) order so
/// the first-occurrence rule is observable; callers must NOT pre-sort.
/// Shared between MODE / MODE.SNGL (`src/eval/builtins/stats/stats_order.cpp`) and
/// AGGREGATE function 13 (`src/eval/aggregate_lazy.cpp`) so the two cannot
/// diverge on the tie-break rule.
Expected<double, ErrorCode> mode_first_occurrence(const std::vector<double>& xs);

/// A run of exactly equal values: the value, the smallest source position
/// holding it, and how many positions do.
struct ValueRun {
  double value;
  std::size_t first;
  std::size_t count;
};

/// Groups `xs` into runs of exactly equal values, in ascending value order.
/// Shared by the mode kernel here and `build_mode_frequencies` in
/// `src/eval/builtins/stats.cpp` so both apply the same first-occurrence grouping.
std::vector<ValueRun> group_equal_values(const std::vector<double>& xs);

}  // namespace numeric_aggregate_kernels
}  // namespace formulon

#endif  // FORMULON_NUMERIC_AGGREGATE_KERNELS_H_
