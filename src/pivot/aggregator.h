//
// Pivot aggregation primitives.
//
// Each aggregator ignores `Blank`. Errors propagate through the
// arithmetic aggregations: the first error in the input dominates the
// output, matching `SUM(#DIV/0!, 1) -> #DIV/0!`. `Count` and
// `CountNumbers` are the exception -- they classify cells instead of
// coercing values, so an error is counted (COUNTA) or passed over
// (COUNT) but never becomes the result, which is what Excel reports.
// Booleans coerce numerically (TRUE=1, FALSE=0) for arithmetic
// aggregations; for `Count`, booleans are non-blank so they count, which
// also matches Excel's COUNTA on a boolean column.
//
// The header also exposes the "gather record values for a leaf or
// leaf-set" helpers used by both the leaf-cell aggregation in the main
// `evaluate()` driver and the subtotal walk.

#ifndef FORMULON_PIVOT_AGGREGATOR_H_
#define FORMULON_PIVOT_AGGREGATOR_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_types.h"
#include "value.h"

namespace formulon::pivot {

/// Per-record bucketing: `[row_leaf][col_leaf]` -> indices into
/// `cache.records()`.
using RecordBuckets = std::vector<std::vector<std::vector<std::size_t>>>;

/// Applies the named aggregation to `values`. The dispatch handles
/// blank-skip, error-propagation (arithmetic aggregations only, see
/// above), and the Excel-specific empty-set semantics (e.g. MAX over an
/// all-text group returns 0).
Value apply_aggregation(Aggregation agg, const std::vector<Value>& values);

/// Numeric coercion of a single aggregate cell. Returns the underlying
/// double when `v` is a `Number` or `Bool`; `nullopt` otherwise.
std::optional<double> numeric_aggregate_value(const Value& v);

/// Appends `cell_value(cache, cache.records()[i], field_index)` for
/// every `i` in `records` to `out`.
void append_record_field_values(const PivotCache& cache, const std::vector<std::size_t>& records,
                                std::uint32_t field_index, std::vector<Value>& out);

/// Appends the per-record field values that live in `buckets[row_leaf][col_leaf]`.
/// Out-of-range `row_leaf` / `col_leaf` indices are silently ignored.
void append_bucket_field_values(const PivotCache& cache, const RecordBuckets& buckets, std::size_t row_leaf,
                                std::size_t col_leaf, std::uint32_t field_index, std::vector<Value>& out);

/// Appends per-record field values across a Cartesian set of row leaves
/// and column leaves; used by the row x col subtotal intersection pass.
void append_leaf_set_field_values(const PivotCache& cache, const RecordBuckets& buckets,
                                  const std::vector<std::size_t>& row_leaves,
                                  const std::vector<std::size_t>& col_leaves, std::uint32_t field_index,
                                  std::vector<Value>& out);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_AGGREGATOR_H_
