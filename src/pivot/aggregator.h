//
// Pivot aggregation primitives.
//
// Each aggregator ignores `Blank`. Errors propagate through the
// arithmetic aggregations: the first error in the input dominates the
// output, matching `SUM(#DIV/0!, 1) -> #DIV/0!`. `Count` and
// `CountNumbers` are the exception -- they classify cells instead of
// coercing values, so an error is counted (COUNTA) or passed over
// (COUNT) but never becomes the result, which is what Excel reports.
// Booleans and text are non-numeric for every aggregation; for `Count` they
// are non-blank so they count, which matches Excel's COUNTA.

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

class PivotTable;
struct PivotResult;

/// One borrowed run of record indices. `cumulative_end` is the exclusive
/// end position of this run in the concatenated view. Spans point into the
/// cache's bucket / surviving-index storage and never own their data.
struct RecordIndexSpan {
  const std::size_t* data = nullptr;
  std::size_t count = 0;
  std::size_t cumulative_end = 0;
};

/// Borrowed record-index sequence used by record aggregations. A direct
/// vector is kept as a pointer/count pair so single-bucket, value-sort, and
/// grand-total paths do not allocate a one-span array. Multi-bucket callers
/// provide an already-built span array; the view only borrows it for the
/// duration of the synchronous aggregation call.
class RecordIndexView {
 public:
  RecordIndexView() = default;
  explicit RecordIndexView(const std::vector<std::size_t>& records)
      : direct_data_(records.data()), direct_count_(records.size()), total_count_(records.size()) {}
  explicit RecordIndexView(std::vector<std::size_t>&&) = delete;

  static RecordIndexView from_spans(const RecordIndexSpan* spans, std::size_t span_count, std::size_t total_count) {
    RecordIndexView view;
    view.spans_ = spans;
    view.span_count_ = span_count;
    view.total_count_ = total_count;
    return view;
  }

  std::size_t size() const { return total_count_; }
  bool empty() const { return total_count_ == 0; }

  /// Returns the source record index at the view position. Sequential reads
  /// stay in the cached span; a non-sequential read searches the cumulative
  /// ends and resets that cache. Callers only pass positions below `size()`.
  std::size_t record_at(std::size_t position) const;

 private:
  const std::size_t* direct_data_ = nullptr;
  std::size_t direct_count_ = 0;
  const RecordIndexSpan* spans_ = nullptr;
  std::size_t span_count_ = 0;
  std::size_t total_count_ = 0;

  mutable std::size_t cached_span_ = 0;
  mutable std::size_t cached_begin_ = 0;
  mutable std::size_t cached_end_ = 0;
};

/// Applies the named aggregation to `values`. The dispatch handles
/// blank-skip, error-propagation (arithmetic aggregations only, see
/// above), and the Excel-specific empty-set semantics (e.g. MAX over an
/// all-text group returns 0).
Value apply_aggregation(Aggregation agg, const std::vector<Value>& values);

/// Applies an aggregation directly to cache records addressed by `records`.
/// The view is borrowed and the returned value is fully materialized.
Value apply_record_aggregation(Aggregation agg, const PivotCache& cache, const RecordIndexView& records,
                               std::uint32_t field_index);
Value apply_record_aggregation(Aggregation agg, const PivotCache& cache, const std::vector<std::size_t>& records,
                               std::uint32_t field_index);

/// Numeric coercion of a single aggregate cell. Returns the underlying
/// double when `v` is a `Number`; booleans and text are non-numeric in a
/// pivot, so `nullopt` for them and everything else.
std::optional<double> numeric_aggregate_value(const Value& v);

/// Builds a borrowed concatenation of `buckets[row][col]` in column-major
/// order (column leaves outer, row leaves inner). `spans` owns only the
/// span descriptors and must outlive the returned view and every aggregation
/// using it. Out-of-range leaves and empty buckets are skipped.
RecordIndexView make_leaf_set_record_view(const RecordBuckets& buckets, const std::vector<std::size_t>& row_leaves,
                                          const std::vector<std::size_t>& col_leaves,
                                          std::vector<RecordIndexSpan>& spans);

/// Single-axis variants avoid constructing a temporary one-element vector
/// for row-leaf and column-leaf subtotal paths.
RecordIndexView make_leaf_set_record_view(const RecordBuckets& buckets, std::size_t row_leaf,
                                          const std::vector<std::size_t>& col_leaves,
                                          std::vector<RecordIndexSpan>& spans);
RecordIndexView make_leaf_set_record_view(const RecordBuckets& buckets, const std::vector<std::size_t>& row_leaves,
                                          std::size_t col_leaf, std::vector<RecordIndexSpan>& spans);

/// Rebuilds every total of a Sum data field from the rendered leaf cells in
/// `result.values`, the way Excel does: a total with one axis on a leaf is a
/// running cancellation-snapped sum of the leaf cells beneath it, a total on
/// both axes a plain running sum of the row leaves' one-axis totals. Totals
/// of other aggregations, and custom subtotal functions other than Sum, are
/// left as they are. The leaf sets are in `result.values` position space.
void recompute_sum_totals_from_cells(const PivotTable& table, PivotResult& result,
                                     const std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                                     const std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets);

/// Aggregates cache records while preserving the sparse-intersection rule:
/// no record positions or only blank values means a blank result; otherwise
/// normal aggregation identities and error semantics apply.
Value aggregate_records_or_blank(Aggregation aggregation, const PivotCache& cache, const RecordIndexView& records,
                                 std::uint32_t field_index, PivotResult& result);
Value aggregate_records_or_blank(Aggregation aggregation, const PivotCache& cache,
                                 const std::vector<std::size_t>& records, std::uint32_t field_index,
                                 PivotResult& result);

/// A total across both axes: every row leaf's one-axis state over `col_leaves`
/// is combined raw (no snap, no flush) top to bottom. Sum keeps its record
/// chain, since its totals are rebuilt from the rendered cells afterwards.
Value aggregate_row_partials_or_blank(Aggregation aggregation, const PivotCache& cache, const RecordBuckets& buckets,
                                      const std::vector<std::size_t>& row_leaves,
                                      const std::vector<std::size_t>& col_leaves, std::uint32_t field_index,
                                      PivotResult& result);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_AGGREGATOR_H_
