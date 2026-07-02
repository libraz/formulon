// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Pivot filter engine.
//
// Two filter passes live here:
//
//   * Pre-aggregation, per-record: `record_passes_manual_filter` runs
//     the manual `items[]` visibility check on every field that declares
//     one, then applies axis-level label / date filters from
//     `active_filters`.
//
//   * Post-aggregation, per-leaf: `score_*_axis` + `build_value_filter_keep`
//     score each axis leaf for the active value filter and produce a
//     boolean keep mask aligned to `result.values`'s DFS pre-order leaf
//     enumeration; `compact_*_axis_values` then collapses the matrix
//     to the surviving leaves.

#ifndef FORMULON_PIVOT_FILTER_ENGINE_H_
#define FORMULON_PIVOT_FILTER_ENGINE_H_

#include <cstddef>
#include <optional>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "value.h"

namespace formulon::pivot {

/// True iff `record` survives the manual `items[]` filter on every
/// field that declares one, AND the axis-level label filters in
/// `active_filters`. Empty `items` lists match all values (Excel default
/// — items[] is only authored when the user has hidden at least one
/// value); axis filters with unresolved field names are skipped.
bool record_passes_manual_filter(const PivotTable& table, const PivotCache& cache, const PivotCacheRecord& record);

/// Per-axis scoring intermediate used by `build_value_filter_keep`.
/// `scores[i]` is the numeric aggregate sum at leaf `i`; `all_blank[i]`
/// is true when leaf `i` had no numeric content (such leaves can be
/// excluded even by `ValueGreaterThan`-style filters).
struct AxisScores {
  std::vector<double> scores;
  std::vector<bool> all_blank;
};

/// Sums leaf scores along the row axis (one entry per row leaf,
/// reducing across the column axis). Used to back row-axis value
/// filters such as `ValueTop10` and `ValueGreaterThan`.
AxisScores score_row_axis(const PivotResult& result, std::size_t row_count, std::size_t col_count);

/// Mirror of `score_row_axis` for the column axis.
AxisScores score_col_axis(const PivotResult& result, std::size_t col_count, std::size_t row_count);

/// Builds a per-leaf keep mask for `f`. Returns `nullopt` for filter
/// shapes that should degrade to a no-op (e.g. unbounded `ValueBetween`).
std::optional<std::vector<bool>> build_value_filter_keep(const PivotFilter& f, const AxisScores& axis);

/// Compacts `result.values` along the row axis according to `keep`.
/// Rows whose mask is false are dropped; surviving rows preserve their
/// relative order. Mirrors what `prune_top_level` does to the row
/// hierarchy.
void compact_row_axis_values(std::vector<std::vector<std::vector<Value>>>& values, const std::vector<bool>& keep);

/// Compacts every row's per-column slice in `result.values` according
/// to `keep`. Mirror of `compact_row_axis_values` for the column axis.
void compact_col_axis_values(std::vector<std::vector<std::vector<Value>>>& values, const std::vector<bool>& keep);

/// Compacts a per-leaf totals vector (`row_leaf_totals` / `col_leaf_totals`,
/// shape `[leaf][data_field]`) according to `keep`, dropping the leaves a
/// value filter pruned so the totals stay index-aligned with
/// `result.values`.
void compact_leaf_totals(std::vector<std::vector<Value>>& totals, const std::vector<bool>& keep);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_FILTER_ENGINE_H_
