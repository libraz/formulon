
#include "pivot/value_filter_pass.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <numeric>
#include <optional>
#include <utility>
#include <variant>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/hierarchy_builder.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// Drops the entries of `leaves` whose index is marked false in `keep`,
// preserving survivor order. Mirrors `filter_engine.cpp`'s `keep_marked`
// for the one `std::vector<std::size_t>` case this file needs (see
// `surviving_row_leaves` / `surviving_col_leaves` below); kept local
// rather than exposed from `filter_engine.h` since nothing else needs
// the general template.
void CompactSurvivingLeaves(std::vector<std::size_t>& leaves, const std::vector<bool>& keep) {
  std::vector<std::size_t> kept;
  kept.reserve(leaves.size());
  for (std::size_t i = 0; i < leaves.size(); ++i) {
    if (i < keep.size() && !keep[i]) {
      continue;
    }
    kept.push_back(leaves[i]);
  }
  leaves = std::move(kept);
}

// Maps a subtotal's leaf set (in current, post-filter leaf-position
// space) to `buckets` coordinates via `original_of_position`.
std::vector<std::size_t> ToOriginalLeaves(const std::vector<std::size_t>& leaf_set,
                                          const std::vector<std::size_t>& original_of_position) {
  std::vector<std::size_t> out;
  out.reserve(leaf_set.size());
  for (const std::size_t position : leaf_set) {
    if (position < original_of_position.size()) {
      out.push_back(original_of_position[position]);
    }
  }
  return out;
}

// Recomputes every total `PivotResult` exposes -- grand totals, per-leaf
// totals across the opposite axis, and every row/col subtotal value --
// from `buckets` directly, restricted to the leaves that survived every
// value-axis filter (evaluate()'s step 7). Re-deriving from records
// rather than adjusting the already-computed per-cell aggregates is what
// keeps a non-additive aggregation (Average/Max/Min/StdDev/Var) correct,
// and it is the only way a filtered-out leaf's contribution is actually
// removed rather than merely hidden from the rendered leaf list --
// verified against tests/fixtures/excel/pivot_value_date_filters.xlsx
// and pivot_recurring_period_filter.xlsx, both real Excel-authored
// Top-2 filters whose cached Grand Total cell equals the sum of the two
// surviving rows, not all four source rows.
//
// `surviving_row_leaves[i]` / `surviving_col_leaves[i]` map the current
// (post-filter) leaf position `i` back to its original `buckets` row/col
// index. `row_subtotal_leaf_sets` / `col_subtotal_leaf_sets` are already
// in that same current-position space -- both are recompacted by the
// same `keep` mask, at the same point in the filter loop, as
// `surviving_row_leaves` / `surviving_col_leaves` themselves -- so a
// subtotal's own leaf set maps to `buckets` coordinates through the same
// lookup (`ToOriginalLeaves`).
void ReaggregateTotalsAfterValueFilter(const PivotTable& table, const PivotCache& cache, const RecordBuckets& buckets,
                                       const std::vector<std::size_t>& surviving_row_leaves,
                                       const std::vector<std::size_t>& surviving_col_leaves,
                                       const std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                                       const std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets,
                                       PivotResult& result) {
  const std::size_t data_field_count = table.data_fields().size();
  if (data_field_count == 0) {
    return;
  }

  for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < result.grand_totals.size(); ++df_idx) {
    const PivotDataField& df = table.data_fields()[df_idx];
    std::vector<Value> column;
    append_leaf_set_field_values(cache, buckets, surviving_row_leaves, surviving_col_leaves, df.field_index, column);
    result.grand_totals[df_idx] = aggregate_or_blank(df.aggregation, column, result);
  }
  if (!result.grand_totals.empty()) {
    result.grand_total = result.grand_totals[0];
  }

  for (std::size_t r = 0; r < result.row_leaf_totals.size() && r < surviving_row_leaves.size(); ++r) {
    for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < result.row_leaf_totals[r].size(); ++df_idx) {
      const PivotDataField& df = table.data_fields()[df_idx];
      std::vector<Value> column;
      append_leaf_set_field_values(cache, buckets, {surviving_row_leaves[r]}, surviving_col_leaves, df.field_index,
                                   column);
      result.row_leaf_totals[r][df_idx] = aggregate_or_blank(df.aggregation, column, result);
    }
  }

  for (std::size_t c = 0; c < result.col_leaf_totals.size() && c < surviving_col_leaves.size(); ++c) {
    for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < result.col_leaf_totals[c].size(); ++df_idx) {
      const PivotDataField& df = table.data_fields()[df_idx];
      std::vector<Value> column;
      append_leaf_set_field_values(cache, buckets, surviving_row_leaves, {surviving_col_leaves[c]}, df.field_index,
                                   column);
      result.col_leaf_totals[c][df_idx] = aggregate_or_blank(df.aggregation, column, result);
    }
  }

  for (std::size_t rs = 0; rs < result.row_subtotals.size() && rs < row_subtotal_leaf_sets.size(); ++rs) {
    RowSubtotal& subtotal = result.row_subtotals[rs];
    const std::vector<std::size_t> rows = ToOriginalLeaves(row_subtotal_leaf_sets[rs], surviving_row_leaves);
    for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.values.size(); ++df_idx) {
      const PivotDataField& df = table.data_fields()[df_idx];
      std::vector<Value> column;
      append_leaf_set_field_values(cache, buckets, rows, surviving_col_leaves, df.field_index, column);
      subtotal.values[df_idx] = aggregate_or_blank(subtotal.aggregation.value_or(df.aggregation), column, result);
    }
    for (std::size_t c = 0; c < subtotal.col_values.size() && c < surviving_col_leaves.size(); ++c) {
      for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.col_values[c].size(); ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        std::vector<Value> column;
        append_leaf_set_field_values(cache, buckets, rows, {surviving_col_leaves[c]}, df.field_index, column);
        subtotal.col_values[c][df_idx] =
            aggregate_or_blank(subtotal.aggregation.value_or(df.aggregation), column, result);
      }
    }
    for (std::size_t cs = 0; cs < subtotal.col_subtotal_values.size() && cs < col_subtotal_leaf_sets.size(); ++cs) {
      const std::vector<std::size_t> cols = ToOriginalLeaves(col_subtotal_leaf_sets[cs], surviving_col_leaves);
      const ColSubtotal& col_subtotal = result.col_subtotals[cs];
      for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.col_subtotal_values[cs].size();
           ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        const Aggregation aggregation =
            col_subtotal.aggregation.has_value() ? *col_subtotal.aggregation : df.aggregation;
        std::vector<Value> column;
        append_leaf_set_field_values(cache, buckets, rows, cols, df.field_index, column);
        subtotal.col_subtotal_values[cs][df_idx] = aggregate_or_blank(aggregation, column, result);
      }
    }
  }

  for (std::size_t cs = 0; cs < result.col_subtotals.size() && cs < col_subtotal_leaf_sets.size(); ++cs) {
    ColSubtotal& subtotal = result.col_subtotals[cs];
    const std::vector<std::size_t> cols = ToOriginalLeaves(col_subtotal_leaf_sets[cs], surviving_col_leaves);
    for (std::size_t r = 0; r < subtotal.values.size() && r < surviving_row_leaves.size(); ++r) {
      for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.values[r].size(); ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        const Aggregation aggregation = subtotal.aggregation.has_value() ? *subtotal.aggregation : df.aggregation;
        std::vector<Value> column;
        append_leaf_set_field_values(cache, buckets, {surviving_row_leaves[r]}, cols, df.field_index, column);
        subtotal.values[r][df_idx] = aggregate_or_blank(aggregation, column, result);
      }
    }
    for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.total.size(); ++df_idx) {
      const PivotDataField& df = table.data_fields()[df_idx];
      const Aggregation aggregation = subtotal.aggregation.has_value() ? *subtotal.aggregation : df.aggregation;
      std::vector<Value> column;
      append_leaf_set_field_values(cache, buckets, surviving_row_leaves, cols, df.field_index, column);
      subtotal.total[df_idx] = aggregate_or_blank(aggregation, column, result);
    }
  }
}

// Original-`buckets`-space indices of the row/col leaves still standing,
// maintained in lock-step with the subtotal leaf sets (both recompacted by
// the same `keep` mask, at the same point, by every filter application) so
// the re-aggregation can map any leaf position -- a plain leaf or a
// subtotal's leaf set -- back to `buckets` coordinates.
struct ValueFilterState {
  const PivotTable& table;
  const PivotCache& cache;
  const RecordBuckets& buckets;
  PivotResult& result;
  std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets;
  std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets;
  std::vector<std::size_t> surviving_row_leaves;
  std::vector<std::size_t> surviving_col_leaves;
  bool any_applied = false;
};

// The running-total target rides in the same `value` slot the item count
// uses, since the three flavours share one dialog field.
double FilterTarget(const PivotFilter& f) {
  if (const auto* as_double = std::get_if<double>(&f.value)) {
    return *as_double;
  }
  if (const auto* as_int = std::get_if<int>(&f.value)) {
    return static_cast<double>(*as_int);
  }
  return 0.0;
}

// `basis` selects how a top-N entry counts: `Items` uses the
// item-count rule, the other two accumulate a running total. `top`
// selects the ranking direction: `true` for Top N, `false` for Bottom
// N. Both are only ever non-default for an authored entry, since the
// dialog choice has no embedder-facing counterpart on `PivotFilter`.
// `field_index` names the filtered field; one that is absent from the
// filter's axis falls back to that axis's innermost field.
void ApplyValueFilter(ValueFilterState& state, const PivotFilter& f, std::optional<std::size_t> field_index,
                      TopNBasis basis = TopNBasis::Items, bool top = true) {
  const PivotTable& table = state.table;
  const PivotCache& cache = state.cache;
  const RecordBuckets& buckets = state.buckets;
  PivotResult& result = state.result;
  const std::size_t data_field_count = table.data_fields().size();
  if (f.type != FilterType::ValueTop10 && f.type != FilterType::ValueGreaterThan &&
      f.type != FilterType::ValueBetween) {
    return;  // Label/Date filters handled pre-aggregation.
  }
  if (data_field_count == 0) {
    return;
  }
  // A value filter names the measure whose aggregate determines ranking.
  // Invalid selectors are deliberately a no-op for filters constructed
  // directly against the C++ model; public binding APIs reject them before
  // mutation.
  if (f.data_field_index >= data_field_count) {
    return;
  }
  const bool row_axis = f.axis == PivotAxis::Row;
  const std::vector<std::uint32_t>& order = row_axis ? table.row_field_order() : table.col_field_order();
  if (order.empty()) {
    return;  // An axis with no fields has nothing to prune.
  }
  std::size_t depth = order.size() - 1U;
  if (field_index) {
    const auto it = std::find(order.begin(), order.end(), *field_index);
    if (it != order.end()) {
      depth = static_cast<std::size_t>(it - order.begin());
    }
  }
  // `leaf_count` is the number of leaves currently standing on the
  // filtered axis, which is what `result.values` is indexed by along it.
  const std::vector<std::size_t>& axis_leaves = row_axis ? state.surviving_row_leaves : state.surviving_col_leaves;
  const std::vector<std::size_t>& cross_leaves = row_axis ? state.surviving_col_leaves : state.surviving_row_leaves;
  const std::size_t leaf_count = axis_leaves.size();
  if (leaf_count == 0 || cross_leaves.empty()) {
    return;
  }
  const std::vector<AxisLeafGroup> groups =
      row_axis ? axis_leaf_groups_at_depth(result.rows, depth) : axis_leaf_groups_at_depth(result.cols, depth);
  // Each item's score is its own aggregate re-derived from the records,
  // so a non-additive aggregation ranks by the value the item displays.
  const PivotDataField& df = table.data_fields()[f.data_field_index];
  AxisScores scores;
  scores.scores.assign(groups.size(), 0.0);
  scores.all_blank.assign(groups.size(), true);
  for (std::size_t g = 0; g < groups.size(); ++g) {
    std::vector<std::size_t> group_leaves;
    for (std::size_t i = groups[g].first_leaf; i < groups[g].first_leaf + groups[g].leaf_count && i < leaf_count; ++i) {
      group_leaves.push_back(axis_leaves[i]);
    }
    std::vector<Value> column;
    append_leaf_set_field_values(cache, buckets, row_axis ? group_leaves : cross_leaves,
                                 row_axis ? cross_leaves : group_leaves, df.field_index, column);
    if (column.empty()) {
      continue;
    }
    if (const auto n = numeric_aggregate_value(apply_aggregation(df.aggregation, column))) {
      scores.scores[g] = *n;
      scores.all_blank[g] = false;
    }
  }
  const auto keep_or = build_grouped_value_filter_keep(f, basis, FilterTarget(f), top, groups, scores, leaf_count);
  if (!keep_or) {
    return;
  }
  const std::vector<bool>& keep = *keep_or;
  // Prune the hierarchy: leaves survive when `keep[leaf] == true` and
  // interior nodes survive when at least one descendant leaf does. Then
  // re-express every leaf-indexed structure in the surviving-leaf index
  // space, preserving DFS order.
  if (row_axis) {
    prune_top_level(result.rows, keep);
    compact_leaf_axis(result, keep, LeafAxis::Row, state.row_subtotal_leaf_sets, state.col_subtotal_leaf_sets);
    CompactSurvivingLeaves(state.surviving_row_leaves, keep);
  } else {
    prune_top_level(result.cols, keep);
    compact_leaf_axis(result, keep, LeafAxis::Col, state.row_subtotal_leaf_sets, state.col_subtotal_leaf_sets);
    CompactSurvivingLeaves(state.surviving_col_leaves, keep);
  }
  state.any_applied = true;
}

}  // namespace

void apply_value_filters(const PivotTable& table, const PivotCache& cache, const RecordBuckets& buckets,
                         PivotResult& result, std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                         std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets) {
  ValueFilterState state{table, cache, buckets, result, row_subtotal_leaf_sets, col_subtotal_leaf_sets, {}, {}};
  state.surviving_row_leaves.resize(buckets.size());
  std::iota(state.surviving_row_leaves.begin(), state.surviving_row_leaves.end(), std::size_t{0});
  state.surviving_col_leaves.resize(buckets.front().size());
  std::iota(state.surviving_col_leaves.begin(), state.surviving_col_leaves.end(), std::size_t{0});

  for (const PivotFilter& f : table.active_filters()) {
    const std::optional<std::size_t> field_index = resolve_field_by_any_name(table, f.field_name);
    // An unresolved active filter has no field to act on. Treat it as the
    // no-op promised by PreparedRecordFilter instead of allowing a value
    // filter to fall back to the innermost axis field.
    if (!field_index.has_value()) {
      continue;
    }
    ApplyValueFilter(state, f, field_index);
  }
  for (const AuthoredValueFilter& authored : table.authored_value_filters()) {
    if (const auto projected = authored_value_filter_as_pivot_filter(table, authored)) {
      ApplyValueFilter(state, *projected, static_cast<std::size_t>(authored.field_index), authored.top_n_basis,
                       authored.top);
    }
  }
  if (state.any_applied) {
    ReaggregateTotalsAfterValueFilter(table, cache, buckets, state.surviving_row_leaves, state.surviving_col_leaves,
                                      row_subtotal_leaf_sets, col_subtotal_leaf_sets, result);
  }
}

}  // namespace formulon::pivot
