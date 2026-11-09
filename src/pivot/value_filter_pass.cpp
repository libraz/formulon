
#include "pivot/value_filter_pass.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <utility>
#include <variant>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/calculated.h"
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

// Maps a subtotal's leaf set (in current, post-filter leaf-position
// space) to `buckets` coordinates via `original_of_position`.
std::vector<std::size_t> to_original_leaves(const std::vector<std::size_t>& leaf_set,
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
// lookup (`to_original_leaves`).
void reaggregate_totals_after_value_filter(const PivotTable& table, const PivotCache& cache,
                                           const RecordBuckets& buckets,
                                           const std::vector<std::size_t>& surviving_row_leaves,
                                           const std::vector<std::size_t>& surviving_col_leaves,
                                           const std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                                           const std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets,
                                           PivotResult& result) {
  const std::size_t data_field_count = table.data_fields().size();
  if (data_field_count == 0) {
    return;
  }

  // A total across both axes combines row partials; a one-axis layout chains its records.
  const bool both_axes = !table.row_field_order().empty() && !table.col_field_order().empty();

  std::vector<RecordIndexSpan> grand_total_spans;
  const RecordIndexView grand_total_records =
      make_leaf_set_record_view(buckets, surviving_row_leaves, surviving_col_leaves, grand_total_spans);
  for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < result.grand_totals.size(); ++df_idx) {
    const PivotDataField& df = table.data_fields()[df_idx];
    result.grand_totals[df_idx] =
        both_axes ? aggregate_row_partials_or_blank(df.aggregation, cache, buckets, surviving_row_leaves,
                                                    surviving_col_leaves, df.field_index, result)
                  : aggregate_records_or_blank(df.aggregation, cache, grand_total_records, df.field_index, result);
  }
  if (!result.grand_totals.empty()) {
    result.grand_total = result.grand_totals[0];
  }

  std::vector<RecordIndexSpan> spans;
  for (std::size_t r = 0; r < result.row_leaf_totals.size() && r < surviving_row_leaves.size(); ++r) {
    const RecordIndexView records =
        make_leaf_set_record_view(buckets, surviving_row_leaves[r], surviving_col_leaves, spans);
    for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < result.row_leaf_totals[r].size(); ++df_idx) {
      const PivotDataField& df = table.data_fields()[df_idx];
      result.row_leaf_totals[r][df_idx] =
          aggregate_records_or_blank(df.aggregation, cache, records, df.field_index, result);
    }
  }

  for (std::size_t c = 0; c < result.col_leaf_totals.size() && c < surviving_col_leaves.size(); ++c) {
    const RecordIndexView records =
        make_leaf_set_record_view(buckets, surviving_row_leaves, surviving_col_leaves[c], spans);
    for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < result.col_leaf_totals[c].size(); ++df_idx) {
      const PivotDataField& df = table.data_fields()[df_idx];
      result.col_leaf_totals[c][df_idx] =
          aggregate_records_or_blank(df.aggregation, cache, records, df.field_index, result);
    }
  }

  for (std::size_t rs = 0; rs < result.row_subtotals.size() && rs < row_subtotal_leaf_sets.size(); ++rs) {
    RowSubtotal& subtotal = result.row_subtotals[rs];
    const std::vector<std::size_t> rows = to_original_leaves(row_subtotal_leaf_sets[rs], surviving_row_leaves);
    std::vector<RecordIndexSpan> subtotal_spans;
    const RecordIndexView subtotal_records =
        make_leaf_set_record_view(buckets, rows, surviving_col_leaves, subtotal_spans);
    std::vector<std::vector<RecordIndexSpan>> col_spans(surviving_col_leaves.size());
    std::vector<RecordIndexView> col_records;
    col_records.reserve(surviving_col_leaves.size());
    for (std::size_t c = 0; c < surviving_col_leaves.size(); ++c) {
      col_records.push_back(make_leaf_set_record_view(buckets, rows, surviving_col_leaves[c], col_spans[c]));
    }
    for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.values.size(); ++df_idx) {
      const PivotDataField& df = table.data_fields()[df_idx];
      const Aggregation aggregation = aggregation_for(subtotal.aggregation, df);
      const MissingStats missing =
          subtotal.aggregation.has_value()
              ? missing_custom_subtotal_stats(table, subtotal.depth, aggregation, df.aggregation)
              : MissingStats{};
      subtotal.values[df_idx] =
          both_axes ? aggregate_row_partials_or_blank(aggregation, cache, buckets, rows, surviving_col_leaves,
                                                      df.field_index, result, missing)
                    : aggregate_records_or_blank(aggregation, cache, subtotal_records, df.field_index, result);
    }
    for (std::size_t c = 0; c < subtotal.col_values.size() && c < surviving_col_leaves.size(); ++c) {
      for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.col_values[c].size(); ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        subtotal.col_values[c][df_idx] = aggregate_records_or_blank(aggregation_for(subtotal.aggregation, df), cache,
                                                                    col_records[c], df.field_index, result);
      }
    }
    for (std::size_t cs = 0; cs < subtotal.col_subtotal_values.size() && cs < col_subtotal_leaf_sets.size(); ++cs) {
      const std::vector<std::size_t> cols = to_original_leaves(col_subtotal_leaf_sets[cs], surviving_col_leaves);
      const ColSubtotal& col_subtotal = result.col_subtotals[cs];
      for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.col_subtotal_values[cs].size();
           ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        const Aggregation aggregation = aggregation_for(col_subtotal.aggregation, df);
        subtotal.col_subtotal_values[cs][df_idx] =
            aggregate_row_partials_or_blank(aggregation, cache, buckets, rows, cols, df.field_index, result);
      }
    }
  }

  for (std::size_t cs = 0; cs < result.col_subtotals.size() && cs < col_subtotal_leaf_sets.size(); ++cs) {
    ColSubtotal& subtotal = result.col_subtotals[cs];
    const std::vector<std::size_t> cols = to_original_leaves(col_subtotal_leaf_sets[cs], surviving_col_leaves);
    std::vector<std::vector<RecordIndexSpan>> row_spans(surviving_row_leaves.size());
    std::vector<RecordIndexView> row_records;
    row_records.reserve(surviving_row_leaves.size());
    for (std::size_t r = 0; r < surviving_row_leaves.size(); ++r) {
      row_records.push_back(make_leaf_set_record_view(buckets, surviving_row_leaves[r], cols, row_spans[r]));
    }
    std::vector<RecordIndexSpan> total_spans;
    const RecordIndexView total_records = make_leaf_set_record_view(buckets, surviving_row_leaves, cols, total_spans);
    for (std::size_t r = 0; r < subtotal.values.size() && r < surviving_row_leaves.size(); ++r) {
      for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.values[r].size(); ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        const Aggregation aggregation = aggregation_for(subtotal.aggregation, df);
        subtotal.values[r][df_idx] =
            aggregate_records_or_blank(aggregation, cache, row_records[r], df.field_index, result);
      }
    }
    for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.total.size(); ++df_idx) {
      const PivotDataField& df = table.data_fields()[df_idx];
      const Aggregation aggregation = aggregation_for(subtotal.aggregation, df);
      subtotal.total[df_idx] =
          both_axes ? aggregate_row_partials_or_blank(aggregation, cache, buckets, surviving_row_leaves, cols,
                                                      df.field_index, result)
                    : aggregate_records_or_blank(aggregation, cache, total_records, df.field_index, result);
    }
  }

  // Sum totals are cell-based; the surviving cells are `result.values`.
  recompute_sum_totals_from_cells(table, result, row_subtotal_leaf_sets, col_subtotal_leaf_sets);
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
  const CalculatedPlan& plan;
  bool has_calculated_cells;
  std::vector<std::size_t> surviving_row_leaves;
  std::vector<std::size_t> surviving_col_leaves;
  bool any_applied = false;
  /// Set once a filter kept only items without records.
  bool keep_totals = false;
};

// The value a group of leaves shows for data field `slot`: the fold of its
// cells across the standing cross leaves, a calculated field evaluated over
// the folded Sums. `records` are the group's records, for the functions a
// fold cannot rebuild.
Value displayed_group_value(const ValueFilterState& state, bool row_axis, const std::vector<std::size_t>& positions,
                            const RecordIndexView& records, std::size_t slot) {
  const PivotTable& table = state.table;
  const std::size_t cross_count = row_axis ? state.surviving_col_leaves.size() : state.surviving_row_leaves.size();
  const std::vector<std::size_t> cross = identity_indices(cross_count);
  const std::vector<bool>& calculated = state.plan.calculated_slots();
  std::vector<Value> slots(table.data_fields().size(), Value::blank());
  for (std::size_t j = 0; j < slots.size(); ++j) {
    if (j < calculated.size() && calculated[j]) {
      continue;
    }
    const PivotDataField& df = table.data_fields()[j];
    if (state.has_calculated_cells && aggregation_combines_cells(df.aggregation)) {
      slots[j] =
          fold_leaf_cells(df.aggregation, state.result, row_axis ? positions : cross, row_axis ? cross : positions, j);
    } else if (!records.empty()) {
      slots[j] = apply_record_aggregation(df.aggregation, state.cache, records, df.field_index);
    }
  }
  state.plan.fill_field_slots(slots, state.result);
  return slots[slot];
}

// Applies the value filter `f` to the leaves of `axis`, pruning the items of
// the field at `depth` in that axis's order.
void apply_value_filter(ValueFilterState& state, const PivotFilter& f, PivotAxis axis, std::size_t depth) {
  const PivotTable& table = state.table;
  const PivotCache& cache = state.cache;
  const RecordBuckets& buckets = state.buckets;
  PivotResult& result = state.result;
  const std::size_t data_field_count = table.data_fields().size();
  // A value filter names the measure whose aggregate determines ranking.
  // Invalid selectors are deliberately a no-op for filters constructed
  // directly against the C++ model; public binding APIs reject them before
  // mutation.
  if (f.data_field_index >= data_field_count) {
    return;
  }
  const bool row_axis = axis == PivotAxis::Row;
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
  const std::vector<bool>& calculated = state.plan.calculated_slots();
  const bool displayed =
      state.has_calculated_cells || (f.data_field_index < calculated.size() && calculated[f.data_field_index]);
  AxisScores scores;
  scores.scores.assign(groups.size(), 0.0);
  scores.all_blank.assign(groups.size(), true);
  std::vector<bool> has_records(groups.size(), false);
  std::vector<RecordIndexSpan> spans;
  for (std::size_t g = 0; g < groups.size(); ++g) {
    std::vector<std::size_t> positions;
    std::vector<std::size_t> group_leaves;
    for (std::size_t i = groups[g].first_leaf; i < groups[g].first_leaf + groups[g].leaf_count && i < leaf_count; ++i) {
      positions.push_back(i);
      group_leaves.push_back(axis_leaves[i]);
    }
    const RecordIndexView records = make_leaf_set_record_view(buckets, row_axis ? group_leaves : cross_leaves,
                                                              row_axis ? cross_leaves : group_leaves, spans);
    has_records[g] = !records.empty();
    Value value = Value::blank();
    if (displayed) {
      value = displayed_group_value(state, row_axis, positions, records, f.data_field_index);
    } else if (!records.empty()) {
      value = apply_record_aggregation(df.aggregation, cache, records, df.field_index);
    }
    if (const auto n = numeric_aggregate_value(value)) {
      scores.scores[g] = *n;
      scores.all_blank[g] = false;
    } else if (records.empty()) {
      // An item listed without data compares as 0.
      scores.all_blank[g] = false;
    }
  }
  const auto keep_or = build_grouped_value_filter_keep(f, groups, scores, leaf_count);
  if (!keep_or) {
    return;
  }
  const std::vector<bool>& keep = *keep_or;
  bool kept_any = false;
  bool kept_data = false;
  for (std::size_t g = 0; g < groups.size(); ++g) {
    const std::size_t first = groups[g].first_leaf;
    if (first < keep.size() && keep[first]) {
      kept_any = true;
      kept_data = kept_data || has_records[g];
    }
  }
  state.keep_totals = state.keep_totals || (kept_any && !kept_data);
  // Prune the hierarchy: leaves survive when `keep[leaf] == true` and
  // interior nodes survive when at least one descendant leaf does. Then
  // re-express every leaf-indexed structure in the surviving-leaf index
  // space, preserving DFS order.
  if (row_axis) {
    prune_top_level(result.rows, keep);
    compact_leaf_axis(result, keep, LeafAxis::Row, state.row_subtotal_leaf_sets, state.col_subtotal_leaf_sets);
    keep_marked(state.surviving_row_leaves, keep);
  } else {
    prune_top_level(result.cols, keep);
    compact_leaf_axis(result, keep, LeafAxis::Col, state.row_subtotal_leaf_sets, state.col_subtotal_leaf_sets);
    keep_marked(state.surviving_col_leaves, keep);
  }
  state.any_applied = true;
}

// The axis and depth `f` prunes. The field's own position when it sits on a
// report axis; otherwise an entry added through the API falls back to the
// innermost field of its stated axis, while a decoded one filters nothing.
std::optional<std::pair<PivotAxis, std::size_t>> value_filter_target(const PivotTable& table, const PivotFilter& f,
                                                                     std::size_t field_index) {
  const std::vector<std::uint32_t>& stated =
      f.axis == PivotAxis::Col ? table.col_field_order() : table.row_field_order();
  const PivotAxis stated_axis = f.axis == PivotAxis::Col ? PivotAxis::Col : PivotAxis::Row;
  if (const auto depth = field_position_in_order(stated, field_index)) {
    return std::pair<PivotAxis, std::size_t>{stated_axis, *depth};
  }
  if (!f.source.has_value()) {
    if (stated.empty()) {
      return std::nullopt;  // An axis with no fields has nothing to prune.
    }
    return std::pair<PivotAxis, std::size_t>{stated_axis, stated.size() - 1U};
  }
  const PivotAxis other_axis = stated_axis == PivotAxis::Row ? PivotAxis::Col : PivotAxis::Row;
  if (const auto depth = field_position_in_order(
          other_axis == PivotAxis::Col ? table.col_field_order() : table.row_field_order(), field_index)) {
    return std::pair<PivotAxis, std::size_t>{other_axis, *depth};
  }
  return std::nullopt;
}

}  // namespace

bool apply_value_filters(const PivotTable& table, const PivotCache& cache, const RecordBuckets& buckets,
                         std::size_t row_leaf_count, std::size_t col_leaf_count, PivotResult& result,
                         std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                         std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets, const CalculatedPlan& plan,
                         bool has_calculated_cells) {
  ValueFilterState state{
      table, cache, buckets, result, row_subtotal_leaf_sets, col_subtotal_leaf_sets, plan, has_calculated_cells,
      {},    {}};
  state.surviving_row_leaves = identity_indices(row_leaf_count);
  state.surviving_col_leaves = identity_indices(col_leaf_count);

  // Each filter prunes what the previous ones left, in document order.
  if (!table.data_fields().empty()) {
    for (const PivotFilter& f : table.active_filters()) {
      if (!is_value_filter_type(f.type)) {
        continue;  // Caption and date filters are applied per record.
      }
      // An unresolved filter has no field to act on: the no-op promised by
      // `PreparedRecordFilter`, never a fallback to the innermost axis field.
      const std::optional<std::size_t> field_index = filter_field_index(table, f);
      if (!field_index.has_value()) {
        continue;
      }
      if (const auto target = value_filter_target(table, f, *field_index)) {
        apply_value_filter(state, f, target->first, target->second);
      }
    }
  }
  // A filter that keeps only empty items leaves the report's totals standing.
  if (!state.any_applied || state.keep_totals) {
    return false;
  }
  reaggregate_totals_after_value_filter(table, cache, buckets, state.surviving_row_leaves, state.surviving_col_leaves,
                                        row_subtotal_leaf_sets, col_subtotal_leaf_sets, result);
  return true;
}

}  // namespace formulon::pivot
