// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Pivot evaluator orchestration. See header / §15.1.3 of the design
// corpus for the algorithm overview. The MVP path implemented here
// produces enough of a `PivotResult` that GETPIVOTDATA can resolve
// label/data tuples against the freshest evaluation snapshot.
//
// This translation unit owns the seven-step pipeline:
//
//   1. validate cache_id + data field bounds;
//   2. filter records (manual + axis label/date filters);
//   3. build the row / col hierarchies and bucket surviving records by
//      (row_leaf, col_leaf);
//   4. aggregate per (row_leaf, col_leaf, data_field);
//   5. emit row / col / row x col subtotals;
//   6. emit grand totals;
//   7. apply value-axis filters and show-values-as transforms.
//
// The heavy lifting lives in sibling TUs (`aggregator`, `filter_engine`,
// `hierarchy_builder`, `layout_generator`); the routines here are just
// glue.

#include "pivot/pivot_evaluator.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/filter_engine.h"
#include "pivot/hierarchy_builder.h"
#include "pivot/layout_generator.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// ---------------------------------------------------------------------------
// Result-side text reification.
// ---------------------------------------------------------------------------
//
// `PivotResult::values` / `subtotals` / grand totals must outlive the
// cache they were computed against (GETPIVOTDATA reads them outside of
// any specific evaluation arena). Numbers, bools, errors, and blanks
// are trivially copyable. Text is the only kind that needs storage —
// we copy the bytes into `result.text_storage` and rebuild a `Value`
// pointing into the deque entry. Pointer/iterator stability of
// `std::deque` keeps the views valid across subsequent appends.
Value reify(const Value& v, PivotResult& result) {
  if (!v.is_text()) {
    return v;
  }
  result.text_storage.emplace_back(v.as_text());
  return Value::text(result.text_storage.back());
}

// `SubtotalFn` and `Aggregation` share the same ordinal layout (Sum=0 ..
// VarP=10), so a custom subtotal function maps to the matching aggregation
// by ordinal.
Aggregation aggregation_from_subtotal_fn(SubtotalFn fn) {
  return static_cast<Aggregation>(static_cast<std::uint8_t>(fn));
}

}  // namespace

Expected<PivotResult, Error> evaluate(const PivotTable& table, const PivotCache& cache) {
  // 1. Validate.
  if (table.pivot_cache_id() != cache.cache_id()) {
    return make_error(FormulonErrorCode::kEvalPivotMissing, "pivot table cache_id does not match supplied PivotCache",
                      "table=" + table.name() + " table.cache_id=" + std::to_string(table.pivot_cache_id()) +
                          " cache.cache_id=" + std::to_string(cache.cache_id()));
  }
  for (std::size_t i = 0; i < table.data_fields().size(); ++i) {
    const PivotDataField& df = table.data_fields()[i];
    if (df.field_index >= cache.fields().size()) {
      return make_error(FormulonErrorCode::kEvalPivotInvalid, "data field references out-of-range cache field",
                        "data_field=" + df.name + " field_index=" + std::to_string(df.field_index) +
                            " cache_fields=" + std::to_string(cache.fields().size()));
    }
  }

  // 2. Filter records.
  std::vector<std::size_t> surviving;
  surviving.reserve(cache.records().size());
  for (std::size_t i = 0; i < cache.records().size(); ++i) {
    if (record_passes_manual_filter(table, cache, cache.records()[i])) {
      surviving.push_back(i);
    }
  }

  // 3. Build hierarchies.
  //
  // A field whose `date_group` is set bucketises the cache value at
  // hierarchy-insertion time; we plumb the optional through `HierLevel`
  // so `insert_path` can call the bucketer without re-walking the table
  // metadata.
  auto level_for = [&](std::uint32_t fi) -> HierLevel {
    const PivotDateGroup* dg = nullptr;
    if (fi < table.fields().size() && table.fields()[fi].date_group.has_value()) {
      dg = &*table.fields()[fi].date_group;
    }
    return HierLevel{fi, dg};
  };
  std::vector<HierLevel> row_levels;
  row_levels.reserve(table.row_field_order().size());
  for (std::uint32_t fi : table.row_field_order()) {
    row_levels.push_back(level_for(fi));
  }
  std::vector<HierLevel> col_levels;
  col_levels.reserve(table.col_field_order().size());
  for (std::uint32_t fi : table.col_field_order()) {
    col_levels.push_back(level_for(fi));
  }

  HierNode row_tree;
  HierNode col_tree;

  // For each surviving record, remember which leaf it lands on (row +
  // col). Indices are looked up after finalisation so we don't need to
  // walk the tree a second time during aggregation.
  std::vector<HierNode*> row_leaves_for_record(surviving.size(), nullptr);
  std::vector<HierNode*> col_leaves_for_record(surviving.size(), nullptr);

  for (std::size_t i = 0; i < surviving.size(); ++i) {
    const PivotCacheRecord& rec = cache.records()[surviving[i]];
    if (!row_levels.empty()) {
      row_leaves_for_record[i] = insert_path(cache, row_levels, rec, row_tree);
    }
    if (!col_levels.empty()) {
      col_leaves_for_record[i] = insert_path(cache, col_levels, rec, col_tree);
    }
  }

  PivotResult result;
  std::vector<HierNode*> row_leaves;
  std::vector<HierNode*> col_leaves;
  finalize_hierarchy<RowHierarchyNode>(row_tree, result.rows, row_leaves);
  finalize_hierarchy<ColHierarchyNode>(col_tree, result.cols, col_leaves);

  // Degenerate axis: if a side has no field configured, treat it as a
  // single implicit leaf so the values matrix still has a slot per
  // surviving record group on the populated axis.
  const std::size_t row_leaf_count = row_levels.empty() ? 1 : row_leaves.size();
  const std::size_t col_leaf_count = col_levels.empty() ? 1 : col_leaves.size();
  const std::size_t data_field_count = table.data_fields().size();

  // Guards on the dense (row_leaf x col_leaf) matrices, evaluated BEFORE
  // the first dense allocation so a pathological high-cardinality cache
  // cannot commit a huge allocation first:
  //   * checked multiplication — on 32-bit `size_t` (WASM) the product can
  //     wrap, leaving the nested vectors inconsistent and downstream code
  //     indexing past their end;
  //   * result-cell budget — even a non-wrapping product can describe a
  //     matrix far past anything a real pivot produces.
  // Both surface `kFnOverflow` so the caller keeps one recoverable path.
  auto value_count_or = checked_mul_size_t(row_leaf_count, col_leaf_count);
  if (!value_count_or) {
    return value_count_or.error();
  }
  ResourceBudget result_cell_budget(kMaxPivotResultCells, FormulonErrorCode::kFnOverflow);
  auto budget_ok = result_cell_budget.consume(static_cast<std::uint64_t>(value_count_or.value()));
  if (!budget_ok) {
    return budget_ok.error();
  }

  // Bucket surviving record indices by (row_leaf, col_leaf).
  // `[row_leaf][col_leaf]` -> indices into `cache.records()`.
  RecordBuckets buckets(row_leaf_count, std::vector<std::vector<std::size_t>>(col_leaf_count));

  for (std::size_t i = 0; i < surviving.size(); ++i) {
    const std::size_t r = row_levels.empty() ? 0 : row_leaves_for_record[i]->leaf_index;
    const std::size_t c = col_levels.empty() ? 0 : col_leaves_for_record[i]->leaf_index;
    buckets[r][c].push_back(surviving[i]);
  }

  // 4. Aggregate per (row_leaf, col_leaf, data_field). The dense-matrix
  // guards above already validated `row_leaf_count * col_leaf_count`.
  result.values.assign(row_leaf_count, std::vector<std::vector<Value>>(col_leaf_count));
  for (std::size_t r = 0; r < row_leaf_count; ++r) {
    for (std::size_t c = 0; c < col_leaf_count; ++c) {
      result.values[r][c].reserve(data_field_count);
      const std::vector<std::size_t>& records = buckets[r][c];
      for (const PivotDataField& df : table.data_fields()) {
        std::vector<Value> column;
        column.reserve(records.size());
        append_record_field_values(cache, records, df.field_index, column);
        result.values[r][c].push_back(reify(apply_aggregation(df.aggregation, column), result));
      }
    }
  }

  // 4b. Per-leaf totals across the opposite axis, re-aggregated from the
  // records themselves. Non-additive aggregations (Average/Max/Min/StdDev/
  // Var) cannot be recovered by summing the per-cell aggregates, so the
  // rendered "Grand Total" row/column reads these instead. Computed from
  // the pre-value-filter buckets; the value-filter step below compacts
  // them in lock-step with `result.values`.
  if (data_field_count > 0) {
    result.row_leaf_totals.assign(row_leaf_count, std::vector<Value>(data_field_count, Value::blank()));
    for (std::size_t r = 0; r < row_leaf_count; ++r) {
      for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        std::vector<Value> column;
        for (std::size_t c = 0; c < col_leaf_count; ++c) {
          append_record_field_values(cache, buckets[r][c], df.field_index, column);
        }
        result.row_leaf_totals[r][df_idx] = reify(apply_aggregation(df.aggregation, column), result);
      }
    }
    result.col_leaf_totals.assign(col_leaf_count, std::vector<Value>(data_field_count, Value::blank()));
    for (std::size_t c = 0; c < col_leaf_count; ++c) {
      for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        std::vector<Value> column;
        for (std::size_t r = 0; r < row_leaf_count; ++r) {
          append_record_field_values(cache, buckets[r][c], df.field_index, column);
        }
        result.col_leaf_totals[c][df_idx] = reify(apply_aggregation(df.aggregation, column), result);
      }
    }
  }

  // 5. Row-direction subtotals.
  //
  // Walk the row hierarchy; at each non-leaf level whose field declares
  // `subtotal_top` or any `subtotal_fns`, aggregate the union of all
  // descendant leaves' records using the data field's own aggregation.
  // For MVP we surface one subtotal slot per data field. Column-axis
  // subtotals are deferred.
  //
  // The flat-list shape (`subtotals[i]` is one row of the result, no
  // tree mirror) is convenient for GETPIVOTDATA, which addresses
  // subtotals by the sequence in which they appear when walking the row
  // hierarchy in document order.
  std::vector<std::vector<std::size_t>> row_subtotal_leaf_sets;
  std::vector<std::vector<std::size_t>> col_subtotal_leaf_sets;

  if (!row_levels.empty() && data_field_count > 0) {
    std::vector<std::size_t> stack_row_leaves;  // current path's leaf indices
    std::vector<std::vector<Value>>& subtotals = result.subtotals;

    walk_subtotal_tree(
        row_tree, row_levels, table, stack_row_leaves,
        [&](const std::vector<std::string>& labels, std::size_t depth, std::size_t collected_start,
            const std::vector<std::size_t>& leaves) {
          // Gather each data field's underlying record values over the
          // group once (both the flat column and the per-column-leaf
          // split), then aggregate them once per subtotal function.
          std::vector<std::vector<Value>> df_columns(data_field_count);
          std::vector<std::vector<std::vector<Value>>> df_columns_by_col(
              data_field_count, std::vector<std::vector<Value>>(col_leaf_count));
          for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
            const std::uint32_t field_index = table.data_fields()[df_idx].field_index;
            for (std::size_t leaf_idx_iter = collected_start; leaf_idx_iter < leaves.size(); ++leaf_idx_iter) {
              const std::size_t leaf_idx = leaves[leaf_idx_iter];
              for (std::size_t c = 0; c < col_leaf_count; ++c) {
                for (std::size_t rec_idx : buckets[leaf_idx][c]) {
                  Value v = cell_value(cache, cache.records()[rec_idx], field_index);
                  df_columns[df_idx].push_back(v);
                  df_columns_by_col[df_idx][c].push_back(v);
                }
              }
            }
          }
          // A row field with explicit custom subtotal functions emits one
          // subtotal row per selected function; otherwise a single default
          // subtotal uses each data field's own summary function. An empty
          // optional in `specs` marks the default (per-data-field) case.
          std::vector<std::optional<Aggregation>> specs;
          if (depth < table.row_field_order().size()) {
            const std::uint32_t group_fi = table.row_field_order()[depth];
            if (group_fi < table.fields().size() && !table.fields()[group_fi].subtotal_fns.empty()) {
              for (const SubtotalFn fn : table.fields()[group_fi].subtotal_fns) {
                specs.push_back(aggregation_from_subtotal_fn(fn));
              }
            }
          }
          if (specs.empty()) {
            specs.push_back(std::nullopt);
          }
          for (const std::optional<Aggregation>& spec : specs) {
            std::vector<Value> row_values(data_field_count, Value::blank());
            std::vector<std::vector<Value>> col_values(col_leaf_count,
                                                       std::vector<Value>(data_field_count, Value::blank()));
            for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
              const Aggregation agg = spec.has_value() ? *spec : table.data_fields()[df_idx].aggregation;
              row_values[df_idx] = reify(apply_aggregation(agg, df_columns[df_idx]), result);
              for (std::size_t c = 0; c < col_leaf_count; ++c) {
                col_values[c][df_idx] = reify(apply_aggregation(agg, df_columns_by_col[df_idx][c]), result);
              }
            }
            RowSubtotal subtotal;
            subtotal.labels = labels;
            subtotal.depth = static_cast<std::uint32_t>(depth);
            subtotal.values = row_values;
            subtotal.col_values = std::move(col_values);
            row_subtotal_leaf_sets.emplace_back(leaves.begin() + static_cast<std::ptrdiff_t>(collected_start),
                                                leaves.end());
            result.row_subtotals.push_back(std::move(subtotal));
            subtotals.push_back(std::move(row_values));
          }
        });
  }

  // 5b. Column-direction subtotals. The shape mirrors row_subtotals but each
  // subtotal stores one row-leaf x data-field matrix because a rendered
  // subtotal column has one value per row leaf.
  if (!col_levels.empty() && data_field_count > 0) {
    std::vector<std::size_t> stack_col_leaves;

    walk_subtotal_tree(col_tree, col_levels, table, stack_col_leaves,
                       [&](const std::vector<std::string>& labels, std::size_t depth, std::size_t collected_start,
                           const std::vector<std::size_t>& leaves) {
                         ColSubtotal subtotal;
                         subtotal.labels = labels;
                         subtotal.depth = static_cast<std::uint32_t>(depth);
                         subtotal.values.assign(row_leaf_count, std::vector<Value>(data_field_count, Value::blank()));

                         for (std::size_t r = 0; r < row_leaf_count; ++r) {
                           for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
                             const PivotDataField& df = table.data_fields()[df_idx];
                             std::vector<Value> column;
                             for (std::size_t leaf_idx_iter = collected_start; leaf_idx_iter < leaves.size();
                                  ++leaf_idx_iter) {
                               const std::size_t leaf_idx = leaves[leaf_idx_iter];
                               append_bucket_field_values(cache, buckets, r, leaf_idx, df.field_index, column);
                             }
                             subtotal.values[r][df_idx] = reify(apply_aggregation(df.aggregation, column), result);
                           }
                         }
                         col_subtotal_leaf_sets.emplace_back(
                             leaves.begin() + static_cast<std::ptrdiff_t>(collected_start), leaves.end());
                         result.col_subtotals.push_back(std::move(subtotal));
                       });
  }

  if (!result.row_subtotals.empty() && !result.col_subtotals.empty() && data_field_count > 0) {
    for (std::size_t rs = 0; rs < result.row_subtotals.size(); ++rs) {
      RowSubtotal& row_subtotal = result.row_subtotals[rs];
      row_subtotal.col_subtotal_values.assign(result.col_subtotals.size(),
                                              std::vector<Value>(data_field_count, Value::blank()));
      if (rs >= row_subtotal_leaf_sets.size()) {
        continue;
      }
      for (std::size_t cs = 0; cs < result.col_subtotals.size(); ++cs) {
        if (cs >= col_subtotal_leaf_sets.size()) {
          continue;
        }
        for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
          const PivotDataField& df = table.data_fields()[df_idx];
          std::vector<Value> column;
          append_leaf_set_field_values(cache, buckets, row_subtotal_leaf_sets[rs], col_subtotal_leaf_sets[cs],
                                       df.field_index, column);
          row_subtotal.col_subtotal_values[cs][df_idx] = reify(apply_aggregation(df.aggregation, column), result);
        }
      }
    }
  }

  // 6. Grand totals, one slot per data field. The legacy single-value
  // `grand_total` mirrors slot 0 for existing GETPIVOTDATA callers.
  if ((table.grand_totals_rows() || table.grand_totals_cols()) && data_field_count > 0) {
    result.grand_totals.reserve(data_field_count);
    for (const PivotDataField& df : table.data_fields()) {
      std::vector<Value> column;
      column.reserve(surviving.size());
      append_record_field_values(cache, surviving, df.field_index, column);
      result.grand_totals.push_back(reify(apply_aggregation(df.aggregation, column), result));
    }
    if (!result.grand_totals.empty()) {
      result.grand_total = result.grand_totals[0];
    }
  }

  // 7. Value-axis filters (Top-N, GreaterThan, Between).
  //
  // Applied last so the pre-aggregation filter set has already shaped
  // `result.values`; the pruning here only drops surviving leaves.
  // Multi-level hierarchies are honoured: we score each leaf in DFS
  // pre-order (the order `finalize_hierarchy` assigned), compute the
  // keep-mask for the whole leaf array, then collapse the row/col tree
  // by dropping leaves whose mask is false and any interior node whose
  // subtree becomes empty. Subtotals + grand totals retain their
  // pre-filter values so a Top-N report can still surface "X out of
  // total" framing.
  for (const PivotFilter& f : table.active_filters()) {
    if (f.type != FilterType::ValueTop10 && f.type != FilterType::ValueGreaterThan &&
        f.type != FilterType::ValueBetween) {
      continue;  // Label/Date filters handled pre-aggregation.
    }
    if (data_field_count == 0) {
      continue;
    }
    if (f.axis == PivotAxis::Row && !table.row_field_order().empty()) {
      // `n` is the number of row leaves (DFS pre-order), which is what
      // `result.values` is indexed by; `result.rows.size()` would be the
      // number of top-level row nodes and would understate `n` whenever
      // the row hierarchy is multi-level.
      const std::size_t n = result.values.size();
      if (n == 0) {
        continue;
      }
      const auto keep_or =
          build_value_filter_keep(f, score_row_axis(result, n, col_levels.empty() ? 1u : col_leaf_count));
      if (!keep_or) {
        continue;
      }
      const std::vector<bool>& keep = *keep_or;
      // Prune the row hierarchy: leaves survive when `keep[leaf] == true`
      // and interior nodes survive when at least one descendant leaf
      // does. Then compact `result.values` to the surviving leaves,
      // preserving DFS order.
      prune_top_level(result.rows, keep);
      compact_row_axis_values(result.values, keep);
      compact_leaf_totals(result.row_leaf_totals, keep);
    } else if (f.axis == PivotAxis::Col && !table.col_field_order().empty()) {
      // `n` is the number of column leaves (DFS pre-order). When the row
      // axis has at least one materialised slot we read the leaf count
      // from the first row's column slice; otherwise the matrix is empty
      // and the filter is a no-op below.
      const std::size_t n = result.values.empty() ? 0 : result.values[0].size();
      if (n == 0) {
        continue;
      }
      const std::size_t row_n = row_levels.empty() ? 1u : result.values.size();
      const auto keep_or = build_value_filter_keep(f, score_col_axis(result, n, row_n));
      if (!keep_or) {
        continue;
      }
      const std::vector<bool>& keep = *keep_or;
      // Prune the col hierarchy then compact every row's per-col slice.
      prune_top_level(result.cols, keep);
      compact_col_axis_values(result.values, keep);
      compact_leaf_totals(result.col_leaf_totals, keep);
    }
    // Mixed-direction (e.g. row-axis filter referencing a column field)
    // remains out of scope; such filters fall through here as a no-op.
  }

  // 8. Show-values-as transforms.
  apply_show_values_as_transforms(table, result, row_subtotal_leaf_sets, col_subtotal_leaf_sets);

  return result;
}

}  // namespace formulon::pivot
