//
// Pivot evaluator orchestration. See header / §15.1.3 of the design
// corpus for the algorithm overview. The MVP path implemented here
// produces enough of a `PivotResult` that GETPIVOTDATA can resolve
// label/data tuples against the freshest evaluation snapshot.
//
// This translation unit owns the eight-step pipeline:
//
//   1. validate cache_id + data field bounds;
//   2. filter records (manual + axis label/date filters);
//   3. build the row / col hierarchies and bucket surviving records by
//      (row_leaf, col_leaf);
//   4. aggregate per (row_leaf, col_leaf, data_field);
//   5. emit row / col / row x col subtotals;
//   6. emit grand totals;
//   7. apply value-axis filters;
//   8. apply show-values-as transforms.
//
// The heavy lifting lives in sibling TUs (`aggregator`, `filter_engine`,
// `hierarchy_builder`, `value_filter_pass`, `show_values_as`); the routines here are just
// glue.

#include "pivot/pivot_evaluator.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/date_serial.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/hierarchy_builder.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/show_values_as.h"
#include "pivot/value_filter_pass.h"
#include "pivot/value_order.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon::pivot {
namespace {

struct ValueSortSpec {
  std::uint32_t field_index;
  Aggregation aggregation;
};

std::optional<ValueSortSpec> resolve_value_sort(const PivotTable& table, std::uint32_t pivot_field_index) {
  if (pivot_field_index >= table.fields().size()) {
    return std::nullopt;
  }
  const std::string& by_field = table.fields()[pivot_field_index].sort.by_field;
  if (by_field.empty()) {
    return std::nullopt;
  }
  for (const PivotDataField& data_field : table.data_fields()) {
    if (data_field.field_index >= table.fields().size()) {
      continue;
    }
    const PivotField& source = table.fields()[data_field.field_index];
    if (by_field == data_field.name || pivot_field_has_name(source, by_field)) {
      return ValueSortSpec{data_field.field_index, data_field.aggregation};
    }
  }
  return std::nullopt;
}

/// Names what a page (report filter) field is currently showing.
///
/// Settled here rather than in the projection for the same reason a blank
/// axis item is: the answer comes from the bound cache, which `layout` is
/// never handed. An item built by cache index has no name until a load
/// resolves it, so labels are read through `pivot_item_label`.
///
/// Excel records the selection two different ways and both are honoured:
/// a single chosen item as `<pageField item="N">`, and any wider selection
/// as visibility flags on the field's own items. A field that hides
/// nothing is unfiltered, which is also what an empty item list means — a
/// table assembled in memory lists items only once something filters them.
std::string page_item_label(const PivotCache& cache, std::size_t field_index, const PivotField& field,
                            std::optional<std::uint32_t> selected, const PivotLayoutOptions& options) {
  const auto label_of = [&](const PivotItem& item) {
    std::string label = pivot_item_label(cache, field_index, item);
    return label.empty() ? options.blank_item_label : label;
  };
  if (selected.has_value() && *selected < field.items.size()) {
    return label_of(field.items[*selected]);
  }
  const PivotItem* only_visible = nullptr;
  std::size_t visible_count = 0;
  for (const PivotItem& item : field.items) {
    if (!item.visible) {
      continue;
    }
    ++visible_count;
    only_visible = &item;
  }
  if (visible_count == field.items.size()) {
    return options.all_pages_label;
  }
  if (visible_count == 1U && only_visible != nullptr) {
    return label_of(*only_visible);
  }
  return options.multiple_items_label;
}

void resolve_page_selections(const PivotTable& table, const PivotCache& cache, const PivotLayoutOptions& options,
                             PivotResult& result) {
  const std::vector<std::uint32_t> order = table.page_field_order();
  result.page_selections.reserve(order.size());
  for (const std::uint32_t field_index : order) {
    const PivotField& field = table.fields()[field_index];
    std::optional<std::uint32_t> selected;
    for (const PivotPageField& page : table.page_fields()) {
      if (page.field_index == field_index) {
        selected = page.item_index;
        break;
      }
    }
    result.page_selections.push_back(PivotPageSelection{field_index, pivot_field_display_name(field),
                                                        page_item_label(cache, field_index, field, selected, options)});
  }
}

// `SubtotalFn` and `Aggregation` share the same ordinal layout (Sum=0 ..
// VarP=10), so a custom subtotal function maps to the matching aggregation
// by ordinal.
Aggregation aggregation_from_subtotal_fn(SubtotalFn fn) {
  return static_cast<Aggregation>(static_cast<std::uint8_t>(fn));
}

}  // namespace

Expected<PivotResult, Error> evaluate(const PivotTable& table, const PivotCache& cache,
                                      const PivotLayoutOptions& options, const PivotFilterEnv& env) {
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
  // Grouping-derived fields (`databaseField="0"`, backed by a
  // `<fieldGroup>`) are not modelled by the reader: date grouping
  // (Years/Quarters/Months/...) sets no `PivotField::date_group`, and
  // number grouping has no representation in the model at all, so every
  // record's value for such a field is the reader's `Value::blank()`
  // placeholder. Evaluating a table that places one of these fields on
  // an axis would silently collapse that axis to a single blank item
  // instead of the groups Excel actually drew, so refuse outright
  // rather than render a wrong answer.
  for (std::size_t fi = 0; fi < table.fields().size(); ++fi) {
    if (table.fields()[fi].axis == PivotAxis::None || fi >= cache.fields().size()) {
      continue;
    }
    if (table.fields()[fi].date_group.has_value()) {
      continue;  // The caller (C API) has already supplied a grouping.
    }
    const PivotCacheField& cache_field = cache.fields()[fi];
    if (!cache_field.is_database_field && !cache_field.field_group_xml.empty()) {
      return make_error(FormulonErrorCode::kEvalPivotInvalid,
                        "pivot field is Excel date/number grouping, which this reader does not decode; refusing "
                        "rather than evaluating every record as blank",
                        "field=" + pivot_field_display_name(table.fields()[fi]) + " field_index=" + std::to_string(fi));
    }
  }
  for (std::size_t fi = 0; fi < table.fields().size(); ++fi) {
    const auto& date_group = table.fields()[fi].date_group;
    if (!date_group.has_value() || date_group->granularity != DateGrouping::Days) {
      continue;
    }
    if (date_group->interval_days == 0) {
      return make_error(FormulonErrorCode::kEvalPivotInvalid, "date grouping interval_days must be non-zero",
                        "field_index=" + std::to_string(fi));
    }
    // Compare authored bounds before flooring so the core model agrees with
    // the C API for inverted fractional windows (e.g. 0.9 > 0.1).
    if (date_group->start_serial.has_value() && date_group->end_serial.has_value() &&
        *date_group->start_serial > *date_group->end_serial) {
      return make_error(FormulonErrorCode::kEvalPivotInvalid, "date grouping start must not exceed end",
                        "field_index=" + std::to_string(fi));
    }
    if ((date_group->start_serial.has_value() && !is_valid_date_serial(*date_group->start_serial, env.date1904)) ||
        (date_group->end_serial.has_value() && !is_valid_date_serial(*date_group->end_serial, env.date1904))) {
      return make_error(FormulonErrorCode::kEvalPivotInvalid, "date grouping bound is outside the date serial domain",
                        "field_index=" + std::to_string(fi));
    }
  }

  // 2. Filter records.
  //
  // Everything the filter derives from the table alone — which items each
  // field hides, which field an axis filter names, which window a relative
  // period spans — is derived once here rather than per record. That is what
  // keeps the pass linear in the record count instead of scaling with the
  // item lists too, and it is also what makes the clock reading single: an
  // evaluation spanning midnight would otherwise filter its early records
  // against one day and its later ones against the next.
  PivotFilterEnv resolved_env = env;
  if (!resolved_env.pinned_now.has_value()) {
    resolved_env.pinned_now = date_time::host_civil_time();
  }
  const PreparedRecordFilter record_filter(table, cache, resolved_env);
  std::vector<std::size_t> surviving;
  surviving.reserve(cache.records().size());
  for (std::size_t i = 0; i < cache.records().size(); ++i) {
    if (record_filter.passes(cache.records()[i])) {
      surviving.push_back(i);
    }
  }

  // 3. Build hierarchies.
  //
  // A field whose `date_group` is set bucketises the cache value at
  // hierarchy-insertion time; we plumb the optional through `HierLevel`
  // so `insert_path` can call the bucketer without re-walking the table
  // metadata.
  //
  // `manual_order_cache` backs `HierLevel::manual_order` for every field
  // with `SortSpec::manual` set: a value->document-position map built
  // once per field (from `<items>`'s authored `x=` cache indices, not
  // from label text) and reused across every hierarchy level that field
  // occupies. It has to outlive `row_levels`/`col_levels`, which is why
  // it lives here rather than inside `level_for` itself; `unordered_map`
  // element references stay valid across further insertions, so a
  // pointer taken now stays good for the rest of this function.
  std::unordered_map<std::uint32_t, std::map<Value, std::size_t, ValueLess>> manual_order_cache;
  auto manual_order_for = [&](std::uint32_t fi) -> const std::map<Value, std::size_t, ValueLess>* {
    if (auto it = manual_order_cache.find(fi); it != manual_order_cache.end()) {
      return &it->second;
    }
    std::map<Value, std::size_t, ValueLess> positions;
    if (fi < table.fields().size() && fi < cache.fields().size()) {
      const std::vector<PivotItem>& items = table.fields()[fi].items;
      const std::vector<Value>& shared_items = cache.fields()[fi].shared_items;
      for (std::size_t i = 0; i < items.size(); ++i) {
        if (items[i].has_cache_index && items[i].cache_index < shared_items.size()) {
          positions.emplace(shared_items[items[i].cache_index], i);
        }
      }
    }
    return &manual_order_cache.emplace(fi, std::move(positions)).first->second;
  };
  // A `Days` field's auto Start/End resolves once per field from the
  // whole cache (matching Excel's Group dialog, unaffected by report
  // filters); every hierarchy level for that field reuses the result.
  std::unordered_map<std::uint32_t, PivotDateGroup> resolved_days_group_cache;
  auto resolved_date_group = [&](std::uint32_t fi, const PivotDateGroup& dg) -> const PivotDateGroup* {
    if (dg.granularity != DateGrouping::Days) {
      return &dg;
    }
    if (auto it = resolved_days_group_cache.find(fi); it != resolved_days_group_cache.end()) {
      return &it->second;
    }
    PivotDateGroup resolved = dg;
    bool have_bound = false;
    double data_min = 0.0;
    double data_max = 0.0;
    for (const PivotCacheRecord& rec : cache.records()) {
      const Value v = cell_value(cache, rec, fi);
      if (!v.is_number() || !is_valid_date_serial(v.as_number(), resolved_env.date1904)) {
        continue;  // Mirrors bucket_date's own valid-serial domain.
      }
      const double n = std::floor(v.as_number());
      if (!have_bound) {
        data_min = n;
        data_max = n;
        have_bound = true;
      } else {
        data_min = std::min(data_min, n);
        data_max = std::max(data_max, n);
      }
    }
    if (resolved.start_serial.has_value()) {
      resolved.start_serial = std::floor(*resolved.start_serial);
    } else {
      resolved.start_serial = have_bound ? data_min : 0.0;
    }
    if (!resolved.end_serial.has_value()) {
      const double auto_end = have_bound ? (data_max + 1.0) : *resolved.start_serial;
      resolved.end_serial = std::min(auto_end, last_valid_date_serial(resolved_env.date1904));
    } else {
      resolved.end_serial = std::floor(*resolved.end_serial);
    }
    return &resolved_days_group_cache.emplace(fi, resolved).first->second;
  };
  auto level_for = [&](std::uint32_t fi) -> HierLevel {
    const PivotDateGroup* dg = nullptr;
    if (fi < table.fields().size() && table.fields()[fi].date_group.has_value()) {
      dg = resolved_date_group(fi, *table.fields()[fi].date_group);
    }
    const bool manual = fi < table.fields().size() && table.fields()[fi].sort.manual;
    const bool ascending = fi >= table.fields().size() || table.fields()[fi].sort.ascending;
    const std::optional<ValueSortSpec> value_sort = resolve_value_sort(table, fi);
    return HierLevel{fi,
                     dg,
                     ascending,
                     value_sort.has_value() ? std::optional<std::uint32_t>(value_sort->field_index) : std::nullopt,
                     value_sort.has_value() ? std::optional<Aggregation>(value_sort->aggregation) : std::nullopt,
                     manual ? manual_order_for(fi) : nullptr};
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
      row_leaves_for_record[i] =
          insert_path(cache, row_levels, rec, surviving[i], row_tree, options.blank_item_label, resolved_env.date1904);
    }
    if (!col_levels.empty()) {
      col_leaves_for_record[i] =
          insert_path(cache, col_levels, rec, surviving[i], col_tree, options.blank_item_label, resolved_env.date1904);
    }
  }

  // A non-empty SortSpec::by_field orders siblings by the aggregate of the
  // named value field rather than by their display labels. Populate those
  // keys before finalising either hierarchy so leaf indices, values, and the
  // rendered axis all receive the same permutation.
  auto assign_value_sort_keys = [&](auto&& self, HierNode& node, const std::vector<HierLevel>& levels,
                                    std::size_t depth) -> void {
    if (depth >= levels.size()) {
      return;
    }
    const HierLevel& level = levels[depth];
    for (auto& [unused_key, child] : node.children) {
      (void)unused_key;
      if (level.value_sort_field.has_value() && level.value_sort_aggregation.has_value()) {
        std::vector<Value> values;
        append_record_field_values(cache, child.record_indices, *level.value_sort_field, values);
        child.value_sort_key = apply_aggregation(*level.value_sort_aggregation, values);
      }
      self(self, child, levels, depth + 1U);
    }
  };
  assign_value_sort_keys(assign_value_sort_keys, row_tree, row_levels, 0U);
  assign_value_sort_keys(assign_value_sort_keys, col_tree, col_levels, 0U);

  PivotResult result;
  std::vector<HierNode*> row_leaves;
  std::vector<HierNode*> col_leaves;
  finalize_hierarchy(row_tree, row_levels, 0U, result.rows, row_leaves, result.text_storage);
  finalize_hierarchy(col_tree, col_levels, 0U, result.cols, col_leaves, result.text_storage);
  resolve_page_selections(table, cache, options, result);

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
        result.values[r][c].push_back(aggregate_or_blank(df.aggregation, column, result));
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
        result.row_leaf_totals[r][df_idx] = aggregate_or_blank(df.aggregation, column, result);
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
        result.col_leaf_totals[c][df_idx] = aggregate_or_blank(df.aggregation, column, result);
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

    walk_subtotal_tree(row_tree, row_levels, table, stack_row_leaves,
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
                           for (std::size_t leaf_idx_iter = collected_start; leaf_idx_iter < leaves.size();
                                ++leaf_idx_iter) {
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
                           std::vector<std::vector<Value>> col_values(
                               col_leaf_count, std::vector<Value>(data_field_count, Value::blank()));
                           for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
                             const Aggregation agg = spec.has_value() ? *spec : table.data_fields()[df_idx].aggregation;
                             row_values[df_idx] = aggregate_or_blank(agg, df_columns[df_idx], result);
                             for (std::size_t c = 0; c < col_leaf_count; ++c) {
                               col_values[c][df_idx] = aggregate_or_blank(agg, df_columns_by_col[df_idx][c], result);
                             }
                           }
                           RowSubtotal subtotal;
                           subtotal.labels = labels;
                           subtotal.depth = static_cast<std::uint32_t>(depth);
                           subtotal.aggregation = spec;
                           subtotal.values = row_values;
                           subtotal.col_values = std::move(col_values);
                           row_subtotal_leaf_sets.emplace_back(
                               leaves.begin() + static_cast<std::ptrdiff_t>(collected_start), leaves.end());
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

    walk_subtotal_tree(
        col_tree, col_levels, table, stack_col_leaves,
        [&](const std::vector<std::string>& labels, std::size_t depth, std::size_t collected_start,
            const std::vector<std::size_t>& leaves) {
          // Mirror row-axis custom subtotal behavior: each selected
          // function emits a distinct subtotal column. An empty spec
          // retains the data field's own aggregation.
          std::vector<std::optional<Aggregation>> specs;
          if (depth < table.col_field_order().size()) {
            const std::uint32_t group_fi = table.col_field_order()[depth];
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
            ColSubtotal subtotal;
            subtotal.labels = labels;
            subtotal.depth = static_cast<std::uint32_t>(depth);
            subtotal.aggregation = spec;
            subtotal.values.assign(row_leaf_count, std::vector<Value>(data_field_count, Value::blank()));
            subtotal.total.assign(data_field_count, Value::blank());

            for (std::size_t r = 0; r < row_leaf_count; ++r) {
              for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
                const PivotDataField& df = table.data_fields()[df_idx];
                std::vector<Value> column;
                for (std::size_t leaf_idx_iter = collected_start; leaf_idx_iter < leaves.size(); ++leaf_idx_iter) {
                  const std::size_t leaf_idx = leaves[leaf_idx_iter];
                  append_bucket_field_values(cache, buckets, r, leaf_idx, df.field_index, column);
                }
                const Aggregation aggregation = spec.has_value() ? *spec : df.aggregation;
                subtotal.values[r][df_idx] = aggregate_or_blank(aggregation, column, result);
              }
            }
            // The subtotal's own total across every row leaf,
            // re-aggregated the same way `RowSubtotal::values`
            // is (never summed from `subtotal.values`, which
            // would be wrong for a non-additive aggregation).
            for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
              const PivotDataField& df = table.data_fields()[df_idx];
              std::vector<Value> column;
              for (std::size_t r = 0; r < row_leaf_count; ++r) {
                for (std::size_t leaf_idx_iter = collected_start; leaf_idx_iter < leaves.size(); ++leaf_idx_iter) {
                  const std::size_t leaf_idx = leaves[leaf_idx_iter];
                  append_bucket_field_values(cache, buckets, r, leaf_idx, df.field_index, column);
                }
              }
              const Aggregation aggregation = spec.has_value() ? *spec : df.aggregation;
              subtotal.total[df_idx] = aggregate_or_blank(aggregation, column, result);
            }
            col_subtotal_leaf_sets.emplace_back(leaves.begin() + static_cast<std::ptrdiff_t>(collected_start),
                                                leaves.end());
            result.col_subtotals.push_back(std::move(subtotal));
          }
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
          const ColSubtotal& col_subtotal = result.col_subtotals[cs];
          const Aggregation aggregation =
              col_subtotal.aggregation.has_value() ? *col_subtotal.aggregation : df.aggregation;
          row_subtotal.col_subtotal_values[cs][df_idx] = aggregate_or_blank(aggregation, column, result);
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
      result.grand_totals.push_back(aggregate_or_blank(df.aggregation, column, result));
    }
    if (!result.grand_totals.empty()) {
      result.grand_total = result.grand_totals[0];
    }
  }

  // 7. Value-axis filters (Top-N, GreaterThan, Between).
  apply_value_filters(table, cache, buckets, result, row_subtotal_leaf_sets, col_subtotal_leaf_sets);

  // 8. Show-values-as transforms.
  apply_show_values_as_transforms(table, cache, result, row_subtotal_leaf_sets, col_subtotal_leaf_sets);

  return result;
}

}  // namespace formulon::pivot
