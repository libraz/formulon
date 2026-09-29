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

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <numeric>
#include <optional>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/hierarchy_builder.h"
#include "pivot/layout_generator.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
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

/// An empty row/column intersection is not the same as aggregating an empty
/// value sequence. Excel leaves a sparse pivot cell blank; aggregation
/// identities such as SUM's zero and AVERAGE's #DIV/0! only apply when a
/// group exists and its values themselves are empty/non-numeric.
Value aggregate_or_blank(Aggregation aggregation, const std::vector<Value>& values, PivotResult& result) {
  if (values.empty()) {
    return Value::blank();
  }
  return reify(apply_aggregation(aggregation, values), result);
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
      subtotal.values[df_idx] = aggregate_or_blank(df.aggregation, column, result);
    }
    for (std::size_t c = 0; c < subtotal.col_values.size() && c < surviving_col_leaves.size(); ++c) {
      for (std::size_t df_idx = 0; df_idx < data_field_count && df_idx < subtotal.col_values[c].size(); ++df_idx) {
        const PivotDataField& df = table.data_fields()[df_idx];
        std::vector<Value> column;
        append_leaf_set_field_values(cache, buckets, rows, {surviving_col_leaves[c]}, df.field_index, column);
        subtotal.col_values[c][df_idx] = aggregate_or_blank(df.aggregation, column, result);
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
    resolved_env.pinned_now = eval::date_time::host_civil_time();
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
  auto level_for = [&](std::uint32_t fi) -> HierLevel {
    const PivotDateGroup* dg = nullptr;
    if (fi < table.fields().size() && table.fields()[fi].date_group.has_value()) {
      dg = &*table.fields()[fi].date_group;
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
  finalize_hierarchy<RowHierarchyNode>(row_tree, row_levels, 0U, result.rows, row_leaves);
  finalize_hierarchy<ColHierarchyNode>(col_tree, col_levels, 0U, result.cols, col_leaves);
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
  //
  // Applied last so the pre-aggregation filter set has already shaped
  // `result.values`; the pruning here only drops surviving leaves. A filter
  // on the field at axis depth d keeps or drops whole depth-d items, scored
  // by the item's own aggregate over the records still standing and ranked
  // within each depth-(d-1) parent independently, as Excel does: Top 2 on an
  // inner field keeps two items under every outer item. The row/col tree is
  // then collapsed by dropping unkept leaves and any interior node whose
  // subtree becomes empty. A surviving subtotal is kept; a subtotal
  // whose group is filtered away entirely is dropped along with the
  // leaves it covered. Every total -- grand totals, per-leaf totals, and
  // surviving subtotal values -- is then recomputed from just the
  // surviving leaves (`ReaggregateTotalsAfterValueFilter`, below the
  // filter loop): Excel's own Top-N grand total covers only the visible
  // rows, not the pre-filter set (verified against
  // pivot_value_date_filters.xlsx / pivot_recurring_period_filter.xlsx).
  //
  // Both filter lists feed this pass. A slicer selection and a rule the
  // reader decoded out of `<filters>` prune identically once the latter
  // has recovered the axis a file does not spell out, so the authored
  // entries are projected onto the same shape rather than given a second
  // copy of the pruning logic.
  // The running-total target rides in the same `value` slot the item count
  // uses, since the three flavours share one dialog field.
  // Original-`buckets`-space indices of the row/col leaves still
  // standing, maintained in lock-step with `row_subtotal_leaf_sets` /
  // `col_subtotal_leaf_sets` (both recompacted by the same `keep` mask,
  // at the same point, by every filter application) so the
  // re-aggregation pass after the filter loop can map any leaf position
  // -- a plain leaf or a subtotal's leaf set -- back to `buckets`
  // coordinates regardless of how many filters ran or which axis each
  // one pruned.
  std::vector<std::size_t> surviving_row_leaves(row_leaf_count);
  std::iota(surviving_row_leaves.begin(), surviving_row_leaves.end(), std::size_t{0});
  std::vector<std::size_t> surviving_col_leaves(col_leaf_count);
  std::iota(surviving_col_leaves.begin(), surviving_col_leaves.end(), std::size_t{0});
  bool any_value_filter_applied = false;
  const auto filter_target = [](const PivotFilter& f) {
    if (const auto* as_double = std::get_if<double>(&f.value)) {
      return *as_double;
    }
    if (const auto* as_int = std::get_if<int>(&f.value)) {
      return static_cast<double>(*as_int);
    }
    return 0.0;
  };
  // `basis` selects how a top-N entry counts: `Items` uses the
  // item-count rule, the other two accumulate a running total. `top`
  // selects the ranking direction: `true` for Top N, `false` for Bottom
  // N. Both are only ever non-default for an authored entry, since the
  // dialog choice has no embedder-facing counterpart on `PivotFilter`.
  // `field_index` names the filtered field; one that is absent from the
  // filter's axis falls back to that axis's innermost field.
  const auto apply_value_filter = [&](const PivotFilter& f, std::optional<std::size_t> field_index,
                                      TopNBasis basis = TopNBasis::Items, bool top = true) {
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
    const std::vector<std::size_t>& axis_leaves = row_axis ? surviving_row_leaves : surviving_col_leaves;
    const std::vector<std::size_t>& cross_leaves = row_axis ? surviving_col_leaves : surviving_row_leaves;
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
      for (std::size_t i = groups[g].first_leaf; i < groups[g].first_leaf + groups[g].leaf_count && i < leaf_count;
           ++i) {
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
    const auto keep_or = build_grouped_value_filter_keep(f, basis, filter_target(f), top, groups, scores, leaf_count);
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
      compact_leaf_axis(result, keep, LeafAxis::Row, row_subtotal_leaf_sets, col_subtotal_leaf_sets);
      CompactSurvivingLeaves(surviving_row_leaves, keep);
    } else {
      prune_top_level(result.cols, keep);
      compact_leaf_axis(result, keep, LeafAxis::Col, row_subtotal_leaf_sets, col_subtotal_leaf_sets);
      CompactSurvivingLeaves(surviving_col_leaves, keep);
    }
    any_value_filter_applied = true;
  };

  for (const PivotFilter& f : table.active_filters()) {
    apply_value_filter(f, resolve_field_by_any_name(table, f.field_name));
  }
  for (const AuthoredValueFilter& authored : table.authored_value_filters()) {
    if (const auto projected = authored_value_filter_as_pivot_filter(table, authored)) {
      apply_value_filter(*projected, static_cast<std::size_t>(authored.field_index), authored.top_n_basis,
                         authored.top);
    }
  }
  if (any_value_filter_applied) {
    ReaggregateTotalsAfterValueFilter(table, cache, buckets, surviving_row_leaves, surviving_col_leaves,
                                      row_subtotal_leaf_sets, col_subtotal_leaf_sets, result);
  }

  // 8. Show-values-as transforms.
  apply_show_values_as_transforms(table, cache, result, row_subtotal_leaf_sets, col_subtotal_leaf_sets);

  return result;
}

}  // namespace formulon::pivot
