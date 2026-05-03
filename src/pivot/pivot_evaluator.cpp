// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Pivot evaluator implementation. See header / §15.1.3 of the design
// corpus for the algorithm overview. The MVP path implemented here
// produces enough of a `PivotResult` that GETPIVOTDATA can resolve
// label/data tuples against the freshest evaluation snapshot.

#include "pivot/pivot_evaluator.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// ---------------------------------------------------------------------------
// Cross-kind ordering for hierarchy keys.
// ---------------------------------------------------------------------------
//
// Excel's default sort on a row/column field is ascending by the field's
// display value. The cache stores discrete values as the typed `Value`
// they came from, so we need a total order that:
//
//   * groups same-kind values together (so we don't interleave numbers
//     into strings);
//   * sorts numerically within `Number`;
//   * sorts case-sensitively within `Text` (good enough for MVP — the
//     primary oracle is ja-JP which Excel collates byte-wise on the
//     wire; the locale-aware compare will arrive with date grouping);
//   * is total across kinds so `std::map` stays well-formed.
//
// Kind ordering: Number < Bool < Text < Error < everything else. The
// cache only ever produces the first four for discrete fields, but the
// fall-through keeps the comparator total in case a future cache field
// surfaces an Array or similar.
int kind_rank(ValueKind k) noexcept {
  switch (k) {
    case ValueKind::Number:
      return 0;
    case ValueKind::Bool:
      return 1;
    case ValueKind::Text:
      return 2;
    case ValueKind::Error:
      return 3;
    case ValueKind::Blank:
      return 4;
    case ValueKind::Array:
      return 5;
    case ValueKind::Ref:
      return 6;
    case ValueKind::Lambda:
      return 7;
  }
  return 8;
}

bool value_less(const Value& a, const Value& b) noexcept {
  const int ra = kind_rank(a.kind());
  const int rb = kind_rank(b.kind());
  if (ra != rb) {
    return ra < rb;
  }
  switch (a.kind()) {
    case ValueKind::Number:
      return a.as_number() < b.as_number();
    case ValueKind::Bool:
      // false < true.
      return !a.as_boolean() && b.as_boolean();
    case ValueKind::Text:
      return a.as_text() < b.as_text();
    case ValueKind::Error:
      return static_cast<std::uint16_t>(a.as_error()) < static_cast<std::uint16_t>(b.as_error());
    default:
      // Same kind, no in-kind ordering defined — treat as equal.
      return false;
  }
}

struct ValueLess {
  bool operator()(const Value& a, const Value& b) const noexcept { return value_less(a, b); }
};

// ---------------------------------------------------------------------------
// Display / equality helpers for filter visibility.
// ---------------------------------------------------------------------------

// Renders a cache `Value` to the same string that would appear in a
// `PivotItem::name`. Mirrors what the OOXML reader would have written
// out: numbers as their canonical decimal, bools as TRUE/FALSE, errors
// as their `#…` token, blanks as the empty string. This is sufficient
// for matching against `PivotItem::name`, which is the only place we
// use it (manual-filter visibility check).
std::string display_string(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Blank:
      return std::string{};
    case ValueKind::Number: {
      // Excel renders integers without a trailing `.0`. We don't go
      // through the full number-format pipeline here: pivot item names
      // are produced by the OOXML reader from cache `<n v="…"/>` /
      // `<s v="…"/>` literals, and matching on the textual form is
      // robust enough for MVP. A more faithful renderer can replace
      // this when item-level filter parity is required.
      const double d = v.as_number();
      const auto i = static_cast<long long>(d);
      if (static_cast<double>(i) == d) {
        return std::to_string(i);
      }
      return std::to_string(d);
    }
    case ValueKind::Bool:
      return v.as_boolean() ? "TRUE" : "FALSE";
    case ValueKind::Text:
      return std::string{v.as_text()};
    case ValueKind::Error:
      return std::string{display_name(v.as_error())};
    default:
      return std::string{};
  }
}

// ---------------------------------------------------------------------------
// Cache-record value extraction.
// ---------------------------------------------------------------------------

// Pulls the effective `Value` for `(record, field)`. When the field is
// shared (`shared_items` non-empty), the record stores a `Number` index
// into `shared_items`; otherwise the record stores the value inline.
// Out-of-range record/field/index references collapse to `Blank` so the
// rest of the evaluator can stay branch-free; cache-record corruption
// is the OOXML reader's responsibility to surface.
Value cell_value(const PivotCache& cache, const PivotCacheRecord& record, std::size_t field_index) {
  if (field_index >= cache.fields().size() || field_index >= record.cells.size()) {
    return Value::blank();
  }
  const auto& field = cache.fields()[field_index];
  const Value& cell = record.cells[field_index];
  if (field.shared_items.empty()) {
    return cell;
  }
  if (!cell.is_number()) {
    return cell;  // Inline override (rare; Excel allows it).
  }
  const double idx = cell.as_number();
  if (idx < 0.0) {
    return Value::blank();
  }
  const auto i = static_cast<std::size_t>(idx);
  if (i >= field.shared_items.size()) {
    return Value::blank();
  }
  return field.shared_items[i];
}

// ---------------------------------------------------------------------------
// Aggregation primitives.
// ---------------------------------------------------------------------------
//
// Each aggregator ignores `Blank`. Errors propagate: the first error in
// the input dominates the output, matching `SUM(#DIV/0!, 1) -> #DIV/0!`.
// Booleans coerce numerically (TRUE=1, FALSE=0) for arithmetic
// aggregations; for `Count`, booleans are non-blank so they count, which
// also matches Excel's COUNTA on a boolean column.

// Returns the first `Value::error` found in `values`, or `std::nullopt`.
const Value* first_error(const std::vector<Value>& values) {
  for (const auto& v : values) {
    if (v.is_error()) {
      return &v;
    }
  }
  return nullptr;
}

// Numeric coercion for arithmetic aggregations. Booleans coerce; text
// is skipped (Excel's SUM/MAX/MIN over a Value column ignore text).
// `out` receives the coerced number on success.
bool coerce_arithmetic(const Value& v, double& out) noexcept {
  switch (v.kind()) {
    case ValueKind::Number:
      out = v.as_number();
      return true;
    case ValueKind::Bool:
      out = v.as_boolean() ? 1.0 : 0.0;
      return true;
    default:
      return false;
  }
}

Value AggregateSum(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double sum = 0.0;
  for (const auto& v : values) {
    double n = 0.0;
    if (coerce_arithmetic(v, n)) {
      sum += n;
    }
  }
  return Value::number(sum);
}

// Excel's pivot `Count` mirrors COUNTA: any non-blank cell counts,
// including text and booleans.
Value AggregateCount(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double count = 0.0;
  for (const auto& v : values) {
    if (!v.is_blank()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

// Excel's pivot `CountNumbers` mirrors COUNT: only numeric cells
// (booleans included, per Excel).
Value AggregateCountNumbers(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double count = 0.0;
  for (const auto& v : values) {
    if (v.is_number() || v.is_boolean()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

Value AggregateAverage(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double sum = 0.0;
  std::size_t n = 0;
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      sum += x;
      ++n;
    }
  }
  if (n == 0) {
    return Value::error(ErrorCode::Div0);
  }
  return Value::number(sum / static_cast<double>(n));
}

Value AggregateMax(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  bool seen = false;
  double best = -std::numeric_limits<double>::infinity();
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      if (!seen || x > best) {
        best = x;
        seen = true;
      }
    }
  }
  // Excel's pivot MAX over an empty/all-text group returns 0.
  return Value::number(seen ? best : 0.0);
}

Value AggregateMin(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  bool seen = false;
  double best = std::numeric_limits<double>::infinity();
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      if (!seen || x < best) {
        best = x;
        seen = true;
      }
    }
  }
  return Value::number(seen ? best : 0.0);
}

Value AggregateProduct(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  bool seen = false;
  double product = 1.0;
  for (const auto& v : values) {
    double x = 0.0;
    if (coerce_arithmetic(v, x)) {
      product *= x;
      seen = true;
    }
  }
  // Excel's pivot PRODUCT on an empty/all-text group returns 0.
  return Value::number(seen ? product : 0.0);
}

Value apply_aggregation(Aggregation agg, const std::vector<Value>& values) {
  switch (agg) {
    case Aggregation::Sum:
      return AggregateSum(values);
    case Aggregation::Count:
      return AggregateCount(values);
    case Aggregation::Average:
      return AggregateAverage(values);
    case Aggregation::Max:
      return AggregateMax(values);
    case Aggregation::Min:
      return AggregateMin(values);
    case Aggregation::Product:
      return AggregateProduct(values);
    case Aggregation::CountNumbers:
      return AggregateCountNumbers(values);
    case Aggregation::StdDev:
    case Aggregation::StdDevP:
    case Aggregation::Var:
    case Aggregation::VarP:
      // Deferred to a follow-up. Surface as #N/A so GETPIVOTDATA
      // returns a sentinel rather than silently aggregating to 0.
      return Value::error(ErrorCode::NA);
  }
  return Value::error(ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Filter (manual-only, via PivotItem::visible).
// ---------------------------------------------------------------------------

// True iff `record` survives the manual filter on every field that
// declares an `items` list. A field with empty `items` matches all
// values (Excel default — items[] is only authored when the user has
// hidden at least one value).
bool record_passes_manual_filter(const PivotTable& table, const PivotCache& cache, const PivotCacheRecord& record) {
  for (std::size_t fi = 0; fi < table.fields().size(); ++fi) {
    const PivotField& field = table.fields()[fi];
    if (field.items.empty()) {
      continue;
    }
    const Value v = cell_value(cache, record, fi);
    const std::string name = display_string(v);
    for (const PivotItem& item : field.items) {
      if (!item.visible && item.name == name) {
        return false;
      }
    }
  }
  return true;
}

// ---------------------------------------------------------------------------
// Hierarchy construction.
// ---------------------------------------------------------------------------
//
// We build the tree as a nested `std::map<Value, ...>` so the ordering
// emerges naturally from `ValueLess`. After all records are inserted we
// flatten into `RowHierarchyNode` / `ColHierarchyNode` and remember the
// leaf path that each surviving record lands on so the per-leaf
// aggregation pass can reuse the work without rewalking the tree.

struct HierNode {
  std::map<Value, HierNode, ValueLess> children;
  // Index into the flat leaf array assigned during finalisation. Leaves
  // only.
  std::size_t leaf_index = static_cast<std::size_t>(-1);
};

struct HierLevel {
  std::uint32_t field_index;  // Index into PivotTable::fields().
};

// Inserts `record` into `tree`, walking `levels`. Returns the leaf
// `HierNode*`. The caller assigns leaf indices in a second pass.
HierNode* insert_path(const PivotCache& cache, const std::vector<HierLevel>& levels, const PivotCacheRecord& record,
                      HierNode& root) {
  HierNode* cursor = &root;
  for (const HierLevel& level : levels) {
    const Value v = cell_value(cache, record, level.field_index);
    auto [it, _] = cursor->children.emplace(v, HierNode{});
    cursor = &it->second;
  }
  return cursor;
}

// Recursively flattens a hierarchy into `Node` form (template so we can
// produce both `RowHierarchyNode` and `ColHierarchyNode` from one
// implementation). On the way, assigns each leaf a dense index and
// pushes the corresponding `HierNode*` into `leaves` so a second pass
// can attach record indices.
template <class Node>
void finalize_hierarchy(HierNode& tree, std::vector<Node>& out, std::vector<HierNode*>& leaves) {
  if (tree.children.empty()) {
    return;
  }
  out.reserve(tree.children.size());
  for (auto& [key, child] : tree.children) {
    Node node;
    node.label = display_string(key);
    if (child.children.empty()) {
      child.leaf_index = leaves.size();
      leaves.push_back(&child);
    } else {
      finalize_hierarchy<Node>(child, node.children, leaves);
    }
    out.push_back(std::move(node));
  }
}

// ---------------------------------------------------------------------------
// Result-side text reification.
// ---------------------------------------------------------------------------
//
// `PivotResult::values` / `subtotals` / `grand_total` must outlive the
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
  std::vector<HierLevel> row_levels;
  row_levels.reserve(table.row_field_order().size());
  for (std::uint32_t fi : table.row_field_order()) {
    row_levels.push_back({fi});
  }
  std::vector<HierLevel> col_levels;
  col_levels.reserve(table.col_field_order().size());
  for (std::uint32_t fi : table.col_field_order()) {
    col_levels.push_back({fi});
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

  // Bucket surviving record indices by (row_leaf, col_leaf).
  // `[row_leaf][col_leaf]` -> indices into `cache.records()`.
  std::vector<std::vector<std::vector<std::size_t>>> buckets(row_leaf_count,
                                                             std::vector<std::vector<std::size_t>>(col_leaf_count));

  for (std::size_t i = 0; i < surviving.size(); ++i) {
    const std::size_t r = row_levels.empty() ? 0 : row_leaves_for_record[i]->leaf_index;
    const std::size_t c = col_levels.empty() ? 0 : col_leaves_for_record[i]->leaf_index;
    buckets[r][c].push_back(surviving[i]);
  }

  // 4. Aggregate per (row_leaf, col_leaf, data_field).
  result.values.assign(row_leaf_count, std::vector<std::vector<Value>>(col_leaf_count));
  for (std::size_t r = 0; r < row_leaf_count; ++r) {
    for (std::size_t c = 0; c < col_leaf_count; ++c) {
      result.values[r][c].reserve(data_field_count);
      const std::vector<std::size_t>& records = buckets[r][c];
      for (const PivotDataField& df : table.data_fields()) {
        std::vector<Value> column;
        column.reserve(records.size());
        for (std::size_t rec_idx : records) {
          column.push_back(cell_value(cache, cache.records()[rec_idx], df.field_index));
        }
        result.values[r][c].push_back(reify(apply_aggregation(df.aggregation, column), result));
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
  if (!row_levels.empty() && data_field_count > 0) {
    // A reusable closure that walks the tree depth-first and emits
    // subtotal rows where appropriate.
    std::vector<std::size_t> stack_row_leaves;  // current path's leaf indices
    std::vector<std::vector<Value>>& subtotals = result.subtotals;

    // Per-level cursor into row_levels keyed by depth.
    auto field_at_depth = [&](std::size_t depth) -> const PivotField* {
      if (depth >= row_levels.size()) {
        return nullptr;
      }
      const std::uint32_t fi = row_levels[depth].field_index;
      if (fi >= table.fields().size()) {
        return nullptr;
      }
      return &table.fields()[fi];
    };

    // Recursive walk implemented with an explicit stack to avoid lambda
    // recursion gymnastics. Each `Frame` owns iterators into its level.
    struct Frame {
      HierNode* node;
      std::map<Value, HierNode, ValueLess>::iterator it;
      std::size_t depth;
      std::size_t collected_start;  // index into `stack_row_leaves` at frame entry
    };

    std::vector<Frame> stack;
    stack.push_back({&row_tree, row_tree.children.begin(), 0, 0});

    while (!stack.empty()) {
      Frame& top = stack.back();
      if (top.it == top.node->children.end()) {
        // All children processed: emit a subtotal for this non-root,
        // non-leaf node when the field requests it.
        if (top.depth > 0 && !top.node->children.empty()) {
          const PivotField* field = field_at_depth(top.depth - 1);
          const bool wants_subtotal = field != nullptr && (field->subtotal_top || !field->subtotal_fns.empty());
          if (wants_subtotal) {
            // Aggregate over all surviving records whose row-leaf
            // index is in [collected_start .. stack_row_leaves.size()).
            std::vector<Value> row_values(data_field_count, Value::blank());
            for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
              const PivotDataField& df = table.data_fields()[df_idx];
              std::vector<Value> column;
              for (std::size_t leaf_idx_iter = top.collected_start; leaf_idx_iter < stack_row_leaves.size();
                   ++leaf_idx_iter) {
                const std::size_t leaf_idx = stack_row_leaves[leaf_idx_iter];
                for (std::size_t c = 0; c < col_leaf_count; ++c) {
                  for (std::size_t rec_idx : buckets[leaf_idx][c]) {
                    column.push_back(cell_value(cache, cache.records()[rec_idx], df.field_index));
                  }
                }
              }
              row_values[df_idx] = reify(apply_aggregation(df.aggregation, column), result);
            }
            subtotals.push_back(std::move(row_values));
          }
        }
        stack.pop_back();
        continue;
      }
      HierNode* child = &top.it->second;
      ++top.it;
      if (child->children.empty()) {
        // Leaf: contributes to its enclosing frames' subtotal.
        stack_row_leaves.push_back(child->leaf_index);
      } else {
        stack.push_back({child, child->children.begin(), top.depth + 1, stack_row_leaves.size()});
      }
    }
  }

  // 6. Grand total.
  //
  // MVP simplification: when grand totals are requested, surface the
  // aggregation of every surviving record using the FIRST data field's
  // aggregation. Excel's UI shows one grand-total cell per data field
  // along the totals strip; the single-Value slot on `PivotResult`
  // suffices for GETPIVOTDATA's `"Sum of Amount"` resolution path but
  // not for rendering the full strip.
  // TODO: extend `PivotResult::grand_total` to one Value per data field
  // once the renderer needs it (likely as `std::vector<Value>`).
  if ((table.grand_totals_rows() || table.grand_totals_cols()) && data_field_count > 0) {
    const PivotDataField& df = table.data_fields()[0];
    std::vector<Value> column;
    column.reserve(surviving.size());
    for (std::size_t rec_idx : surviving) {
      column.push_back(cell_value(cache, cache.records()[rec_idx], df.field_index));
    }
    result.grand_total = reify(apply_aggregation(df.aggregation, column), result);
  }

  return result;
}

}  // namespace formulon::pivot
