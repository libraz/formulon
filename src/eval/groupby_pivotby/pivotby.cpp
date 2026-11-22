
#include "eval/groupby_pivotby/pivotby.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/array_context.h"
#include "eval/error_origin.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/groupby_pivotby/common.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

/// PIVOTBY `relative_to` values.
enum RelativeTo : int { kColumnTotal, kRowTotal, kGrandTotal, kParentColumn, kParentRow };

/// Parsed and validated PIVOTBY arguments.
struct PivotArgs {
  const ArrayValue* row_fields = nullptr;
  const ArrayValue* col_fields = nullptr;
  const ArrayValue* values = nullptr;
  Aggregator agg_arg;
  int field_headers = kFieldHeadersAuto;
  int row_total_depth = 1;
  std::vector<int> row_sort_order;
  int col_total_depth = 1;
  std::vector<int> col_sort_order;
  HeaderLayout layout;
  std::vector<bool> include_row;
  int relative_to = kColumnTotal;
};

/// Row and column groups of the filtered data rows.
struct PivotGroups {
  std::vector<std::uint32_t> row_repr;
  std::vector<std::vector<std::uint32_t>> row_members;  // rows in each row-group
  std::vector<bool> row_is_error;
  std::vector<std::uint32_t> col_repr;
  std::vector<std::vector<std::uint32_t>> col_members;  // rows in each col-group
  std::vector<bool> col_is_error;
  // Per-row tags so we can later compute (row_g, col_g) intersections by
  // walking the data rows once.
  std::vector<std::size_t> row_tag;
  std::vector<std::size_t> col_tag;
};

/// Reads and validates arguments 0..10; on failure `err` holds the result.
bool parse_pivotby_args(const parser::AstNode& call, std::uint32_t arity, Arena& arena,
                        const FunctionRegistry& registry, const EvalContext& ctx, PivotArgs& a, Value& err) {
  // -- arg 0: row_fields ----------------------------------------------------
  a.row_fields = read_array_arg(call.as_call_arg(0), arena, registry, ctx, &err);
  if (a.row_fields == nullptr) {
    return false;
  }

  // -- arg 1: col_fields ----------------------------------------------------
  a.col_fields = read_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (a.col_fields == nullptr) {
    return false;
  }

  // -- arg 2: values --------------------------------------------------------
  a.values = read_array_arg(call.as_call_arg(2), arena, registry, ctx, &err);
  if (a.values == nullptr) {
    return false;
  }
  // A lone error in values is reported ahead of the other arguments.
  if (const Value* lone = lone_error_cell(*a.values)) {
    err = *lone;
    return false;
  }

  // Row-count consistency across all three rectangles. Mac Excel surfaces
  // `#VALUE!` when any pair has different row counts.
  if (a.row_fields->rows != a.values->rows || a.col_fields->rows != a.values->rows) {
    err = Value::error(ErrorCode::Value);
    return false;
  }
  if (a.row_fields->rows == 0U || a.row_fields->cols == 0U || a.col_fields->cols == 0U || a.values->cols == 0U) {
    note_unknown_error_origin();
    err = Value::error(ErrorCode::Calc);
    return false;
  }

  // -- arg 3: aggregator ----------------------------------------------------
  if (!resolve_aggregator(call.as_call_arg(3), arena, registry, ctx, &a.agg_arg, &err)) {
    return false;
  }

  // -- arg 4: field_headers ∈ {0,1,2,3}; omitted detects a header row ------
  a.field_headers = kFieldHeadersAuto;
  if (!read_field_headers(call, 4, arity, kFieldHeadersAuto, arena, registry, ctx, &a.field_headers, &err)) {
    return false;
  }
  a.field_headers = resolve_auto_field_headers(a.field_headers, *a.values);

  // -- arg 5: row_total_depth, default +1 ------------------------------------
  // The grand-total row (showing column totals) defaults to the BOTTOM of
  // the result. |depth| >= 2 adds subtotal rows for each row-key prefix of
  // length 1..|depth|-1; the depth cannot exceed the row-key count.
  a.row_total_depth = 1;
  if (!read_total_depth(call, 5, arity, 1, arena, registry, ctx, &a.row_total_depth, &err)) {
    return false;
  }
  if (!total_depth_fits(a.row_total_depth, a.row_fields->cols)) {
    err = Value::error(ErrorCode::Value);
    return false;
  }

  // -- arg 6: row_sort_order; omitted sorts the keys ascending -------------
  // Numbers the row_fields columns, then the value columns (ordering by the
  // row totals); negative descends.
  if (!read_optional_sort_spec(call, 6, arity, arena, registry, ctx, &a.row_sort_order, &err)) {
    return false;
  }

  // -- arg 7: col_total_depth, default 1 --------------------------------------
  // The grand-total column (showing row totals) defaults to the RIGHT of
  // the result. |depth| >= 2 adds subtotal blocks for each column-key prefix
  // of length 1..|depth|-1; the depth cannot exceed the column-key count.
  a.col_total_depth = 1;
  if (!read_total_depth(call, 7, arity, 1, arena, registry, ctx, &a.col_total_depth, &err)) {
    return false;
  }
  if (!total_depth_fits(a.col_total_depth, a.col_fields->cols)) {
    err = Value::error(ErrorCode::Value);
    return false;
  }

  // -- arg 8: col_sort_order; omitted sorts the keys ascending -------------
  // Excel pins the zero rejection on the row slot; this slot takes the same
  // signed-column-index domain, so the same rule is applied to both rather
  // than leaving one half of a symmetric argument pair accepting a value the
  // other rejects.
  if (!read_optional_sort_spec(call, 8, arity, arena, registry, ctx, &a.col_sort_order, &err)) {
    return false;
  }

  // Determine header row layout. Same as GROUPBY but the header / output
  // emission flags drive both the row-axis labels (left edge) and the
  // col-axis labels (top edge).
  // -- arg 9: filter_array --------------------------------------------------
  if (!read_layout_and_mask(call, 9, arity, a.field_headers, a.row_fields->rows, arena, registry, ctx, &a.layout,
                            &a.include_row, &err)) {
    return false;
  }

  // -- arg 10: relative_to ∈ {0..4}, default 0 (column totals) -------------
  static constexpr int kRelativeToValues[] = {kColumnTotal, kRowTotal, kGrandTotal, kParentColumn, kParentRow};
  a.relative_to = kColumnTotal;
  if (!read_optional_int_in_set(call, 10, arity, kColumnTotal, arena, registry, ctx, kRelativeToValues,
                                sizeof(kRelativeToValues) / sizeof(kRelativeToValues[0]), &a.relative_to, &err)) {
    return false;
  }
  return true;
}

/// Assigns every filtered data row to one row-group and one col-group. False
/// when either axis ends up with no group.
bool build_pivotby_groups(const PivotArgs& a, PivotGroups& g) {
  // -- Build row and column groups ----------------------------------------
  // For each filtered data row: assign the row to one row-group (by
  // row_fields key) and one col-group (by col_fields key). The (row_group,
  // col_group) pair is later used to aggregate the values column. We also
  // keep per-group full row-index lists for row totals and col totals.
  const std::uint32_t data_start_row = a.layout.data_start_row;
  const std::uint32_t data_row_count = a.layout.data_row_count;
  std::unordered_map<std::string, std::size_t> row_index;
  std::unordered_map<std::string, std::size_t> col_index;
  row_index.reserve(data_row_count);
  col_index.reserve(data_row_count);

  g.row_tag.assign(data_row_count, 0);
  g.col_tag.assign(data_row_count, 0);

  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    if (!a.include_row[i]) {
      continue;
    }
    const std::uint32_t row = data_start_row + i;
    g.row_tag[i] = find_or_add_group(*a.row_fields, row, &g.row_repr, &g.row_members, &g.row_is_error, &row_index);
    g.col_tag[i] = find_or_add_group(*a.col_fields, row, &g.col_repr, &g.col_members, &g.col_is_error, &col_index);
  }

  return !(g.row_repr.empty() || g.col_repr.empty());
}

/// Plans both axis sorts and applies the body-size cap; on failure `err` holds
/// the result.
bool plan_pivotby_sorts(const PivotArgs& a, const PivotGroups& g, GroupSort& row_plan, GroupSort& col_plan,
                        Value& err) {
  const std::size_t n_rows = g.row_repr.size();
  const std::size_t n_cols = g.col_repr.size();
  const std::uint32_t key_cols = a.row_fields->cols;
  const std::uint32_t col_levels = a.col_fields->cols;
  const std::uint32_t val_cols = a.values->cols;

  // Each sort order numbers its own axis's key columns first, then the value
  // columns, whose axis totals order the groups.
  if (!plan_group_sort(a.row_sort_order, key_cols, val_cols, false, &row_plan) ||
      !plan_group_sort(a.col_sort_order, col_levels, val_cols, false, &col_plan)) {
    err = Value::error(ErrorCode::Value);
    return false;
  }
  // The output holds every body cell, so a body past the cap is the output's
  // #NUM! before the body is built.
  if (static_cast<std::uint64_t>(n_rows) * n_cols * val_cols > kMaxDerivedArrayCells) {
    err = Value::error(ErrorCode::Num);
    return false;
  }
  return true;
}

/// Aggregation state of one PIVOTBY call: axis hierarchies, the leaf / total
/// cells and the ordered output slots.
struct PivotCalc {
  PivotCalc(const PivotArgs& a, const PivotGroups& g, Arena& arena_in, const FunctionRegistry& registry_in,
            const EvalContext& ctx_in)
      : row_fields(a.row_fields),
        col_fields(a.col_fields),
        values(a.values),
        agg_arg(a.agg_arg),
        include_row(a.include_row),
        data_row_count(a.layout.data_row_count),
        data_start_row(a.layout.data_start_row),
        relative_to(a.relative_to),
        row_total_depth(a.row_total_depth),
        col_total_depth(a.col_total_depth),
        row_repr(g.row_repr),
        row_members(g.row_members),
        row_is_error(g.row_is_error),
        col_repr(g.col_repr),
        col_members(g.col_members),
        col_is_error(g.col_is_error),
        row_tag(g.row_tag),
        col_tag(g.col_tag),
        arena(arena_in),
        registry(registry_in),
        ctx(ctx_in),
        n_rows(g.row_repr.size()),
        n_cols(g.col_repr.size()),
        key_cols(a.row_fields->cols),
        col_levels(a.col_fields->cols),
        val_cols(a.values->cols),
        emit_row_totals_col(a.col_total_depth != 0),
        emit_col_totals_row(a.row_total_depth != 0) {
    // Subtotals need every key prefix; a parent total (kParentColumn / kParentRow) needs
    // them on its axis only for a two-parameter aggregator.
    emit_row_subtotals = depth_has_subtotals(row_total_depth);
    emit_col_subtotals = depth_has_subtotals(col_total_depth);
    const bool parent_totals = agg_arg.arity >= 2U;
    row_hierarchy = build_axis_hierarchy(*row_fields, row_repr, row_members,
                                         emit_row_subtotals || (parent_totals && relative_to == kParentRow));
    col_hierarchy = build_axis_hierarchy(*col_fields, col_repr, col_members,
                                         emit_col_subtotals || (parent_totals && relative_to == kParentColumn));

    // A two-parameter aggregator's second argument is the rows of the total
    // `relative_to` names for the cell at (row, col): its column total, row
    // total, the grand total, or the parent column / row total. Bound once per
    // distinct total.
    all_rows = collect_included_rows(include_row, data_start_row);
  }

  const ArrayValue* row_fields;
  const ArrayValue* col_fields;
  const ArrayValue* values;
  const Aggregator& agg_arg;
  const std::vector<bool>& include_row;
  const std::uint32_t data_row_count;
  const std::uint32_t data_start_row;
  const int relative_to;
  const int row_total_depth;
  const int col_total_depth;
  const std::vector<std::uint32_t>& row_repr;
  const std::vector<std::vector<std::uint32_t>>& row_members;
  const std::vector<bool>& row_is_error;
  const std::vector<std::uint32_t>& col_repr;
  const std::vector<std::vector<std::uint32_t>>& col_members;
  const std::vector<bool>& col_is_error;
  const std::vector<std::size_t>& row_tag;
  const std::vector<std::size_t>& col_tag;
  Arena& arena;
  const FunctionRegistry& registry;
  const EvalContext& ctx;
  const std::size_t n_rows;
  const std::size_t n_cols;
  const std::uint32_t key_cols;
  const std::uint32_t col_levels;
  const std::uint32_t val_cols;
  const bool emit_row_totals_col;
  const bool emit_col_totals_row;
  bool emit_row_subtotals = false;
  bool emit_col_subtotals = false;
  AxisHierarchy row_hierarchy;
  AxisHierarchy col_hierarchy;
  std::vector<std::uint32_t> all_rows;
  std::map<std::array<std::size_t, 4>, Aggregator> agg_by_total;
  std::vector<std::vector<std::vector<Value>>> body;
  std::vector<std::vector<Value>> row_totals;
  std::vector<std::vector<Value>> col_totals;
  std::vector<Value> grand_totals;
  std::vector<std::size_t> row_order;
  std::vector<std::size_t> col_order;
  std::vector<AxisEntity> row_slots;
  std::vector<AxisEntity> col_slots;

  bool in_row_entity(std::uint32_t row, AxisEntity e) const {
    return axis_group_at(row_hierarchy, e.level, row_tag[row - data_start_row]) == e.index;
  }
  bool in_col_entity(std::uint32_t row, AxisEntity e) const {
    return axis_group_at(col_hierarchy, e.level, col_tag[row - data_start_row]) == e.index;
  }

  /// Null on allocation failure.
  const Aggregator* cell_agg(AxisEntity row, AxisEntity col);

  /// Cells where row node `re` meets column node `ce`: the precomputed leaf and
  /// total cells, else aggregated over the rows common to both. An empty
  /// intersection stays blank; `*failed` is set on allocation failure.
  std::vector<Value> node_cells(AxisEntity re, AxisEntity ce, bool* failed);

  /// Aggregates the body, row totals, column totals and grand totals. False on
  /// allocation failure.
  bool aggregate();

  /// Orders both axes and plans their output slots. False on allocation
  /// failure.
  bool order_axes(const GroupSort& row_plan, const GroupSort& col_plan);
};

const Aggregator* PivotCalc::cell_agg(AxisEntity row, AxisEntity col) {
  if (agg_arg.arity < 2U) {
    return &agg_arg;
  }
  AxisEntity row_total;
  AxisEntity col_total;
  switch (relative_to) {
    case kColumnTotal:
      col_total = col;
      break;
    case kRowTotal:
      row_total = row;
      break;
    case kGrandTotal:
      break;
    case kParentColumn:
      row_total = row;
      col_total = axis_parent(col_hierarchy, col);
      break;
    default:
      row_total = axis_parent(row_hierarchy, row);
      col_total = col;
      break;
  }
  const std::array<std::size_t, 4> key = {row_total.level, row_total.index, col_total.level, col_total.index};
  const auto found = agg_by_total.find(key);
  if (found != agg_by_total.end()) {
    return &found->second;
  }
  const std::vector<std::uint32_t>& col_rows = axis_entity_rows(col_hierarchy, col_members, all_rows, col_total);
  std::vector<std::uint32_t> rows;
  if (row_total.level != 0U) {
    for (const std::uint32_t r : col_rows) {
      if (in_row_entity(r, row_total)) {
        rows.push_back(r);
      }
    }
  }
  Aggregator bound;
  if (!bind_aggregator_totals(*values, row_total.level == 0U ? col_rows : rows, arena, agg_arg, &bound)) {
    return nullptr;
  }
  return &agg_by_total.emplace(key, std::move(bound)).first->second;
}

std::vector<Value> PivotCalc::node_cells(AxisEntity re, AxisEntity ce, bool* failed) {
  const bool row_leaf = re.level == row_hierarchy.keys;
  const bool col_leaf = ce.level == col_hierarchy.keys;
  if (row_leaf && col_leaf) {
    return body[re.index][ce.index];
  }
  if (row_leaf && ce.level == 0U) {
    return row_totals[re.index];
  }
  if (re.level == 0U && col_leaf) {
    return col_totals[ce.index];
  }
  if (re.level == 0U && ce.level == 0U) {
    return grand_totals;
  }
  std::vector<Value> cells(val_cols, placeholder_cell());
  // The common rows keep the order of the row-side node, or of the column
  // node when the row side is the whole axis.
  std::vector<std::uint32_t> common;
  const std::vector<std::uint32_t>* rows = nullptr;
  if (re.level == 0U) {
    rows = &axis_entity_rows(col_hierarchy, col_members, all_rows, ce);
  } else if (ce.level == 0U) {
    rows = &axis_entity_rows(row_hierarchy, row_members, all_rows, re);
  } else {
    for (const std::uint32_t r : axis_entity_rows(row_hierarchy, row_members, all_rows, re)) {
      if (in_col_entity(r, ce)) {
        common.push_back(r);
      }
    }
    rows = &common;
  }
  if (rows->empty()) {
    return cells;
  }
  const Aggregator* agg = cell_agg(re, ce);
  if (agg == nullptr) {
    *failed = true;
    return cells;
  }
  return aggregate_value_columns(*values, val_cols, *rows, *agg, arena, registry, ctx);
}

bool PivotCalc::aggregate() {
  const AxisEntity kGrandEntity;
  const auto row_leaf_entity = [&](std::size_t index) { return AxisEntity{row_hierarchy.keys, index}; };
  const auto col_leaf_entity = [&](std::size_t index) { return AxisEntity{col_hierarchy.keys, index}; };

  // -- Aggregate per (row_group, col_group, value_col) --------------------
  // For each (rg, cg) pair, build the intersection row list once and reuse
  // it across every value column. Per-cell error isolation: an aggregator
  // failure for one (rg, cg, v) tuple lands in that cell; the rest of the
  // body is still computed.
  body = std::vector<std::vector<std::vector<Value>>>(
      n_rows, std::vector<std::vector<Value>>(n_cols, std::vector<Value>(val_cols, placeholder_cell())));
  std::vector<std::vector<std::uint32_t>> buckets(n_rows * n_cols);
  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    if (include_row[i]) {
      buckets[row_tag[i] * n_cols + col_tag[i]].push_back(data_start_row + i);
    }
  }
  for (std::size_t rg = 0; rg < n_rows; ++rg) {
    for (std::size_t cg = 0; cg < n_cols; ++cg) {
      const std::vector<std::uint32_t>& intersection = buckets[rg * n_cols + cg];
      if (intersection.empty()) {
        // No data points at this intersection. Mac Excel surfaces an empty
        // (Blank) cell here — the aggregator is not called for empty
        // groups in the pivot body. Already initialized to Blank.
        continue;
      }
      const Aggregator* agg = cell_agg(row_leaf_entity(rg), col_leaf_entity(cg));
      if (agg == nullptr) {
        return false;
      }
      body[rg][cg] = aggregate_value_columns(*values, val_cols, intersection, *agg, arena, registry, ctx);
    }
  }

  // -- Compute row totals (one per (row group, value col)) ----------------
  // A "row total" aggregates an entire row group across all col groups,
  // per value column — it lives in the GRAND-TOTAL COLUMN block, whose
  // presence is governed by `col_total_depth`.
  row_totals = std::vector<std::vector<Value>>(n_rows, std::vector<Value>(val_cols, placeholder_cell()));
  if (emit_row_totals_col) {
    for (std::size_t rg = 0; rg < n_rows; ++rg) {
      const Aggregator* agg = cell_agg(row_leaf_entity(rg), kGrandEntity);
      if (agg == nullptr) {
        return false;
      }
      row_totals[rg] = aggregate_value_columns(*values, val_cols, row_members[rg], *agg, arena, registry, ctx);
    }
  }

  // -- Compute column totals (one per (col group, value col)) -------------
  // A "column total" aggregates an entire col group across all row groups,
  // per value column — it lives in the GRAND-TOTAL ROW, whose presence is
  // governed by `row_total_depth`.
  col_totals = std::vector<std::vector<Value>>(n_cols, std::vector<Value>(val_cols, placeholder_cell()));
  if (emit_col_totals_row) {
    for (std::size_t cg = 0; cg < n_cols; ++cg) {
      const Aggregator* agg = cell_agg(kGrandEntity, col_leaf_entity(cg));
      if (agg == nullptr) {
        return false;
      }
      col_totals[cg] = aggregate_value_columns(*values, val_cols, col_members[cg], *agg, arena, registry, ctx);
    }
  }

  // -- Compute grand totals (one per value col) ---------------------------
  // The grand total cells exist only when both axes emit totals.
  grand_totals = std::vector<Value>(val_cols, placeholder_cell());
  if (emit_row_totals_col && emit_col_totals_row) {
    const Aggregator* agg = cell_agg(kGrandEntity, kGrandEntity);
    if (agg == nullptr) {
      return false;
    }
    grand_totals = aggregate_value_columns(*values, val_cols, all_rows, *agg, arena, registry, ctx);
  }
  return true;
}

bool PivotCalc::order_axes(const GroupSort& row_plan, const GroupSort& col_plan) {
  const AxisEntity kGrandEntity;
  const auto row_leaf_entity = [&](std::size_t index) { return AxisEntity{row_hierarchy.keys, index}; };
  const auto col_leaf_entity = [&](std::size_t index) { return AxisEntity{col_hierarchy.keys, index}; };

  // -- Sort row and col groups -------------------------------------------
  // The first/only value column reduces to the axis total in
  // single-column-values scope.
  std::vector<const Aggregator*> row_axis_agg(n_rows);
  for (std::size_t rg = 0; rg < n_rows; ++rg) {
    row_axis_agg[rg] = cell_agg(row_leaf_entity(rg), kGrandEntity);
    if (row_axis_agg[rg] == nullptr) {
      return false;
    }
  }
  std::vector<const Aggregator*> col_axis_agg(n_cols);
  for (std::size_t cg = 0; cg < n_cols; ++cg) {
    col_axis_agg[cg] = cell_agg(kGrandEntity, col_leaf_entity(cg));
    if (col_axis_agg[cg] == nullptr) {
      return false;
    }
  }
  order_axis_groups(row_plan, *row_fields, row_repr, row_members, row_is_error, row_totals, emit_row_totals_col,
                    /*totals_offset=*/0U, /*hierarchical=*/true, *values, row_axis_agg, arena, registry, ctx,
                    &row_order);
  order_axis_groups(col_plan, *col_fields, col_repr, col_members, col_is_error, col_totals, emit_col_totals_row,
                    /*totals_offset=*/0U, /*hierarchical=*/true, *values, col_axis_agg, arena, registry, ctx,
                    &col_order);

  // Output order of both axes: leaf groups plus the subtotal nodes the depths
  // request.
  row_slots = plan_axis_slots(row_hierarchy, row_order, row_total_depth);
  col_slots = plan_axis_slots(col_hierarchy, col_order, col_total_depth);
  return true;
}

// Multi-column row_fields / col_fields / values are supported. Let
//   K = row_fields->cols, L = col_fields->cols, V = values->cols.
// The output rectangle has:
//   rows = L (col-axis label rows) + (output_emits_header ? 1 : 0) +
//          nR (one per row group) + (emit_col_totals_row ? 1 : 0)
//          [for the multi-column layout K>1 OR L>1 OR V>1]
//   rows = 1 (col-axis/header row) + nR + (emit_col_totals_row ? 1 : 0)
//          [for the merged singletons layout K==1 AND L==1 AND V==1;
//           the col-axis row and the header row collapse into a single
//           combined top row, always emitted -- the row_fields header
//           label in its corner cell is blank when output_emits_header
//           is false, but the col-axis labels always render, preserved
//           for backwards compatibility with the original single-column
//           impl]
//   cols = K + S * V + (emit_row_totals_col ? V : 0), where S is the
//          physical column-slot count (leaf groups plus one outer subtotal
//          slot per key-prefix subtotal when |col_total_depth| >= 2)
//
// Layout summary, in row order:
//   1. L col-axis label rows: cells [0..K-1] are blank; each physical
//      column slot occupies V tiled cells. Leaf slots hold col_fields'
//      keys, subtotal slots hold the outer key on level 0 and blanks on
//      inner levels; the optional grand-total block holds "Grand Total"
//      on the outermost level only and is blank on inner levels.
//   2. Optional header row: cells [0..K-1] hold the row_fields header
//      labels (or generated labels); each physical slot holds the values
//      header labels (V cells) tiled per slot; the optional grand-
//      total block holds the same V values-header tile.
//   3. Optional TOP grand-total row (when row_total_depth < 0): "Grand
//      Total" label, blanks across the row keys, col totals, then the
//      grand totals over each value col.
//   4. nR body rows: row keys, then per (col group, value col) cells,
//      then optional row totals per value col.
//   5. Optional BOTTOM grand-total row (when row_total_depth > 0): same
//      shape as TOP.
//
// Reference: Mac Excel ja-JP 16.109 layouts captured via xlwings-backed
// oracle goldens. Generated field/value labels follow the profile's locale.

/// Lays out the output rectangle from the aggregated state.
Value render_pivotby_output(PivotCalc& calc, const PivotArgs& a) {
  Arena& arena = calc.arena;
  const EvalContext& ctx = calc.ctx;
  const ArrayValue* row_fields = calc.row_fields;
  const ArrayValue* col_fields = calc.col_fields;
  const ArrayValue* values = calc.values;
  const int field_headers = a.field_headers;
  const HeaderLayout& layout = a.layout;
  const int row_total_depth = calc.row_total_depth;
  const int col_total_depth = calc.col_total_depth;
  const std::size_t n_cols = calc.n_cols;
  const std::uint32_t key_cols = calc.key_cols;
  const std::uint32_t col_levels = calc.col_levels;
  const std::uint32_t val_cols = calc.val_cols;
  const bool emit_row_totals_col = calc.emit_row_totals_col;
  const bool emit_col_totals_row = calc.emit_col_totals_row;
  const bool emit_row_subtotals = calc.emit_row_subtotals;
  const bool emit_col_subtotals = calc.emit_col_subtotals;
  const AxisHierarchy& row_hierarchy = calc.row_hierarchy;
  const AxisHierarchy& col_hierarchy = calc.col_hierarchy;
  const std::vector<std::uint32_t>& row_repr = calc.row_repr;
  const std::vector<std::uint32_t>& col_repr = calc.col_repr;
  const std::vector<std::size_t>& col_order = calc.col_order;
  const std::vector<AxisEntity>& row_slots = calc.row_slots;
  const std::vector<AxisEntity>& col_slots = calc.col_slots;
  const AxisEntity kGrandEntity;

  // -- Assemble output ----------------------------------------------------
  // The output has K row-key columns on the left, one V-wide block per
  // physical column slot in the middle, and an optional V-wide grand-total
  // block on either the LEFT (col_total_depth < 0) or RIGHT (positive).
  //
  // Vertically: L col-axis label rows, then optional header row, then
  // optional TOP totals row, then nR body rows, then optional BOTTOM
  // totals row.
  const bool grand_total_left = emit_row_totals_col && (col_total_depth < 0);

  // Column index helpers. The body block is laid out as
  //   [K key cols][optional V grand-total cols if left]
  //   [col_slots.size() * V blocks][optional V grand-total cols if right].
  const std::uint32_t body_block_start = key_cols + (grand_total_left ? val_cols : 0U);
  const std::uint32_t grand_total_block_start =
      grand_total_left ? key_cols : (key_cols + static_cast<std::uint32_t>(col_slots.size()) * val_cols);
  const std::uint32_t out_cols =
      key_cols + static_cast<std::uint32_t>(col_slots.size()) * val_cols + (emit_row_totals_col ? val_cols : 0U);

  // Resolve the values-header label cells (V wide) once. With
  // field_headers == 3 we copy values->cells[v] for v=0..V-1 (the
  // input's row 0). With field_headers == 2 we synth "Value <v+1>".
  // field_headers == 0 leaves these blank (output_emits_header is false).
  std::vector<Value> values_header_labels(val_cols, placeholder_cell());
  if (layout.output_emits_header) {
    for (std::uint32_t v = 0; v < val_cols; ++v) {
      values_header_labels[v] = header_label(arena, ctx, field_headers, HeaderField::kValue, values->cells, v);
    }
  }

  // Helper: render one of the L col-axis label rows. `level` is the
  // 0-based col_fields column index for this row.
  auto render_col_axis_row = [&](std::uint32_t level) {
    std::vector<Value> row(out_cols, placeholder_cell());
    // Cells [0..K-1] stay blank.
    for (std::size_t ci = 0; ci < col_slots.size(); ++ci) {
      const AxisEntity& slot = col_slots[ci];
      Value label = placeholder_cell();
      // A subtotal node restates its key prefix and leaves the deeper header
      // cells blank.
      if (level < slot.level) {
        label = col_fields
                    ->cells[static_cast<std::size_t>(axis_entity_repr_row(col_hierarchy, col_repr, slot)) * col_levels +
                            level];
      }
      const std::uint32_t base = body_block_start + static_cast<std::uint32_t>(ci) * val_cols;
      // Tile the label V times across the body cells for this col group.
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        row[base + v] = label;
      }
    }
    if (emit_row_totals_col) {
      // Outermost level (level == 0): the grand-total label tiled V times.
      // With nested subtotals Excel promotes this label to the hierarchy
      // form (総計 in ja-JP); inner levels remain blank.
      if (level == 0U) {
        const Value gt = Value::text(arena.intern(total_label(ctx, emit_col_subtotals)));
        for (std::uint32_t v = 0; v < val_cols; ++v) {
          row[grand_total_block_start + v] = gt;
        }
      }
    }
    return row;
  };

  // Helper: render the (single) header row when output_emits_header.
  auto render_header_row = [&]() {
    std::vector<Value> row(out_cols, placeholder_cell());
    // Row-fields header labels (cells [0..K-1]).
    for (std::uint32_t c = 0; c < key_cols; ++c) {
      row[c] = header_label(arena, ctx, field_headers, HeaderField::kRow, row_fields->cells, c);
    }
    // Values-header labels tiled V cells per physical column slot. Subtotal
    // blocks carry the same values-header tile as their leaf siblings.
    for (std::size_t ci = 0; ci < col_slots.size(); ++ci) {
      const std::uint32_t base = body_block_start + static_cast<std::uint32_t>(ci) * val_cols;
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        row[base + v] = values_header_labels[v];
      }
    }
    // Grand-total block: same V values-header tile.
    if (emit_row_totals_col) {
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        row[grand_total_block_start + v] = values_header_labels[v];
      }
    }
    return row;
  };

  auto render_col_fields_header_row = [&]() {
    std::vector<Value> row(out_cols, placeholder_cell());
    for (std::uint32_t level = 0; level < col_levels; ++level) {
      const std::uint32_t dst = body_block_start + level * val_cols;
      if (dst < out_cols) {
        row[dst] = header_label(arena, ctx, field_headers, HeaderField::kColumn, col_fields->cells, level);
      }
    }
    return row;
  };

  // Helper: render one output row for row node `re`: the whole-axis node is
  // the grand-total row, a subtotal node restates its key prefix.
  bool cell_failed = false;
  const auto render_node_row = [&](AxisEntity re) {
    std::vector<Value> row(out_cols, placeholder_cell());
    if (re.level == 0U) {
      row[0] = Value::text(arena.intern(total_label(ctx, emit_row_subtotals)));
    } else {
      const std::uint32_t repr_row = axis_entity_repr_row(row_hierarchy, row_repr, re);
      for (std::uint32_t c = 0; c < re.level; ++c) {
        row[c] = row_fields->cells[static_cast<std::size_t>(repr_row) * key_cols + c];
      }
    }
    for (std::size_t ci = 0; ci < col_slots.size(); ++ci) {
      const std::vector<Value> cells = calc.node_cells(re, col_slots[ci], &cell_failed);
      const std::uint32_t base = body_block_start + static_cast<std::uint32_t>(ci) * val_cols;
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        row[base + v] = cells[v];
      }
    }
    // Grand-total block: Mac Excel ja-JP only populates this strip for the
    // single-value-column layout; with val_cols > 1 it carries headers but
    // blank values.
    if (emit_row_totals_col && val_cols == 1U) {
      row[grand_total_block_start] = calc.node_cells(re, kGrandEntity, &cell_failed)[0];
    }
    return row;
  };

  // Without an output header, K==1, L==1 and V==1 use a compact layout whose
  // single top row carries the col-axis labels (and the grand-total label),
  // with a placeholder corner. Anything else, including every shape that
  // emits headers, uses the multi-column layout.
  const bool merged_single_col_layout =
      (key_cols == 1U && col_levels == 1U && val_cols == 1U && !layout.output_emits_header);

  // Helper: render the merged top row of the compact layout.
  auto render_merged_header_row = [&]() {
    std::vector<Value> row(out_cols, placeholder_cell());
    for (std::size_t ci = 0; ci < n_cols; ++ci) {
      const std::size_t cg = col_order[ci];
      const std::uint32_t out_col_idx = body_block_start + static_cast<std::uint32_t>(ci);
      row[out_col_idx] = col_fields->cells[static_cast<std::size_t>(col_repr[cg]) * col_levels];
    }
    if (emit_row_totals_col) {
      row[grand_total_block_start] = Value::text(arena.intern(grand_total_label(ctx)));
    }
    return row;
  };

  std::vector<std::vector<Value>> out_rows;
  out_rows.reserve(static_cast<std::size_t>(col_levels) + row_slots.size() + 3U);
  if (merged_single_col_layout) {
    // Merged single-column layout: the combined top row (col-axis labels,
    // plus the row_fields header label when output_emits_header) is
    // always emitted -- see the comment on merged_single_col_layout above.
    out_rows.push_back(render_merged_header_row());
  } else {
    // Multi-column layout: when output_emits_header, the column-field header
    // row comes first and the row/value header row follows the L col-axis
    // label rows.
    if (layout.output_emits_header) {
      out_rows.push_back(render_col_fields_header_row());
    }
    for (std::uint32_t level = 0; level < col_levels; ++level) {
      out_rows.push_back(render_col_axis_row(level));
    }
    if (layout.output_emits_header) {
      out_rows.push_back(render_header_row());
    }
  }
  // Optional TOP grand-total row.
  if (emit_col_totals_row && row_total_depth < 0) {
    out_rows.push_back(render_node_row(kGrandEntity));
  }
  // Body rows and row subtotals in output order.
  for (const AxisEntity& slot : row_slots) {
    out_rows.push_back(render_node_row(slot));
  }
  // Optional BOTTOM grand-total row.
  if (emit_col_totals_row && row_total_depth > 0) {
    out_rows.push_back(render_node_row(kGrandEntity));
  }
  if (cell_failed) {
    return Value::error(ErrorCode::Num);
  }

  return rows_to_array_value(out_rows, out_cols, arena);
}

}  // namespace

Value eval_pivotby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 4U || arity > 11U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();
  PivotArgs args;
  if (!parse_pivotby_args(call, arity, arena, registry, ctx, args, err)) {
    return err;
  }

  PivotGroups groups;
  if (!build_pivotby_groups(args, groups)) {
    return Value::error(ErrorCode::Value);
  }

  GroupSort row_plan;
  GroupSort col_plan;
  if (!plan_pivotby_sorts(args, groups, row_plan, col_plan, err)) {
    return err;
  }

  PivotCalc calc(args, groups, arena, registry, ctx);
  if (!calc.aggregate() || !calc.order_axes(row_plan, col_plan)) {
    return Value::error(ErrorCode::Num);
  }

  return render_pivotby_output(calc, args);
}

}  // namespace eval
}  // namespace formulon
