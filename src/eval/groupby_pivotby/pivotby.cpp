
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

// Orders one axis's groups by `plan` (see `plan_group_sort`). A value sort
// reads each group's axis total for that value column: `totals` when the
// layout emits them, otherwise aggregated here with the group's aggregator.
// The axis is a hierarchy, so a value sort ranks each key prefix first.
void order_axis_groups(const GroupSort& plan, const ArrayValue& keys, const std::vector<std::uint32_t>& repr,
                       const std::vector<std::vector<std::uint32_t>>& members, const std::vector<bool>& is_error,
                       const std::vector<std::vector<Value>>& totals, bool totals_emitted, const ArrayValue& values,
                       const std::vector<const Aggregator*>& agg_of_group, Arena& arena,
                       const FunctionRegistry& registry, const EvalContext& ctx, std::vector<std::size_t>* out_order) {
  // One sort key per group, aggregated before the sort so the comparator
  // never invokes the aggregator.
  std::vector<Value> value_keys;
  std::vector<OuterSortLevel> outer_levels;
  if (plan.value_col != 0U) {
    const std::uint32_t v = plan.value_col - 1U;
    outer_levels = build_outer_sort_levels(keys, repr, members, values, v, agg_of_group, arena, registry, ctx);
    value_keys.reserve(repr.size());
    for (std::size_t g = 0; g < repr.size(); ++g) {
      if (totals_emitted) {
        value_keys.push_back(totals[g][v]);
        continue;
      }
      const ArrayValue* slice = build_group_slice(values, v, members[g], arena);
      value_keys.push_back((slice != nullptr)
                               ? invoke_aggregator_for_group(*agg_of_group[g], v, slice, arena, registry, ctx)
                               : Value::error(ErrorCode::Num));
    }
  }
  order_groups(plan, keys, repr, is_error, value_keys, outer_levels, out_order);
}

}  // namespace

Value eval_pivotby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 4U || arity > 11U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();

  // -- arg 0: row_fields ----------------------------------------------------
  const ArrayValue* row_fields = read_array_arg(call.as_call_arg(0), arena, registry, ctx, &err);
  if (row_fields == nullptr) {
    return err;
  }

  // -- arg 1: col_fields ----------------------------------------------------
  const ArrayValue* col_fields = read_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (col_fields == nullptr) {
    return err;
  }

  // -- arg 2: values --------------------------------------------------------
  const ArrayValue* values = read_array_arg(call.as_call_arg(2), arena, registry, ctx, &err);
  if (values == nullptr) {
    return err;
  }

  // Row-count consistency across all three rectangles. Mac Excel surfaces
  // `#VALUE!` when any pair has different row counts.
  if (row_fields->rows != values->rows || col_fields->rows != values->rows) {
    return Value::error(ErrorCode::Value);
  }
  if (row_fields->rows == 0U || row_fields->cols == 0U || col_fields->cols == 0U || values->cols == 0U) {
    note_unknown_error_origin();
    return Value::error(ErrorCode::Calc);
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

  // -- arg 3: aggregator ----------------------------------------------------
  Aggregator agg_arg;
  if (!resolve_aggregator(call.as_call_arg(3), arena, registry, ctx, &agg_arg, &err)) {
    return err;
  }

  // -- arg 4: field_headers ∈ {0,1,2,3}; omitted detects a header row ------
  int field_headers = kFieldHeadersAuto;
  if (!read_field_headers(call, 4, arity, kFieldHeadersAuto, arena, registry, ctx, &field_headers, &err)) {
    return err;
  }
  field_headers = resolve_auto_field_headers(field_headers, *values);

  // -- arg 5: row_total_depth, default +1 ------------------------------------
  // The grand-total row (showing column totals) defaults to the BOTTOM of
  // the result. |depth| >= 2 adds subtotal rows for each row-key prefix of
  // length 1..|depth|-1; the depth cannot exceed the row-key count.
  int row_total_depth = 1;
  if (!read_total_depth(call, 5, arity, 1, arena, registry, ctx, &row_total_depth, &err)) {
    return err;
  }
  if (!total_depth_fits(row_total_depth, row_fields->cols)) {
    return Value::error(ErrorCode::Value);
  }

  // -- arg 6: row_sort_order; omitted sorts the keys ascending -------------
  // Numbers the row_fields columns, then the value columns (ordering by the
  // row totals); negative descends.
  std::vector<int> row_sort_order;
  if (!read_optional_sort_spec(call, 6, arity, arena, registry, ctx, &row_sort_order, &err)) {
    return err;
  }

  // -- arg 7: col_total_depth, default 1 --------------------------------------
  // The grand-total column (showing row totals) defaults to the RIGHT of
  // the result. |depth| >= 2 adds subtotal blocks for each column-key prefix
  // of length 1..|depth|-1; the depth cannot exceed the column-key count.
  int col_total_depth = 1;
  if (!read_total_depth(call, 7, arity, 1, arena, registry, ctx, &col_total_depth, &err)) {
    return err;
  }
  if (!total_depth_fits(col_total_depth, col_fields->cols)) {
    return Value::error(ErrorCode::Value);
  }

  // -- arg 8: col_sort_order; omitted sorts the keys ascending -------------
  // Excel pins the zero rejection on the row slot; this slot takes the same
  // signed-column-index domain, so the same rule is applied to both rather
  // than leaving one half of a symmetric argument pair accepting a value the
  // other rejects.
  std::vector<int> col_sort_order;
  if (!read_optional_sort_spec(call, 8, arity, arena, registry, ctx, &col_sort_order, &err)) {
    return err;
  }

  // Determine header row layout. Same as GROUPBY but the header / output
  // emission flags drive both the row-axis labels (left edge) and the
  // col-axis labels (top edge).
  // -- arg 9: filter_array --------------------------------------------------
  HeaderLayout layout;
  std::vector<bool> include_row;
  if (!read_layout_and_mask(call, 9, arity, field_headers, row_fields->rows, arena, registry, ctx, &layout,
                            &include_row, &err)) {
    return err;
  }
  const std::uint32_t data_start_row = layout.data_start_row;
  const std::uint32_t data_row_count = layout.data_row_count;

  // -- arg 10: relative_to ∈ {0..4}, default 0 (column totals) -------------
  static constexpr int kRelativeTo[] = {0, 1, 2, 3, 4};
  int relative_to = 0;
  if (!read_optional_int_in_set(call, 10, arity, 0, arena, registry, ctx, kRelativeTo,
                                sizeof(kRelativeTo) / sizeof(kRelativeTo[0]), &relative_to, &err)) {
    return err;
  }

  // -- Build row and column groups ----------------------------------------
  // For each filtered data row: assign the row to one row-group (by
  // row_fields key) and one col-group (by col_fields key). The (row_group,
  // col_group) pair is later used to aggregate the values column. We also
  // keep per-group full row-index lists for row totals and col totals.
  std::vector<std::uint32_t> row_repr;
  std::vector<std::vector<std::uint32_t>> row_members;  // rows in each row-group
  std::vector<bool> row_is_error;
  std::vector<std::uint32_t> col_repr;
  std::vector<std::vector<std::uint32_t>> col_members;  // rows in each col-group
  std::vector<bool> col_is_error;
  std::unordered_map<std::string, std::size_t> row_index;
  std::unordered_map<std::string, std::size_t> col_index;
  row_index.reserve(data_row_count);
  col_index.reserve(data_row_count);

  // Per-row tags so we can later compute (row_g, col_g) intersections by
  // walking the data rows once.
  std::vector<std::size_t> row_tag(data_row_count, 0);
  std::vector<std::size_t> col_tag(data_row_count, 0);

  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    if (!include_row[i]) {
      continue;
    }
    const std::uint32_t row = data_start_row + i;
    row_tag[i] = find_or_add_group(*row_fields, row, &row_repr, &row_members, &row_is_error, &row_index);
    col_tag[i] = find_or_add_group(*col_fields, row, &col_repr, &col_members, &col_is_error, &col_index);
  }

  if (row_repr.empty() || col_repr.empty()) {
    return Value::error(ErrorCode::Value);
  }

  const std::size_t n_rows = row_repr.size();
  const std::size_t n_cols = col_repr.size();
  const std::uint32_t key_cols = row_fields->cols;
  const std::uint32_t col_levels = col_fields->cols;
  const std::uint32_t val_cols = values->cols;

  // Subtotals need every key prefix; a parent total (relative_to 3 / 4) needs
  // them on its axis only for a two-parameter aggregator.
  const bool emit_row_subtotals = row_total_depth >= 2 || row_total_depth <= -2;
  const bool emit_col_subtotals = col_total_depth >= 2 || col_total_depth <= -2;
  const bool parent_totals = agg_arg.arity >= 2U;
  const AxisHierarchy row_hierarchy = build_axis_hierarchy(*row_fields, row_repr, row_members,
                                                           emit_row_subtotals || (parent_totals && relative_to == 4));
  const AxisHierarchy col_hierarchy = build_axis_hierarchy(*col_fields, col_repr, col_members,
                                                           emit_col_subtotals || (parent_totals && relative_to == 3));

  // A two-parameter aggregator's second argument is the rows of the total
  // `relative_to` names for the cell at (row, col): its column total, row
  // total, the grand total, or the parent column / row total. Bound once per
  // distinct total.
  const std::vector<std::uint32_t> all_rows = collect_included_rows(include_row, data_start_row);
  const auto in_row_entity = [&](std::uint32_t row, AxisEntity e) {
    return axis_group_at(row_hierarchy, e.level, row_tag[row - data_start_row]) == e.index;
  };
  const auto in_col_entity = [&](std::uint32_t row, AxisEntity e) {
    return axis_group_at(col_hierarchy, e.level, col_tag[row - data_start_row]) == e.index;
  };
  std::map<std::array<std::size_t, 4>, Aggregator> agg_by_total;
  // Null on allocation failure.
  const auto cell_agg = [&](AxisEntity row, AxisEntity col) -> const Aggregator* {
    if (agg_arg.arity < 2U) {
      return &agg_arg;
    }
    AxisEntity row_total;
    AxisEntity col_total;
    switch (relative_to) {
      case 0:
        col_total = col;
        break;
      case 1:
        row_total = row;
        break;
      case 2:
        break;
      case 3:
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
  };
  const AxisEntity kGrandEntity;
  const auto row_leaf_entity = [&](std::size_t index) { return AxisEntity{row_hierarchy.keys, index}; };
  const auto col_leaf_entity = [&](std::size_t index) { return AxisEntity{col_hierarchy.keys, index}; };

  // -- Aggregate per (row_group, col_group, value_col) --------------------
  // For each (rg, cg) pair, build the intersection row list once and reuse
  // it across every value column. Per-cell error isolation: an aggregator
  // failure for one (rg, cg, v) tuple lands in that cell; the rest of the
  // body is still computed.
  std::vector<std::vector<std::vector<Value>>> body(
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
        return Value::error(ErrorCode::Num);
      }
      body[rg][cg] = aggregate_value_columns(*values, val_cols, intersection, *agg, arena, registry, ctx);
    }
  }

  // -- Compute row totals (one per (row group, value col)) ----------------
  // A "row total" aggregates an entire row group across all col groups,
  // per value column — it lives in the GRAND-TOTAL COLUMN block, whose
  // presence is governed by `col_total_depth`.
  const bool emit_row_totals_col = (col_total_depth != 0);
  std::vector<std::vector<Value>> row_totals(n_rows, std::vector<Value>(val_cols, placeholder_cell()));
  if (emit_row_totals_col) {
    for (std::size_t rg = 0; rg < n_rows; ++rg) {
      const Aggregator* agg = cell_agg(row_leaf_entity(rg), kGrandEntity);
      if (agg == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      row_totals[rg] = aggregate_value_columns(*values, val_cols, row_members[rg], *agg, arena, registry, ctx);
    }
  }

  // -- Compute column totals (one per (col group, value col)) -------------
  // A "column total" aggregates an entire col group across all row groups,
  // per value column — it lives in the GRAND-TOTAL ROW, whose presence is
  // governed by `row_total_depth`.
  const bool emit_col_totals_row = (row_total_depth != 0);
  std::vector<std::vector<Value>> col_totals(n_cols, std::vector<Value>(val_cols, placeholder_cell()));
  if (emit_col_totals_row) {
    for (std::size_t cg = 0; cg < n_cols; ++cg) {
      const Aggregator* agg = cell_agg(kGrandEntity, col_leaf_entity(cg));
      if (agg == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      col_totals[cg] = aggregate_value_columns(*values, val_cols, col_members[cg], *agg, arena, registry, ctx);
    }
  }

  // -- Compute grand totals (one per value col) ---------------------------
  // The grand total cells exist only when both axes emit totals.
  std::vector<Value> grand_totals(val_cols, placeholder_cell());
  if (emit_row_totals_col && emit_col_totals_row) {
    const Aggregator* agg = cell_agg(kGrandEntity, kGrandEntity);
    if (agg == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    grand_totals = aggregate_value_columns(*values, val_cols, all_rows, *agg, arena, registry, ctx);
  }

  // -- Sort row and col groups -------------------------------------------
  // The first/only value column reduces to the axis total in
  // single-column-values scope.
  // Each sort order numbers its own axis's key columns first, then the value
  // columns, whose axis totals order the groups.
  GroupSort row_plan;
  GroupSort col_plan;
  if (!plan_group_sort(row_sort_order, key_cols, val_cols, false, &row_plan) ||
      !plan_group_sort(col_sort_order, col_levels, val_cols, false, &col_plan)) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<const Aggregator*> row_axis_agg(n_rows);
  for (std::size_t rg = 0; rg < n_rows; ++rg) {
    row_axis_agg[rg] = cell_agg(row_leaf_entity(rg), kGrandEntity);
    if (row_axis_agg[rg] == nullptr) {
      return Value::error(ErrorCode::Num);
    }
  }
  std::vector<const Aggregator*> col_axis_agg(n_cols);
  for (std::size_t cg = 0; cg < n_cols; ++cg) {
    col_axis_agg[cg] = cell_agg(kGrandEntity, col_leaf_entity(cg));
    if (col_axis_agg[cg] == nullptr) {
      return Value::error(ErrorCode::Num);
    }
  }
  std::vector<std::size_t> row_order;
  order_axis_groups(row_plan, *row_fields, row_repr, row_members, row_is_error, row_totals, emit_row_totals_col,
                    *values, row_axis_agg, arena, registry, ctx, &row_order);
  std::vector<std::size_t> col_order;
  order_axis_groups(col_plan, *col_fields, col_repr, col_members, col_is_error, col_totals, emit_col_totals_row,
                    *values, col_axis_agg, arena, registry, ctx, &col_order);

  // Output order of both axes: leaf groups plus the subtotal nodes the depths
  // request.
  const std::vector<AxisEntity> row_slots = plan_axis_slots(row_hierarchy, row_order, row_total_depth);
  const std::vector<AxisEntity> col_slots = plan_axis_slots(col_hierarchy, col_order, col_total_depth);

  // Cells where row node `re` meets column node `ce`: the precomputed leaf and
  // total cells, else aggregated over the rows common to both. An empty
  // intersection stays blank; `*failed` is set on allocation failure.
  const auto node_cells = [&](AxisEntity re, AxisEntity ce, bool* failed) -> std::vector<Value> {
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
  };

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
      const std::vector<Value> cells = node_cells(re, col_slots[ci], &cell_failed);
      const std::uint32_t base = body_block_start + static_cast<std::uint32_t>(ci) * val_cols;
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        row[base + v] = cells[v];
      }
    }
    // Grand-total block: Mac Excel ja-JP only populates this strip for the
    // single-value-column layout; with val_cols > 1 it carries headers but
    // blank values.
    if (emit_row_totals_col && val_cols == 1U) {
      row[grand_total_block_start] = node_cells(re, kGrandEntity, &cell_failed)[0];
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

}  // namespace eval
}  // namespace formulon
