
#include "eval/groupby_pivotby/groupby.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

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

Value eval_groupby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 7U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();

  // -- arg 0: row_fields ----------------------------------------------------
  const ArrayValue* row_fields = read_array_arg(call.as_call_arg(0), arena, registry, ctx, &err);
  if (row_fields == nullptr) {
    return err;
  }

  // -- arg 1: values --------------------------------------------------------
  const ArrayValue* values = read_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (values == nullptr) {
    return err;
  }

  // Row-count consistency. Mac Excel surfaces #VALUE! when the two arrays
  // have different row counts.
  if (row_fields->rows != values->rows) {
    return Value::error(ErrorCode::Value);
  }
  if (row_fields->rows == 0U || row_fields->cols == 0U || values->cols == 0U) {
    return Value::error(ErrorCode::Calc);
  }

  // -- arg 2: aggregator ----------------------------------------------------
  Aggregator agg_arg;
  if (!resolve_aggregator(call.as_call_arg(2), arena, registry, ctx, &agg_arg, &err)) {
    return err;
  }

  // -- arg 3: field_headers ∈ {0,1,2,3}; omitted detects a header row ------
  int field_headers = kFieldHeadersAuto;
  if (!read_field_headers(call, 3, arity, kFieldHeadersAuto, arena, registry, ctx, &field_headers, &err)) {
    return err;
  }
  field_headers = resolve_auto_field_headers(field_headers, *values);

  // -- arg 4: total_depth ∈ {-2,-1,0,1,2}; default 1 (grand total last) ----
  int total_depth = 1;
  if (!read_total_depth(call, 4, arity, 1, arena, registry, ctx, &total_depth, &err)) {
    return err;
  }

  // -- arg 5: sort_order ----------------------------------------------------
  std::vector<int> sort_spec;
  if (!read_optional_sort_spec(call, 5, arity, arena, registry, ctx, &sort_spec, &err)) {
    return err;
  }

  // Determine header row layout. Inputs have a header row when
  // field_headers ∈ {1, 3}; outputs emit a header row when
  // field_headers ∈ {1, 2, 3}.
  // -- arg 6: filter_array --------------------------------------------------
  HeaderLayout layout;
  std::vector<bool> include_row;
  if (!read_layout_and_mask(call, 6, arity, field_headers, row_fields->rows, arena, registry, ctx, &layout,
                            &include_row, &err)) {
    return err;
  }
  const std::uint32_t data_start_row = layout.data_start_row;
  const std::uint32_t data_row_count = layout.data_row_count;

  // -- Build groups --------------------------------------------------------
  // Walk filtered data rows in input order. The canonical key folds each
  // text cell once and indexes the composite key, preserving first-occurrence
  // ordering while avoiding a linear scan of existing groups.
  //
  // Group representatives are stored as row indices (into `row_fields`) so
  // that subsequent equality checks can re-use the same column-walk path.
  std::vector<std::uint32_t> group_repr;               // row index of representative
  std::vector<std::vector<std::uint32_t>> group_rows;  // row indices in each group
  std::vector<bool> group_is_error;
  std::unordered_map<std::string, std::size_t> group_index;
  group_index.reserve(data_row_count);

  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    if (!include_row[i]) {
      continue;
    }
    find_or_add_group(*row_fields, data_start_row + i, &group_repr, &group_rows, &group_is_error, &group_index);
  }

  if (group_repr.empty()) {
    return Value::error(ErrorCode::Value);
  }
  // A two-parameter aggregator's second argument is the whole filtered column,
  // for group, subtotal and grand-total cells alike.
  const std::vector<std::uint32_t> all_rows = collect_included_rows(include_row, data_start_row);
  Aggregator agg;
  if (!bind_aggregator_totals(*values, all_rows, arena, agg_arg, &agg)) {
    return Value::error(ErrorCode::Num);
  }

  // -- Aggregate per group -------------------------------------------------
  // Output has `row_fields->cols` key columns followed by `values->cols`
  // aggregated columns. Per-group error isolation: each invocation's error
  // is captured into its cell; the rest of the result is still computed.
  const std::uint32_t key_cols = row_fields->cols;
  const std::uint32_t val_cols = values->cols;
  const std::uint32_t out_cols = key_cols + val_cols;

  std::vector<std::vector<Value>> agg_rows;
  agg_rows.reserve(group_repr.size());
  for (std::size_t g = 0; g < group_repr.size(); ++g) {
    std::vector<Value> row(out_cols, placeholder_cell());
    // Key columns: copy the representative's key cells verbatim.
    const std::uint32_t repr_row = group_repr[g];
    for (std::uint32_t c = 0; c < key_cols; ++c) {
      row[c] = row_fields->cells[static_cast<std::size_t>(repr_row) * key_cols + c];
    }
    // Value columns: invoke the aggregator per column with the group's
    // slice. Errors land in the cell; they do NOT short-circuit the result.
    for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
      const ArrayValue* slice = build_group_slice(*values, vc, group_rows[g], arena);
      if (slice == nullptr) {
        row[key_cols + vc] = Value::error(ErrorCode::Num);
        continue;
      }
      row[key_cols + vc] = invoke_aggregator_for_group(agg, vc, slice, arena, registry, ctx);
    }
    agg_rows.push_back(std::move(row));
  }

  // -- Sort ---------------------------------------------------------------
  // sort_order numbers the output columns, key columns first; omitted sorts
  // by the keys ascending (see `plan_group_sort`).
  GroupSort sort_plan;
  if (!plan_group_sort(sort_spec, key_cols, val_cols, &sort_plan)) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> value_keys;
  if (sort_plan.value_col != 0U) {
    value_keys.reserve(agg_rows.size());
    for (const std::vector<Value>& row : agg_rows) {
      value_keys.push_back(row[key_cols + sort_plan.value_col - 1U]);
    }
  }
  std::vector<std::size_t> order;
  order_groups(sort_plan, *row_fields, group_repr, group_is_error, value_keys, &order);

  // -- Compute grand total (if requested) ---------------------------------
  // The grand total aggregates over EVERY filtered data row. It is rendered
  // with the literal "Grand Total" label in the first key column and blank
  // cells in the remaining key columns.
  // A ±2 request adds one subtotal row per outer group — the first key
  // column alone — on top of the grand total. With a single key column the
  // outer level coincides with the groups themselves and every subtotal row
  // would just restate its group, so that shape stays on the ±1 layout.
  const bool emit_subtotals = (total_depth == 2 || total_depth == -2) && key_cols >= 2U;

  std::vector<Value> grand_total_row;
  bool emit_grand_total = (total_depth != 0);
  if (emit_grand_total) {
    grand_total_row.assign(out_cols, placeholder_cell());
    grand_total_row[0] =
        Value::text(arena.intern(emit_subtotals ? hierarchy_grand_total_label(ctx) : grand_total_label(ctx)));
    const std::vector<Value> totals =
        aggregate_value_columns(*values, val_cols, all_rows, agg, arena, registry, ctx, ErrorCode::Calc);
    for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
      grand_total_row[key_cols + vc] = totals[vc];
    }
  }
  // -- Nested subtotals (|total_depth| == 2) -------------------------------
  OuterGrouping hierarchy;
  std::vector<std::vector<Value>> subtotal_rows;
  std::vector<std::size_t> outer_order;
  std::vector<std::vector<std::size_t>> groups_of_outer;
  if (emit_subtotals) {
    hierarchy = build_outer_grouping(*row_fields, group_repr, group_rows);
    const std::size_t outer_count = hierarchy.repr_of_outer.size();
    // The hierarchy is primary and the sort applies within it: outer groups
    // follow the sorted order of their first member, and each outer group's
    // members keep their relative order from the same sort.
    groups_of_outer.resize(outer_count);
    std::vector<bool> outer_seen(outer_count, false);
    for (std::size_t g : order) {
      const std::size_t outer = hierarchy.outer_of_group[g];
      if (!outer_seen[outer]) {
        outer_seen[outer] = true;
        outer_order.push_back(outer);
      }
      groups_of_outer[outer].push_back(g);
    }
    subtotal_rows.reserve(outer_count);
    for (std::size_t o = 0; o < outer_count; ++o) {
      std::vector<Value> row(out_cols, placeholder_cell());
      // The subtotal row restates its outer key verbatim in the first key
      // column and leaves the inner key columns blank, so it reads as the
      // roll-up of the group rows directly above or below it.
      row[0] = row_fields->cells[static_cast<std::size_t>(hierarchy.repr_of_outer[o]) * key_cols];
      const std::vector<Value> totals = aggregate_value_columns(*values, val_cols, hierarchy.rows_of_outer[o], agg,
                                                                arena, registry, ctx, ErrorCode::Calc);
      for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
        row[key_cols + vc] = totals[vc];
      }
      subtotal_rows.push_back(std::move(row));
    }
  }

  // -- Assemble output ----------------------------------------------------
  std::vector<std::vector<Value>> out_rows;
  out_rows.reserve(agg_rows.size() + subtotal_rows.size() + 2U);

  // Header row (if requested).
  if (layout.output_emits_header) {
    std::vector<Value> header(out_cols, placeholder_cell());
    if (field_headers == 3) {
      // Inputs had a header row (row 0 of each input). Copy it verbatim.
      for (std::uint32_t c = 0; c < key_cols; ++c) {
        header[c] = row_fields->cells[c];
      }
      for (std::uint32_t c = 0; c < val_cols; ++c) {
        header[key_cols + c] = values->cells[c];
      }
    } else {
      // field_headers == 2: generated labels.
      for (std::uint32_t c = 0; c < key_cols; ++c) {
        header[c] = Value::text(arena.intern(row_field_label(ctx, c + 1U)));
      }
      for (std::uint32_t c = 0; c < val_cols; ++c) {
        header[key_cols + c] = Value::text(arena.intern(value_label(ctx, c + 1U)));
      }
    }
    emit_row(&out_rows, header);
  }

  // Grand total at top (negative total_depth).
  if (emit_grand_total && total_depth < 0) {
    emit_row(&out_rows, grand_total_row);
  }

  // Per-group rows in sorted order. With subtotals each outer group's rows
  // are bracketed by its subtotal row, placed on the same side of the block
  // as the grand total.
  if (emit_subtotals) {
    for (std::size_t o : outer_order) {
      if (total_depth < 0) {
        emit_row(&out_rows, subtotal_rows[o]);
      }
      for (std::size_t g : groups_of_outer[o]) {
        emit_row(&out_rows, agg_rows[g]);
      }
      if (total_depth > 0) {
        emit_row(&out_rows, subtotal_rows[o]);
      }
    }
  } else {
    for (std::size_t i : order) {
      emit_row(&out_rows, agg_rows[i]);
    }
  }

  // Grand total at bottom (positive total_depth).
  if (emit_grand_total && total_depth > 0) {
    emit_row(&out_rows, grand_total_row);
  }

  return rows_to_array_value(out_rows, out_cols, arena);
}

}  // namespace eval
}  // namespace formulon
