
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
  if (arity < 3U || arity > 8U) {
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

  // -- arg 4: total_depth, an integer up to the key-column count; default 1 -
  int total_depth = 1;
  if (!read_total_depth(call, 4, arity, 1, arena, registry, ctx, &total_depth, &err)) {
    return err;
  }

  // -- arg 5: sort_order ----------------------------------------------------
  std::vector<int> sort_spec;
  if (!read_optional_sort_spec(call, 5, arity, arena, registry, ctx, &sort_spec, &err)) {
    return err;
  }

  // -- arg 6: filter_array --------------------------------------------------
  // Resolves the header layout (inputs have a header row when field_headers
  // is 1 or 3; outputs emit one when it is 1, 2 or 3) and reads the mask.
  HeaderLayout layout;
  std::vector<bool> include_row;
  if (!read_layout_and_mask(call, 6, arity, field_headers, row_fields->rows, arena, registry, ctx, &layout,
                            &include_row, &err)) {
    return err;
  }
  const std::uint32_t data_start_row = layout.data_start_row;
  const std::uint32_t data_row_count = layout.data_row_count;

  // -- arg 7: field_relationship; 0 hierarchy (default), 1 flat table -------
  // A flat table has no levels to subtotal, and the depth cannot exceed the key columns.
  static constexpr int kFieldRelationships[] = {0, 1};
  int field_relationship = 0;
  if (!read_optional_int_in_set(call, 7, arity, 0, arena, registry, ctx, kFieldRelationships,
                                sizeof(kFieldRelationships) / sizeof(kFieldRelationships[0]), &field_relationship,
                                &err)) {
    return err;
  }
  const bool flat_table = field_relationship == 1;
  if (!total_depth_fits(total_depth, flat_table ? 1U : row_fields->cols)) {
    return Value::error(ErrorCode::Value);
  }

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
    const std::vector<Value> cells =
        aggregate_value_columns(*values, val_cols, group_rows[g], agg, arena, registry, ctx, ErrorCode::Calc);
    for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
      row[key_cols + vc] = cells[vc];
    }
    agg_rows.push_back(std::move(row));
  }

  // -- Sort ---------------------------------------------------------------
  // sort_order numbers the output columns, key columns first; omitted sorts
  // by the keys ascending (see `plan_group_sort`).
  GroupSort sort_plan;
  if (!plan_group_sort(sort_spec, key_cols, val_cols, flat_table, &sort_plan)) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> value_keys;
  std::vector<OuterSortLevel> outer_levels;
  if (sort_plan.value_col != 0U) {
    const std::uint32_t v = sort_plan.value_col - 1U;
    value_keys.reserve(agg_rows.size());
    for (const std::vector<Value>& row : agg_rows) {
      value_keys.push_back(row[key_cols + v]);
    }
    // A hierarchy ranks each prefix of the key columns by its own aggregate.
    if (!flat_table) {
      const std::vector<const Aggregator*> agg_of_group(group_repr.size(), &agg);
      outer_levels =
          build_outer_sort_levels(*row_fields, group_repr, group_rows, *values, v, agg_of_group, arena, registry, ctx);
    }
  }
  std::vector<std::size_t> order;
  order_groups(sort_plan, *row_fields, group_repr, group_is_error, value_keys, outer_levels, &order);

  // -- Compute grand total (if requested) ---------------------------------
  // The grand total aggregates over EVERY filtered data row. It is labelled in
  // the first key column and blank in the remaining key columns.
  // |total_depth| >= 2 adds a subtotal row per key prefix of length
  // 1..|total_depth|-1.
  const bool emit_subtotals = total_depth >= 2 || total_depth <= -2;

  std::vector<Value> grand_total_row;
  bool emit_grand_total = (total_depth != 0);
  if (emit_grand_total) {
    grand_total_row.assign(out_cols, placeholder_cell());
    grand_total_row[0] = Value::text(arena.intern(total_label(ctx, emit_subtotals)));
    const std::vector<Value> totals =
        aggregate_value_columns(*values, val_cols, all_rows, agg, arena, registry, ctx, ErrorCode::Calc);
    for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
      grand_total_row[key_cols + vc] = totals[vc];
    }
  }
  const AxisHierarchy hierarchy = build_axis_hierarchy(*row_fields, group_repr, group_rows, emit_subtotals);
  const std::vector<AxisEntity> slots = plan_axis_slots(hierarchy, order, total_depth);

  // -- Assemble output ----------------------------------------------------
  std::vector<std::vector<Value>> out_rows;
  out_rows.reserve(slots.size() + 2U);

  // Header row (if requested).
  if (layout.output_emits_header) {
    std::vector<Value> header(out_cols, placeholder_cell());
    // field_headers == 3 copies the inputs' header row; 2 generates labels.
    for (std::uint32_t c = 0; c < key_cols; ++c) {
      header[c] = header_label(arena, ctx, field_headers, HeaderField::kRow, row_fields->cells, c);
    }
    for (std::uint32_t c = 0; c < val_cols; ++c) {
      header[key_cols + c] = header_label(arena, ctx, field_headers, HeaderField::kValue, values->cells, c);
    }
    emit_row(&out_rows, std::move(header));
  }

  // Grand total at top (negative total_depth).
  if (emit_grand_total && total_depth < 0) {
    emit_row(&out_rows, std::move(grand_total_row));
  }

  // Group rows and subtotal rows in output order. A subtotal row restates the
  // first `level` key cells of its prefix and leaves the rest blank.
  for (const AxisEntity& slot : slots) {
    if (slot.level == key_cols) {
      emit_row(&out_rows, std::move(agg_rows[slot.index]));
      continue;
    }
    std::vector<Value> row(out_cols, placeholder_cell());
    const std::uint32_t repr_row = axis_entity_repr_row(hierarchy, group_repr, slot);
    for (std::uint32_t c = 0; c < slot.level; ++c) {
      row[c] = row_fields->cells[static_cast<std::size_t>(repr_row) * key_cols + c];
    }
    const std::vector<Value> totals =
        aggregate_value_columns(*values, val_cols, axis_entity_rows(hierarchy, group_rows, all_rows, slot), agg, arena,
                                registry, ctx, ErrorCode::Calc);
    for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
      row[key_cols + vc] = totals[vc];
    }
    emit_row(&out_rows, std::move(row));
  }

  // Grand total at bottom (positive total_depth).
  if (emit_grand_total && total_depth > 0) {
    emit_row(&out_rows, std::move(grand_total_row));
  }

  return rows_to_array_value(out_rows, out_cols, arena);
}

}  // namespace eval
}  // namespace formulon
