// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "eval/groupby_pivotby/groupby.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/groupby_pivotby/common.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/structured_log.h"
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
  AggregatorRef agg;
  if (!resolve_aggregator(call.as_call_arg(2), arena, registry, ctx, &agg, &err)) {
    return err;
  }

  // -- arg 3: field_headers ∈ {0,1,2,3} ------------------------------------
  static constexpr int kFieldHeaders[] = {0, 1, 2, 3};
  int field_headers = 0;
  if (!read_optional_int_in_set(call, 3, arity, 0, arena, registry, ctx, kFieldHeaders,
                                sizeof(kFieldHeaders) / sizeof(kFieldHeaders[0]), &field_headers, &err)) {
    return err;
  }

  // -- arg 4: total_depth ∈ {-2,-1,0,1,2} ----------------------------------
  static constexpr int kTotalDepths[] = {-2, -1, 0, 1, 2};
  int total_depth = -1;
  if (!read_optional_int_in_set(call, 4, arity, -1, arena, registry, ctx, kTotalDepths,
                                sizeof(kTotalDepths) / sizeof(kTotalDepths[0]), &total_depth, &err)) {
    return err;
  }

  // -- arg 5: sort_order ----------------------------------------------------
  int sort_order = 0;
  if (!read_optional_int(call, 5, arity, 0, arena, registry, ctx, &sort_order, &err)) {
    return err;
  }

  // Determine header row layout. Inputs have a header row when
  // field_headers ∈ {1, 3}; outputs emit a header row when
  // field_headers ∈ {1, 2, 3}.
  auto layout_result = resolve_header_layout(field_headers, row_fields->rows);
  if (!layout_result) {
    return Value::error(layout_result.error());
  }
  const HeaderLayout layout = layout_result.take();
  const std::uint32_t data_start_row = layout.data_start_row;
  const std::uint32_t data_row_count = layout.data_row_count;

  // -- arg 6: filter_array --------------------------------------------------
  std::vector<bool> include_row(data_row_count, true);
  if (arity == 7U) {
    if (!read_filter_mask(call.as_call_arg(6), arena, registry, ctx, data_row_count, &include_row, &err)) {
      return err;
    }
  }

  // -- Build groups --------------------------------------------------------
  // Walk filtered data rows in input order; for each row, look up its group
  // key against the existing list of unique keys (linear scan via
  // `group_key_equal`). New keys append; matching keys add the row index to
  // their bucket. This preserves first-occurrence ordering for sort_order=0
  // for free.
  //
  // Group representatives are stored as row indices (into `row_fields`) so
  // that subsequent equality checks can re-use the same column-walk path.
  std::vector<std::uint32_t> group_repr;               // row index of representative
  std::vector<std::vector<std::uint32_t>> group_rows;  // row indices in each group
  std::vector<bool> group_is_error;

  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    if (!include_row[i]) {
      continue;
    }
    const std::uint32_t row = data_start_row + i;
    bool matched = false;
    for (std::size_t g = 0; g < group_repr.size(); ++g) {
      if (group_key_equal(*row_fields, row, group_repr[g])) {
        group_rows[g].push_back(row);
        matched = true;
        break;
      }
    }
    if (!matched) {
      group_repr.push_back(row);
      group_rows.push_back(std::vector<std::uint32_t>{row});
      group_is_error.push_back(row_key_is_error(*row_fields, row));
    }
  }

  if (group_repr.empty()) {
    return Value::error(ErrorCode::Value);
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
    std::vector<Value> row(out_cols, Value::blank());
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
      row[key_cols + vc] = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
    }
    agg_rows.push_back(std::move(row));
  }

  // -- Sort ---------------------------------------------------------------
  // sort_order semantics:
  //   * 0 -> preserve first-occurrence order (already true).
  //   * N>0 -> stable-sort ascending by the N-th aggregated value column
  //     (1-based; out-of-range -> #VALUE!). Tie-break on group key.
  //   * N<0 -> stable-sort descending by |N|-th column.
  // Error-keyed groups always sort to the bottom (after all valid groups).
  std::vector<std::size_t> order(agg_rows.size());
  for (std::size_t i = 0; i < order.size(); ++i) {
    order[i] = i;
  }
  if (sort_order != 0) {
    const int abs_sort = sort_order > 0 ? sort_order : -sort_order;
    if (abs_sort < 1 || static_cast<std::uint32_t>(abs_sort) > val_cols) {
      return Value::error(ErrorCode::Value);
    }
    const std::uint32_t value_col_idx = static_cast<std::uint32_t>(abs_sort - 1);
    const bool descending = (sort_order < 0);
    std::stable_sort(order.begin(), order.end(), [&](std::size_t a, std::size_t b) {
      // Error-keyed groups always go last regardless of
      // direction; this matches Mac Excel's "errors trail"
      // surface for the dynamic-array sort family.
      if (group_is_error[a] != group_is_error[b]) {
        return !group_is_error[a];
      }
      const Value& va = agg_rows[a][key_cols + value_col_idx];
      const Value& vb = agg_rows[b][key_cols + value_col_idx];
      const int c = cmp_value_asc(va, vb);
      if (c != 0) {
        return descending ? (c > 0) : (c < 0);
      }
      // Tie-break on group key (always ascending).
      return cmp_keys_asc(*row_fields, group_repr[a], group_repr[b]) < 0;
    });
  } else {
    // Even at sort_order=0, error-keyed groups sink to the bottom in stable
    // first-occurrence order (matching Mac Excel's UNIQUE / FILTER pattern
    // for cells whose evaluation produced an error).
    std::stable_sort(order.begin(), order.end(), [&](std::size_t a, std::size_t b) {
      if (group_is_error[a] != group_is_error[b]) {
        return !group_is_error[a];
      }
      return false;  // preserve original order within bucket
    });
  }

  // -- Compute grand total (if requested) ---------------------------------
  // The grand total aggregates over EVERY filtered data row. It is rendered
  // with the literal "Grand Total" label in the first key column and blank
  // cells in the remaining key columns.
  std::vector<Value> grand_total_row;
  bool emit_grand_total = (total_depth != 0);
  if (emit_grand_total) {
    grand_total_row.assign(out_cols, Value::blank());
    grand_total_row[0] = Value::text(arena.intern(grand_total_label(ctx)));
    const std::vector<std::uint32_t> all_rows = collect_included_rows(include_row, data_start_row);
    const std::vector<Value> totals =
        aggregate_value_columns(*values, val_cols, all_rows, agg, arena, registry, ctx, ErrorCode::Calc);
    for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
      grand_total_row[key_cols + vc] = totals[vc];
    }
  }
  // Subtotals (|total_depth| == 2) request per-outer-group subtotal rows
  // in addition to the grand total. Implementing them requires replacing
  // the flat composite-key grouping with an outer/inner hierarchy plus the
  // corresponding subtotal-row placement, which is not yet built. Until
  // then a ±2 request falls back to the ±1 (grand-total-only) layout. The
  // fallback is recorded in tests/divergence.yaml; emit a diagnostic so it
  // is observable rather than silent.
  if (total_depth == 2 || total_depth == -2) {
    StructuredLog("eval.groupby.subtotals_unsupported")
        .field("function", std::string_view("GROUPBY"))
        .field("total_depth", static_cast<int64_t>(total_depth))
        .field("fallback", std::string_view("grand_total_only"))
        .warn();
  }

  // -- Assemble output ----------------------------------------------------
  std::vector<std::vector<Value>> out_rows;
  out_rows.reserve(agg_rows.size() + 2U);

  // Header row (if requested).
  if (layout.output_emits_header) {
    std::vector<Value> header(out_cols, Value::blank());
    if (field_headers == 1 || field_headers == 3) {
      // Inputs had a header row (row 0 of each input). Copy it verbatim.
      for (std::uint32_t c = 0; c < key_cols; ++c) {
        header[c] = row_fields->cells[c];
      }
      for (std::uint32_t c = 0; c < val_cols; ++c) {
        header[key_cols + c] = values->cells[c];
      }
    } else {
      // field_headers == 2: synthesize English defaults. Mac Excel may
      // emit Japanese labels in ja-JP; this divergence is logged for the
      // first oracle run.
      for (std::uint32_t c = 0; c < key_cols; ++c) {
        const std::string label = "Field " + std::to_string(c + 1U);
        header[c] = Value::text(arena.intern(label));
      }
      for (std::uint32_t c = 0; c < val_cols; ++c) {
        const std::string label = "Value " + std::to_string(c + 1U);
        header[key_cols + c] = Value::text(arena.intern(label));
      }
    }
    emit_row(&out_rows, header);
  }

  // Grand total at top (negative total_depth).
  if (emit_grand_total && total_depth < 0) {
    emit_row(&out_rows, grand_total_row);
  }

  // Per-group rows in sorted order.
  for (std::size_t i : order) {
    emit_row(&out_rows, agg_rows[i]);
  }

  // Grand total at bottom (positive total_depth).
  if (emit_grand_total && total_depth > 0) {
    emit_row(&out_rows, grand_total_row);
  }

  return rows_to_array_value(out_rows, out_cols, arena);
}

}  // namespace eval
}  // namespace formulon
