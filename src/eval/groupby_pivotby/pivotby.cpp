
#include "eval/groupby_pivotby/pivotby.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
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

namespace {

enum class ColSlotKind { Leaf, OuterSubtotal, GrandTotal };

struct ColSlot {
  ColSlotKind kind = ColSlotKind::Leaf;
  std::size_t col_group = 0;
  std::size_t outer_group = 0;
};

// Maps a single row's group key to its group index, appending a new group on
// first occurrence. `keys` is the source key array (row_fields or
// col_fields); `row` is the absolute row index. Returns the group index.
//
// A composite copy of `representative_rows` is kept so callers can later
// pull the canonical key cells via `keys.cells[representative_rows[g] *
// keys.cols + c]`. Mirrors the inline group-build loop in
// `eval_groupby_lazy`; factored here so both the row and column axes share
// the same first-occurrence semantics.
std::size_t find_or_add_group(const ArrayValue& keys, std::uint32_t row,
                              std::vector<std::uint32_t>* representative_rows,
                              std::vector<std::vector<std::uint32_t>>* member_rows, std::vector<bool>* is_error_group,
                              std::unordered_map<std::string, std::size_t>* index) {
  const std::string key = normalized_group_key(keys, row);
  const auto existing = index->find(key);
  if (existing != index->end()) {
    (*member_rows)[existing->second].push_back(row);
    return existing->second;
  }
  const std::size_t group = representative_rows->size();
  index->emplace(key, group);
  representative_rows->push_back(row);
  member_rows->push_back(std::vector<std::uint32_t>{row});
  is_error_group->push_back(row_key_is_error(keys, row));
  return group;
}

}  // namespace

Value eval_pivotby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 4U || arity > 10U) {
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
  //          slot per outer group when |col_total_depth| == 2)
  //
  // Layout summary, in row order:
  //   1. L col-axis label rows: cells [0..K-1] are blank; each physical
  //      column slot occupies V tiled cells. Leaf slots hold col_fields'
  //      keys, subtotal slots hold the outer key on level 0 and blanks on
  //      inner levels; the optional grand-total block holds "Grand Total"
  //      on the outermost level only and is blank on inner levels.
  //   2. Optional header row: cells [0..K-1] hold the row_fields header
  //      labels (or "Field N" synth); each physical slot holds the values
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
  // oracle goldens. Known divergence from Mac Excel's localized surface:
  // synthesized field/value labels still use English defaults
  // ("Field N" / "Value N") rather than ja-JP labels.

  // -- arg 3: aggregator ----------------------------------------------------
  AggregatorRef agg;
  if (!resolve_aggregator(call.as_call_arg(3), arena, registry, ctx, &agg, &err)) {
    return err;
  }

  // -- arg 4: field_headers ∈ {0,1,2,3}, default 3 -------------------------
  // PIVOTBY's default differs from GROUPBY's (0): pivot output typically
  // wants both the input row to be treated as a header AND a header to be
  // emitted on the output's left/top edges.
  static constexpr int kFieldHeaders[] = {0, 1, 2, 3};
  int field_headers = 3;
  if (!read_optional_int_in_set(call, 4, arity, 3, arena, registry, ctx, kFieldHeaders,
                                sizeof(kFieldHeaders) / sizeof(kFieldHeaders[0]), &field_headers, &err)) {
    return err;
  }

  // -- arg 5: row_total_depth ∈ {-2,-1,0,1,2}, default +1 ------------------
  // The grand-total row (showing column totals) defaults to the BOTTOM of
  // the result. ±2 adds one subtotal row per outer row group; with a single
  // row-key column there is no outer level to roll up, so that shape stays
  // on the ±1 grand-total-only layout.
  static constexpr int kTotalDepths[] = {-2, -1, 0, 1, 2};
  int row_total_depth = 1;
  if (!read_optional_int_in_set(call, 5, arity, 1, arena, registry, ctx, kTotalDepths,
                                sizeof(kTotalDepths) / sizeof(kTotalDepths[0]), &row_total_depth, &err)) {
    return err;
  }

  // -- arg 6: row_sort_order, default 0 ------------------------------------
  // Sort the row groups by their row totals (`SUM`-like aggregation over
  // every (row_group, col_group) cell of the row). 0 preserves first-
  // occurrence order; positive means ascending; negative descending.
  int row_sort_order = 0;
  if (!read_optional_sort_order(call, 6, arity, arena, registry, ctx, &row_sort_order, &err)) {
    return err;
  }

  // -- arg 7: col_total_depth ∈ {-2,-1,0,1,2}, default 1 -------------------
  // The grand-total column (showing row totals) defaults to the RIGHT of
  // the result. ±2 adds one subtotal block per outer column group when the
  // column axis has at least two levels; a one-level column axis retains the
  // ordinary ±1 layout because there is no inner level to roll up.
  int col_total_depth = 1;
  if (!read_optional_int_in_set(call, 7, arity, 1, arena, registry, ctx, kTotalDepths,
                                sizeof(kTotalDepths) / sizeof(kTotalDepths[0]), &col_total_depth, &err)) {
    return err;
  }

  // -- arg 8: col_sort_order, default 0 ------------------------------------
  // Excel pins the zero rejection on the row slot; this slot takes the same
  // signed-column-index domain, so the same rule is applied to both rather
  // than leaving one half of a symmetric argument pair accepting a value the
  // other rejects.
  int col_sort_order = 0;
  if (!read_optional_sort_order(call, 8, arity, arena, registry, ctx, &col_sort_order, &err)) {
    return err;
  }

  // Determine header row layout. Same as GROUPBY but the header / output
  // emission flags drive both the row-axis labels (left edge) and the
  // col-axis labels (top edge).
  auto layout_result = resolve_header_layout(field_headers, row_fields->rows);
  if (!layout_result) {
    return Value::error(layout_result.error());
  }
  const HeaderLayout layout = layout_result.take();
  const std::uint32_t data_start_row = layout.data_start_row;
  const std::uint32_t data_row_count = layout.data_row_count;

  // -- arg 9: filter_array --------------------------------------------------
  std::vector<bool> include_row(data_row_count, true);
  if (arity == 10U) {
    if (!read_filter_mask(call.as_call_arg(9), arena, registry, ctx, data_row_count, &include_row, &err)) {
      return err;
    }
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

  // -- Aggregate per (row_group, col_group, value_col) --------------------
  // For each (rg, cg) pair, build the intersection row list once and reuse
  // it across every value column. Per-cell error isolation: an aggregator
  // failure for one (rg, cg, v) tuple lands in that cell; the rest of the
  // body is still computed.
  std::vector<std::vector<std::vector<Value>>> body(
      n_rows, std::vector<std::vector<Value>>(n_cols, std::vector<Value>(val_cols, Value::blank())));
  for (std::size_t rg = 0; rg < n_rows; ++rg) {
    for (std::size_t cg = 0; cg < n_cols; ++cg) {
      std::vector<std::uint32_t> intersection;
      // Linear walk; the intersection is small for typical workbooks. (For
      // large pivots this hot loop should be revisited with a hash bucket.)
      for (std::uint32_t i = 0; i < data_row_count; ++i) {
        if (!include_row[i]) {
          continue;
        }
        if (row_tag[i] == rg && col_tag[i] == cg) {
          intersection.push_back(data_start_row + i);
        }
      }
      if (intersection.empty()) {
        // No data points at this intersection. Mac Excel surfaces an empty
        // (Blank) cell here — the aggregator is not called for empty
        // groups in the pivot body. Already initialized to Blank.
        continue;
      }
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        const ArrayValue* slice = build_group_slice(*values, v, intersection, arena);
        if (slice == nullptr) {
          body[rg][cg][v] = Value::error(ErrorCode::Num);
          continue;
        }
        body[rg][cg][v] = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
      }
    }
  }

  // -- Compute row totals (one per (row group, value col)) ----------------
  // A "row total" aggregates an entire row group across all col groups,
  // per value column — it lives in the GRAND-TOTAL COLUMN block, whose
  // presence is governed by `col_total_depth`.
  const bool emit_row_totals_col = (col_total_depth != 0);
  std::vector<std::vector<Value>> row_totals(n_rows, std::vector<Value>(val_cols, Value::blank()));
  if (emit_row_totals_col) {
    for (std::size_t rg = 0; rg < n_rows; ++rg) {
      row_totals[rg] =
          aggregate_value_columns(*values, val_cols, row_members[rg], agg, arena, registry, ctx, ErrorCode::Calc);
    }
  }

  // -- Compute column totals (one per (col group, value col)) -------------
  // A "column total" aggregates an entire col group across all row groups,
  // per value column — it lives in the GRAND-TOTAL ROW, whose presence is
  // governed by `row_total_depth`.
  const bool emit_col_totals_row = (row_total_depth != 0);
  std::vector<std::vector<Value>> col_totals(n_cols, std::vector<Value>(val_cols, Value::blank()));
  if (emit_col_totals_row) {
    for (std::size_t cg = 0; cg < n_cols; ++cg) {
      col_totals[cg] =
          aggregate_value_columns(*values, val_cols, col_members[cg], agg, arena, registry, ctx, ErrorCode::Calc);
    }
  }

  // -- Compute grand totals (one per value col) ---------------------------
  // The grand total cells exist only when both axes emit totals.
  std::vector<Value> grand_totals(val_cols, Value::blank());
  if (emit_row_totals_col && emit_col_totals_row) {
    const std::vector<std::uint32_t> all_rows = collect_included_rows(include_row, data_start_row);
    grand_totals = aggregate_value_columns(*values, val_cols, all_rows, agg, arena, registry, ctx, ErrorCode::Calc);
  }

  // -- Sort row groups ----------------------------------------------------
  // Sort by row-total (the "first/only value column" reduces to the row
  // total in single-column-values scope). Error-keyed groups sink to the
  // bottom; sort_order=0 preserves first-occurrence order modulo that
  // sink.
  std::vector<std::size_t> row_order(n_rows);
  for (std::size_t i = 0; i < n_rows; ++i) {
    row_order[i] = i;
  }
  if (row_sort_order != 0) {
    if (row_sort_order != 1 && row_sort_order != -1) {
      // Only ±1 / 0 are supported. With V > 1 the spec for |sort_order|
      // indexing a specific value column is not yet captured by the
      // oracle; reserved for future work.
      return Value::error(ErrorCode::Value);
    }
    const bool descending = (row_sort_order < 0);
    std::stable_sort(row_order.begin(), row_order.end(), [&](std::size_t a, std::size_t b) {
      if (row_is_error[a] != row_is_error[b]) {
        return !row_is_error[a];
      }
      // Sort by the first value column's row total. Compute on the fly if
      // row totals weren't emitted (col_total_depth == 0).
      Value va = row_totals[a][0];
      Value vb = row_totals[b][0];
      if (!emit_row_totals_col) {
        const ArrayValue* sa = build_group_slice(*values, 0U, row_members[a], arena);
        const ArrayValue* sb = build_group_slice(*values, 0U, row_members[b], arena);
        va =
            (sa != nullptr) ? invoke_aggregator_for_group(agg, sa, arena, registry, ctx) : Value::error(ErrorCode::Num);
        vb =
            (sb != nullptr) ? invoke_aggregator_for_group(agg, sb, arena, registry, ctx) : Value::error(ErrorCode::Num);
      }
      const int c = cmp_value_asc(va, vb);
      if (c != 0) {
        return descending ? (c > 0) : (c < 0);
      }
      return cmp_keys_asc(*row_fields, row_repr[a], row_repr[b]) < 0;
    });
  } else {
    std::stable_sort(row_order.begin(), row_order.end(), [&](std::size_t a, std::size_t b) {
      if (row_is_error[a] != row_is_error[b]) {
        return !row_is_error[a];
      }
      return false;
    });
  }

  // -- Sort col groups ----------------------------------------------------
  std::vector<std::size_t> col_order(n_cols);
  for (std::size_t i = 0; i < n_cols; ++i) {
    col_order[i] = i;
  }
  if (col_sort_order != 0) {
    if (col_sort_order != 1 && col_sort_order != -1) {
      // Only ±1 / 0 are supported; see row_sort_order note above.
      return Value::error(ErrorCode::Value);
    }
    const bool descending = (col_sort_order < 0);
    std::stable_sort(col_order.begin(), col_order.end(), [&](std::size_t a, std::size_t b) {
      if (col_is_error[a] != col_is_error[b]) {
        return !col_is_error[a];
      }
      Value va = col_totals[a][0];
      Value vb = col_totals[b][0];
      if (!emit_col_totals_row) {
        const ArrayValue* sa = build_group_slice(*values, 0U, col_members[a], arena);
        const ArrayValue* sb = build_group_slice(*values, 0U, col_members[b], arena);
        va =
            (sa != nullptr) ? invoke_aggregator_for_group(agg, sa, arena, registry, ctx) : Value::error(ErrorCode::Num);
        vb =
            (sb != nullptr) ? invoke_aggregator_for_group(agg, sb, arena, registry, ctx) : Value::error(ErrorCode::Num);
      }
      const int c = cmp_value_asc(va, vb);
      if (c != 0) {
        return descending ? (c > 0) : (c < 0);
      }
      return cmp_keys_asc(*col_fields, col_repr[a], col_repr[b]) < 0;
    });
  } else {
    std::stable_sort(col_order.begin(), col_order.end(), [&](std::size_t a, std::size_t b) {
      if (col_is_error[a] != col_is_error[b]) {
        return !col_is_error[a];
      }
      return false;
    });
  }

  // A two-level column axis can expose one value-wide subtotal block for
  // every outer key. Build the physical block plan once and let all of the
  // renderers below consume it; keeping the plan explicit avoids the fixed
  // `ci * val_cols` assumption that the ordinary leaf-only layout uses.
  const bool emit_col_subtotals = (col_total_depth == 2 || col_total_depth == -2) && col_levels >= 2U;
  OuterGrouping col_hierarchy;
  std::vector<std::size_t> outer_col_order;
  std::vector<ColSlot> col_slots;
  if (emit_col_subtotals) {
    col_hierarchy = build_outer_grouping(*col_fields, col_repr, col_members);
    std::vector<bool> outer_seen(col_hierarchy.repr_of_outer.size(), false);
    for (const std::size_t cg : col_order) {
      const std::size_t outer = col_hierarchy.outer_of_group[cg];
      if (!outer_seen[outer]) {
        outer_seen[outer] = true;
        outer_col_order.push_back(outer);
      }
    }
    for (const std::size_t outer : outer_col_order) {
      if (col_total_depth < 0) {
        col_slots.push_back({ColSlotKind::OuterSubtotal, 0U, outer});
      }
      for (const std::size_t cg : col_order) {
        if (col_hierarchy.outer_of_group[cg] == outer) {
          col_slots.push_back({ColSlotKind::Leaf, cg, outer});
        }
      }
      if (col_total_depth > 0) {
        col_slots.push_back({ColSlotKind::OuterSubtotal, 0U, outer});
      }
    }
  } else {
    col_slots.reserve(n_cols);
    for (const std::size_t cg : col_order) {
      col_slots.push_back({ColSlotKind::Leaf, cg, 0U});
    }
  }

  // Subtotal cells are the intersection of one flat row group and one outer
  // column group. Empty intersections stay blank, just like ordinary pivot
  // body cells; an aggregator error is retained in that one cell.
  std::vector<std::vector<std::vector<Value>>> col_subtotal_body;
  std::vector<std::vector<Value>> col_subtotal_totals;
  if (emit_col_subtotals) {
    const std::size_t outer_count = col_hierarchy.repr_of_outer.size();
    col_subtotal_body.assign(
        n_rows, std::vector<std::vector<Value>>(outer_count, std::vector<Value>(val_cols, Value::blank())));
    for (std::size_t rg = 0; rg < n_rows; ++rg) {
      for (std::size_t outer = 0; outer < outer_count; ++outer) {
        std::vector<std::uint32_t> intersection;
        for (const std::uint32_t row : row_members[rg]) {
          const std::size_t offset = static_cast<std::size_t>(row - data_start_row);
          if (offset < col_tag.size() && col_hierarchy.outer_of_group[col_tag[offset]] == outer) {
            intersection.push_back(row);
          }
        }
        if (!intersection.empty()) {
          col_subtotal_body[rg][outer] =
              aggregate_value_columns(*values, val_cols, intersection, agg, arena, registry, ctx, ErrorCode::Calc);
        }
      }
    }
    if (emit_col_totals_row) {
      col_subtotal_totals.assign(outer_count, std::vector<Value>(val_cols, Value::blank()));
      for (std::size_t outer = 0; outer < outer_count; ++outer) {
        col_subtotal_totals[outer] = aggregate_value_columns(*values, val_cols, col_hierarchy.rows_of_outer[outer], agg,
                                                             arena, registry, ctx, ErrorCode::Calc);
      }
    }
  }

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
  // field_headers ∈ {1,3} we copy values->cells[v] for v=0..V-1 (the
  // input's row 0). With field_headers == 2 we synth "Value <v+1>".
  // field_headers == 0 leaves these blank (output_emits_header is false).
  std::vector<Value> values_header_labels(val_cols, Value::blank());
  if (layout.output_emits_header) {
    if (field_headers == 1 || field_headers == 3) {
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        values_header_labels[v] = values->cells[v];
      }
    } else {
      // field_headers == 2: synth English defaults. ja-JP "値 N" is a
      // documented divergence (not emitted).
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        const std::string label = "Value " + std::to_string(v + 1U);
        values_header_labels[v] = Value::text(arena.intern(label));
      }
    }
  }

  // Helper: render one of the L col-axis label rows. `level` is the
  // 0-based col_fields column index for this row.
  auto render_col_axis_row = [&](std::uint32_t level) {
    std::vector<Value> row(out_cols, Value::blank());
    // Cells [0..K-1] stay blank.
    for (std::size_t ci = 0; ci < col_slots.size(); ++ci) {
      const ColSlot& slot = col_slots[ci];
      Value label = Value::blank();
      if (slot.kind == ColSlotKind::Leaf) {
        label = col_fields->cells[static_cast<std::size_t>(col_repr[slot.col_group]) * col_levels + level];
      } else if (slot.kind == ColSlotKind::OuterSubtotal) {
        // A subtotal is identified by the outer key on level 0; its inner
        // header cells remain blank. This is the shape observed in Excel's
        // PIVOTBY column subtotal output.
        if (level == 0U) {
          label =
              col_fields->cells[static_cast<std::size_t>(col_hierarchy.repr_of_outer[slot.outer_group]) * col_levels];
        }
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
        const Value gt =
            Value::text(arena.intern(emit_col_subtotals ? hierarchy_grand_total_label(ctx) : grand_total_label(ctx)));
        for (std::uint32_t v = 0; v < val_cols; ++v) {
          row[grand_total_block_start + v] = gt;
        }
      }
    }
    return row;
  };

  // Helper: render the (single) header row when output_emits_header.
  auto render_header_row = [&]() {
    std::vector<Value> row(out_cols, Value::blank());
    // Row-fields header labels (cells [0..K-1]).
    if (field_headers == 1 || field_headers == 3) {
      for (std::uint32_t c = 0; c < key_cols; ++c) {
        row[c] = row_fields->cells[c];
      }
    } else {
      // field_headers == 2: synth "Field N". ja-JP "行フィールド N" is a
      // documented divergence.
      for (std::uint32_t c = 0; c < key_cols; ++c) {
        const std::string label = "Field " + std::to_string(c + 1U);
        row[c] = Value::text(arena.intern(label));
      }
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
    std::vector<Value> row(out_cols, Value::blank());
    for (std::uint32_t level = 0; level < col_levels; ++level) {
      const std::uint32_t dst = body_block_start + level * val_cols;
      if (dst < out_cols) {
        row[dst] = col_fields->cells[level];
      }
    }
    return row;
  };

  // Helper: render one body row (per row-group rg).
  auto render_body_row = [&](std::size_t rg) {
    std::vector<Value> row(out_cols, Value::blank());
    // Row keys: copy K cells from the representative row.
    for (std::uint32_t c = 0; c < key_cols; ++c) {
      row[c] = row_fields->cells[static_cast<std::size_t>(row_repr[rg]) * key_cols + c];
    }
    // Body cells: per physical column slot, value column.
    for (std::size_t ci = 0; ci < col_slots.size(); ++ci) {
      const ColSlot& slot = col_slots[ci];
      const std::uint32_t base = body_block_start + static_cast<std::uint32_t>(ci) * val_cols;
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        if (slot.kind == ColSlotKind::Leaf) {
          row[base + v] = body[rg][slot.col_group][v];
        } else if (slot.kind == ColSlotKind::OuterSubtotal) {
          row[base + v] = col_subtotal_body[rg][slot.outer_group][v];
        }
      }
    }
    // Grand-total block: row totals per value column. Mac Excel ja-JP only
    // populates this strip for the single-value-column layout; with
    // val_cols > 1 the grand-total columns carry headers but blank values.
    if (emit_row_totals_col && val_cols == 1U) {
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        row[grand_total_block_start + v] = row_totals[rg][v];
      }
    }
    return row;
  };

  // A ±2 row request adds one subtotal row per outer row group — the first
  // row-key column alone. With a single row-key column the outer level
  // coincides with the row groups themselves, so that shape stays on the ±1
  // layout.
  const bool emit_row_subtotals = (row_total_depth == 2 || row_total_depth == -2) && key_cols >= 2U;

  // Helper: render a TOP/BOTTOM grand-total row (column totals).
  auto render_totals_row = [&]() {
    std::vector<Value> row(out_cols, Value::blank());
    row[0] = Value::text(arena.intern(emit_row_subtotals ? hierarchy_grand_total_label(ctx) : grand_total_label(ctx)));
    // Cells [1..K-1] stay blank (the rest of the row-keys columns).
    for (std::size_t ci = 0; ci < col_slots.size(); ++ci) {
      const ColSlot& slot = col_slots[ci];
      const std::uint32_t base = body_block_start + static_cast<std::uint32_t>(ci) * val_cols;
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        if (slot.kind == ColSlotKind::Leaf) {
          row[base + v] = col_totals[slot.col_group][v];
        } else if (slot.kind == ColSlotKind::OuterSubtotal) {
          row[base + v] = col_subtotal_totals[slot.outer_group][v];
        }
      }
    }
    // Grand-total block: the bottom-right grand total, populated only for
    // the single-value-column layout (see render_body_row note).
    if (emit_row_totals_col && val_cols == 1U) {
      for (std::uint32_t v = 0; v < val_cols; ++v) {
        row[grand_total_block_start + v] = grand_totals[v];
      }
    }
    return row;
  };

  // The all-singletons case (K==1 AND L==1 AND V==1) uses a compact
  // merged layout that pre-dates the multi-column extension and is
  // preserved for backwards compatibility. Mac Excel always shows the
  // pivot's column keys (X/Y) on this row regardless of field_headers --
  // they identify which output column is which, independent of whether
  // descriptive field-name headers are requested -- so this row is
  // unconditionally emitted, mirroring the multi-column layout below:
  //   - When output_emits_header is true: the row_fields header label
  //     occupies col 0 alongside the col-axis labels at cols 1..; the
  //     values-header tile collapses into the col-axis row.
  //   - When output_emits_header is false: col 0 (the corner cell) stays
  //     blank, but the col-axis labels and the grand-total label (if
  //     col_total_depth != 0) still render.
  // For K > 1 OR L > 1 OR V > 1, the multi-column layout always emits L
  // col-axis label rows and (optionally) a separate header row beneath
  // them.
  const bool merged_single_col_layout = (key_cols == 1U && col_levels == 1U && val_cols == 1U);

  // Helper: render the merged top row for the single-column layout. The
  // row carries:
  //   - col 0: row_fields header label (or "Field 1" for fh=2), left blank
  //     when output_emits_header is false
  //   - cols 1..nC: col-axis labels (level 0 keys), one per col group --
  //     always rendered, independent of output_emits_header
  //   - last col (if emit_row_totals_col): "Grand Total"
  // V is always 1 in this branch (key_cols == 1 == col_levels), so no
  // tiling is required.
  auto render_merged_header_row = [&]() {
    std::vector<Value> row(out_cols, Value::blank());
    if (layout.output_emits_header) {
      if (field_headers == 1 || field_headers == 3) {
        row[0] = row_fields->cells[0];
      } else {
        // field_headers == 2: synth "Field 1".
        row[0] = Value::text(arena.intern("Field 1"));
      }
    }
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

  // -- Nested row subtotals (|row_total_depth| == 2) -----------------------
  // A ±2 request adds one subtotal row per outer row group — the first row-
  // key column alone. Each subtotal row carries the outer group's aggregate
  // for every (col group, value col) cell, plus its total in the grand-total
  // column.
  OuterGrouping row_hierarchy;
  std::vector<std::size_t> outer_row_order;
  std::vector<std::vector<std::size_t>> row_groups_of_outer;
  std::vector<std::vector<Value>> row_subtotal_rows;
  if (emit_row_subtotals) {
    row_hierarchy = build_outer_grouping(*row_fields, row_repr, row_members);
    const std::size_t outer_count = row_hierarchy.repr_of_outer.size();
    // The hierarchy is primary and the row sort applies within it.
    row_groups_of_outer.resize(outer_count);
    std::vector<bool> outer_seen(outer_count, false);
    for (std::size_t ri = 0; ri < n_rows; ++ri) {
      const std::size_t rg = row_order[ri];
      const std::size_t outer = row_hierarchy.outer_of_group[rg];
      if (!outer_seen[outer]) {
        outer_seen[outer] = true;
        outer_row_order.push_back(outer);
      }
      row_groups_of_outer[outer].push_back(rg);
    }
    row_subtotal_rows.reserve(outer_count);
    for (std::size_t o = 0; o < outer_count; ++o) {
      // Split the outer group's rows by column group once, then aggregate
      // each bucket per value column.
      std::vector<std::vector<std::uint32_t>> by_col_group(n_cols);
      for (std::uint32_t row : row_hierarchy.rows_of_outer[o]) {
        by_col_group[col_tag[row - data_start_row]].push_back(row);
      }
      std::vector<Value> row(out_cols, Value::blank());
      // The subtotal row restates its outer key verbatim in the first row-key
      // column and leaves the inner row-key columns blank.
      row[0] = row_fields->cells[static_cast<std::size_t>(row_hierarchy.repr_of_outer[o]) * key_cols];
      for (std::size_t ci = 0; ci < col_slots.size(); ++ci) {
        const ColSlot& slot = col_slots[ci];
        std::vector<std::uint32_t> intersection;
        if (slot.kind == ColSlotKind::Leaf) {
          intersection = by_col_group[slot.col_group];
        } else if (slot.kind == ColSlotKind::OuterSubtotal) {
          // Row and column subtotals intersect at their two outer keys.
          for (const std::uint32_t source_row : row_hierarchy.rows_of_outer[o]) {
            const std::size_t offset = static_cast<std::size_t>(source_row - data_start_row);
            if (offset < col_tag.size() && col_hierarchy.outer_of_group[col_tag[offset]] == slot.outer_group) {
              intersection.push_back(source_row);
            }
          }
        }
        if (intersection.empty()) {
          continue;  // no data at this intersection; Excel leaves it blank.
        }
        const std::vector<Value> cells =
            aggregate_value_columns(*values, val_cols, intersection, agg, arena, registry, ctx, ErrorCode::Calc);
        const std::uint32_t base = body_block_start + static_cast<std::uint32_t>(ci) * val_cols;
        for (std::uint32_t v = 0; v < val_cols; ++v) {
          row[base + v] = cells[v];
        }
      }
      // The grand-total strip follows the body rows' rule: populated only
      // for the single-value-column layout.
      if (emit_row_totals_col && val_cols == 1U) {
        const std::vector<Value> totals = aggregate_value_columns(*values, val_cols, row_hierarchy.rows_of_outer[o],
                                                                  agg, arena, registry, ctx, ErrorCode::Calc);
        for (std::uint32_t v = 0; v < val_cols; ++v) {
          row[grand_total_block_start + v] = totals[v];
        }
      }
      row_subtotal_rows.push_back(std::move(row));
    }
  }

  std::vector<std::vector<Value>> out_rows;
  out_rows.reserve(static_cast<std::size_t>(col_levels) + n_rows + row_subtotal_rows.size() + 3U);
  if (merged_single_col_layout) {
    // Merged single-column layout: the combined top row (col-axis labels,
    // plus the row_fields header label when output_emits_header) is
    // always emitted -- see the comment on merged_single_col_layout above.
    out_rows.push_back(render_merged_header_row());
  } else {
    // Multi-column layout: always emit L col-axis label rows; emit a
    // separate header row when output_emits_header is true.
    if (layout.output_emits_header && (field_headers == 1 || field_headers == 3)) {
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
    out_rows.push_back(render_totals_row());
  }
  // Body rows. With subtotals each outer row group's rows are bracketed by
  // its subtotal row, placed on the same side of the block as the grand
  // total row.
  if (emit_row_subtotals) {
    for (std::size_t o : outer_row_order) {
      if (row_total_depth < 0) {
        out_rows.push_back(row_subtotal_rows[o]);
      }
      for (std::size_t rg : row_groups_of_outer[o]) {
        out_rows.push_back(render_body_row(rg));
      }
      if (row_total_depth > 0) {
        out_rows.push_back(row_subtotal_rows[o]);
      }
    }
  } else {
    for (std::size_t ri = 0; ri < n_rows; ++ri) {
      out_rows.push_back(render_body_row(row_order[ri]));
    }
  }
  // Optional BOTTOM grand-total row.
  if (emit_col_totals_row && row_total_depth > 0) {
    out_rows.push_back(render_totals_row());
  }

  return rows_to_array_value(out_rows, out_cols, arena);
}

}  // namespace eval
}  // namespace formulon
