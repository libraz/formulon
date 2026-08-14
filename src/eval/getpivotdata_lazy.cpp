//
// Implementation of the GETPIVOTDATA lazy form. See
// eval/getpivotdata_lazy.h for the public contract.
//
// Algorithm:
//   1. Validate arity (>= 2 and (count - 2) is even).
//   2. Evaluate arg 0 (data field name) eagerly and coerce to text.
//      Errors propagate; a non-coercible value (e.g. an array) surfaces
//      `#REF!`.
//   3. Inspect arg 1's AST: it must be a `Ref` (single cell) or a
//      `RangeOp` whose left-hand endpoint is a cell `Ref`. The
//      reference's `(sheet, row, col)` becomes the pivot anchor; an
//      empty `sheet` field falls back to `EvalContext::current_sheet()`.
//      Anything else surfaces `#REF!`.
//   4. Resolve the anchor via `pivot::find_pivot_at_anchor`. Missing
//      sheet, missing pivot, or no workbook bound -> `#REF!`.
//   5. If the pivot's `last_result_` cache is empty, refresh it by
//      calling `pivot::evaluate(*pivot, *cache)`. The `last_result_`
//      slot is `mutable` on `PivotTable` so this works through a const
//      pointer (the lazy form only sees the workbook through
//      `EvalContext::workbook()`, which returns `const Workbook*`).
//   6. Walk the (field, item) pairs:
//        * Locate each `field_text` in the pivot table's `fields()` by
//          `custom_name` (or `source_name`) -- case-sensitive, matching
//          Mac Excel's GETPIVOTDATA exact-match contract.
//        * Determine whether the field is on the row axis or the col
//          axis; record the pair as a (depth, item_text) entry on the
//          appropriate axis.
//        * Walk the row hierarchy depth-first to find the leaf whose
//          ancestors match every recorded row-axis pair (in the order
//          declared by `row_field_order`). Same for the col hierarchy.
//        * Partial paths -- caller did not supply every row/col field --
//          surface `#REF!`. Only the (no-args, grand_total) and
//          (full-axis, leaf) shapes are honoured in the MVP.
//   7. Locate the data field by name in `data_fields()` -> data field
//      index `df_idx`. Emit `result.values[row_leaf][col_leaf][df_idx]`
//      reified into the eval arena (see `reify_text` below). If no
//      row/col fields are configured, the (0, 0) cell is the sole leaf;
//      this also covers the "plain `=GETPIVOTDATA("Sum of Amount", A3)`
//      grand-total formula when row/col fields are configured" case via
//      the same path because the grand total of a single-aggregation
//      pivot equals `result.values[0][0][df_idx]` only when there are
//      no axis fields. With axis fields, no axis pairs -> `#REF!`. The
//      design corpus (§15.1.4) keeps this strict because Mac itself
//      surfaces `#REF!` for partial-path lookups.

#include "eval/getpivotdata_lazy.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/pivot_locale.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Mac's GETPIVOTDATA returns `#REF!` for nearly every error path: bad
// anchor, unknown data field, unknown item label, partial address. The
// only exception is an error argument (which propagates) and a
// non-coercible argument shape (which is reported as `#REF!` by the
// design's MVP rules).
constexpr ErrorCode kPivotRefError = ErrorCode::Ref;

// Drills down to the leftmost cell-`Ref` of a `RangeOp` chain, e.g.
// `A1:B2` -> `A1`. Returns nullptr when the leftmost descendant is not
// a `Ref`. RangeOp lhs is itself permitted to be another RangeOp in
// pathological parser output, so we loop rather than relying on a
// single dereference.
const parser::Reference* anchor_ref_of(const parser::AstNode& node) noexcept {
  const parser::AstNode* cur = &node;
  while (cur->kind() == parser::NodeKind::RangeOp) {
    cur = &cur->as_range_lhs();
  }
  if (cur->kind() != parser::NodeKind::Ref) {
    return nullptr;
  }
  return &cur->as_ref();
}

// Locates `name` in `fields` by `custom_name` (preferred when non-empty,
// matching the OOXML reader's behaviour) or `source_name`. Returns the
// 0-based index, or `static_cast<std::size_t>(-1)` when no field
// matches. Comparison is case-sensitive: Mac's GETPIVOTDATA requires an
// exact text match, including whitespace.
std::size_t find_field_index(const std::vector<pivot::PivotField>& fields, std::string_view name) noexcept {
  for (std::size_t i = 0; i < fields.size(); ++i) {
    const pivot::PivotField& f = fields[i];
    if (!f.custom_name.empty() && f.custom_name == name) {
      return i;
    }
    if (f.source_name == name) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}

// Locates `name` in `data_fields` by display name. Returns the 0-based
// index, or `static_cast<std::size_t>(-1)` when no data field matches.
std::size_t find_data_field_index(const std::vector<pivot::PivotDataField>& data_fields,
                                  std::string_view name) noexcept {
  for (std::size_t i = 0; i < data_fields.size(); ++i) {
    if (data_fields[i].name == name) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}

// Returns the depth (0-based) of `field_idx` inside `axis_order`, or
// `static_cast<std::size_t>(-1)` when the field is not on this axis.
std::size_t depth_in_axis(const std::vector<std::uint32_t>& axis_order, std::size_t field_idx) noexcept {
  for (std::size_t i = 0; i < axis_order.size(); ++i) {
    if (axis_order[i] == field_idx) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}

// Walks a hierarchy tree of `Node` (either `RowHierarchyNode` or
// `ColHierarchyNode`) following `path`. Each `path[i]` is the label
// expected at depth `i`. Returns the leaf index assigned to the
// matching leaf, or `static_cast<std::size_t>(-1)` on any mismatch.
//
// The leaf index is computed by counting leaves in document order,
// matching the dense indexing used by `pivot::evaluate` in
// `finalize_hierarchy<Node>`.
template <class Node>
std::size_t walk_hierarchy(const std::vector<Node>& roots, const std::vector<std::string>& path) noexcept {
  if (path.empty() || roots.empty()) {
    return static_cast<std::size_t>(-1);
  }
  std::size_t leaf_index = 0;
  const std::vector<Node>* level = &roots;
  for (std::size_t depth = 0; depth < path.size(); ++depth) {
    const std::string& want = path[depth];
    bool matched = false;
    for (const Node& node : *level) {
      if (node.label == want) {
        // Descend.
        if (depth + 1 == path.size()) {
          // Leaf level: this node must be a leaf for the path to be
          // valid. A non-leaf at the final position means the caller
          // gave a partial address at this axis, which the MVP refuses.
          if (!node.children.empty()) {
            return static_cast<std::size_t>(-1);
          }
          return leaf_index;
        }
        if (node.children.empty()) {
          // Path expects to descend further, but we're already at a
          // leaf -- the caller asked for a deeper item than the pivot
          // exposes.
          return static_cast<std::size_t>(-1);
        }
        level = &node.children;
        matched = true;
        break;
      }
      // Non-matching subtree: skip its leaf count so the leaf_index
      // counter stays aligned with `pivot::evaluate`'s dense ordering.
      if (node.children.empty()) {
        ++leaf_index;
      } else {
        // Count all leaves in this subtree.
        std::vector<const Node*> stack;
        stack.push_back(&node);
        while (!stack.empty()) {
          const Node* top = stack.back();
          stack.pop_back();
          if (top->children.empty()) {
            ++leaf_index;
          } else {
            for (const Node& c : top->children) {
              stack.push_back(&c);
            }
          }
        }
      }
    }
    if (!matched) {
      return static_cast<std::size_t>(-1);
    }
  }
  return static_cast<std::size_t>(-1);
}

// Re-emits `v` so any Text payload is rooted in `arena` (the evaluator's
// per-call arena), independent of the `PivotResult::text_storage` deque.
// Numbers / bools / errors / blanks are trivially copyable.
Value reify_in_arena(const Value& v, Arena& arena) {
  if (!v.is_text()) {
    return v;
  }
  return Value::text(arena.intern(v.as_text()));
}

}  // namespace

Value eval_getpivotdata_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Required: data_field + pivot_anchor. Optional: any number of
  // (field, item) pairs; total count after the first 2 must be even.
  if (arity < 2U || ((arity - 2U) % 2U) != 0U) {
    return Value::error(kPivotRefError);
  }

  // 1. Evaluate data field name (eager). Errors propagate.
  const Value data_field_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (data_field_v.is_error()) {
    return data_field_v;
  }
  Expected<std::string, ErrorCode> data_field_text = coerce_to_text(data_field_v);
  if (!data_field_text) {
    // coerce_to_text yields #VALUE! for arrays / refs / lambdas; remap
    // to the GETPIVOTDATA-specific #REF! surface.
    return Value::error(kPivotRefError);
  }

  // 2. Inspect anchor AST: must be a Ref or a RangeOp's leftmost Ref.
  const parser::AstNode& anchor_node = call.as_call_arg(1);
  // First evaluate it speculatively for error propagation: any error
  // hidden inside (e.g. a malformed cross-sheet ref the parser turned
  // into a placeholder, or an arithmetic expression) must surface
  // before the structural check rejects the shape.
  const parser::Reference* anchor = anchor_ref_of(anchor_node);
  if (anchor == nullptr) {
    // Eager-evaluate so an embedded error (e.g. =GETPIVOTDATA("F", 1/0))
    // propagates instead of being masked by the structural #REF!.
    const Value v = eval_node(anchor_node, arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    return Value::error(kPivotRefError);
  }

  // Anchor sheet identity: explicit sheet qualifier wins; otherwise
  // fall back to the bound current sheet. An unbound context can never
  // address a pivot, so surface #REF!.
  std::string_view sheet_name = anchor->sheet;
  if (sheet_name.empty()) {
    if (ctx.current_sheet() == nullptr) {
      return Value::error(kPivotRefError);
    }
    sheet_name = ctx.current_sheet()->name();
  }

  // 3. Locate the pivot table covering (sheet, row, col).
  if (ctx.workbook() == nullptr) {
    return Value::error(kPivotRefError);
  }
  const pivot::PivotTable* table = pivot::find_pivot_at_anchor(*ctx.workbook(), sheet_name, anchor->row, anchor->col);
  if (table == nullptr) {
    return Value::error(kPivotRefError);
  }

  // 4. Refresh the result cache on demand. The pivot evaluator owns
  // the lifetime of the produced text values via its `text_storage`
  // deque; we keep that result on the table for subsequent lookups.
  std::shared_ptr<const pivot::PivotResult> result = table->last_result();
  if (!result) {
    const pivot::PivotCache* cache = ctx.workbook()->find_pivot_cache(table->pivot_cache_id());
    if (cache == nullptr) {
      return Value::error(kPivotRefError);
    }
    // The workbook's locale decides how an axis group with no source value is
    // named, and that label is what the (field, item) pairs below match, so it
    // has to be resolved before the hierarchy is built.
    Expected<pivot::PivotResult, Error> evaluated =
        pivot::evaluate(*table, *cache, pivot_layout_options_for(ctx.workbook()->excel_profile()));
    if (!evaluated) {
      // Hide internal evaluator error codes behind the Mac-visible
      // GETPIVOTDATA surface.
      return Value::error(kPivotRefError);
    }
    // Evaluation is intentionally outside the table lock. Multiple cells
    // may compute the same immutable result, but publishing only replaces a
    // shared snapshot; any reader that already acquired one stays valid.
    table->set_last_result(std::move(evaluated.value()));
    result = table->last_result();
  }
  if (!result) {
    return Value::error(kPivotRefError);
  }
  const pivot::PivotResult& pivot_result = *result;

  // 5. Evaluate the optional (field, item) pairs and bucket them by
  // axis. Each pair is keyed by the field's depth on its axis so we
  // can sort them into row_field_order / col_field_order order.
  const std::size_t pair_count = (arity - 2U) / 2U;
  // Initialise paths to empty strings; fill the depths we receive and
  // refuse the lookup later if any depth is left blank (partial path).
  std::vector<std::string> row_path(table->row_field_order().size());
  std::vector<bool> row_path_set(table->row_field_order().size(), false);
  std::vector<std::string> col_path(table->col_field_order().size());
  std::vector<bool> col_path_set(table->col_field_order().size(), false);

  for (std::size_t p = 0; p < pair_count; ++p) {
    const parser::AstNode& field_node = call.as_call_arg(2U + 2U * static_cast<std::uint32_t>(p));
    const parser::AstNode& item_node = call.as_call_arg(2U + 2U * static_cast<std::uint32_t>(p) + 1U);

    const Value fv = eval_node(field_node, arena, registry, ctx);
    if (fv.is_error()) {
      return fv;
    }
    const Value iv = eval_node(item_node, arena, registry, ctx);
    if (iv.is_error()) {
      return iv;
    }

    Expected<std::string, ErrorCode> field_text = coerce_to_text(fv);
    if (!field_text) {
      return Value::error(kPivotRefError);
    }
    Expected<std::string, ErrorCode> item_text = coerce_to_text(iv);
    if (!item_text) {
      return Value::error(kPivotRefError);
    }

    const std::size_t fi = find_field_index(table->fields(), field_text.value());
    if (fi == static_cast<std::size_t>(-1)) {
      return Value::error(kPivotRefError);
    }
    const std::size_t row_depth = depth_in_axis(table->row_field_order(), fi);
    const std::size_t col_depth = depth_in_axis(table->col_field_order(), fi);
    if (row_depth != static_cast<std::size_t>(-1)) {
      if (row_path_set[row_depth]) {
        // Caller addressed the same field twice: ambiguous, refuse.
        return Value::error(kPivotRefError);
      }
      row_path[row_depth] = std::move(item_text.value());
      row_path_set[row_depth] = true;
    } else if (col_depth != static_cast<std::size_t>(-1)) {
      if (col_path_set[col_depth]) {
        return Value::error(kPivotRefError);
      }
      col_path[col_depth] = std::move(item_text.value());
      col_path_set[col_depth] = true;
    } else {
      // Field is on neither row nor col axis (or doesn't exist).
      // Page-axis / data-axis fields are not addressable by GETPIVOTDATA
      // in the MVP -- they'd require routing through filter state,
      // which is deferred. Surface #REF! to mirror Mac's response when
      // the field is on an axis it does not expect.
      return Value::error(kPivotRefError);
    }
  }

  // 6. Locate the data field, then read the value matrix. With no
  // row/col axis fields the value matrix is 1x1xN_data; with axis
  // fields the (row_leaf, col_leaf) coordinates address the correct
  // cell. Either way, the third index is the data-field slot.
  const std::size_t df_idx = find_data_field_index(table->data_fields(), data_field_text.value());
  if (df_idx == static_cast<std::size_t>(-1)) {
    return Value::error(kPivotRefError);
  }

  // 7. Special case: no field/item pairs at all -> grand total. This
  // must run before the per-axis path-completeness check below, which
  // would otherwise reject the bare `=GETPIVOTDATA("Sum of Amount", A3)`
  // shape because the row-axis path slot is unset. Mac's surface for
  // the bare call is the grand total even when axis fields are
  // configured.
  if (pair_count == 0 && (!table->row_field_order().empty() || !table->col_field_order().empty())) {
    if (df_idx < pivot_result.grand_totals.size() && !pivot_result.grand_totals[df_idx].is_blank()) {
      return reify_in_arena(pivot_result.grand_totals[df_idx], arena);
    }
    if (df_idx == 0 && !pivot_result.grand_total.is_blank()) {
      return reify_in_arena(pivot_result.grand_total, arena);
    }
    return Value::error(kPivotRefError);
  }

  // 8. Resolve the row / col leaf indices. With axis fields configured,
  // every depth must be filled; otherwise refuse the lookup.
  std::size_t row_leaf = 0;
  if (!table->row_field_order().empty()) {
    for (bool s : row_path_set) {
      if (!s) {
        return Value::error(kPivotRefError);
      }
    }
    row_leaf = walk_hierarchy<pivot::RowHierarchyNode>(pivot_result.rows, row_path);
    if (row_leaf == static_cast<std::size_t>(-1)) {
      return Value::error(kPivotRefError);
    }
  }
  std::size_t col_leaf = 0;
  if (!table->col_field_order().empty()) {
    for (bool s : col_path_set) {
      if (!s) {
        return Value::error(kPivotRefError);
      }
    }
    col_leaf = walk_hierarchy<pivot::ColHierarchyNode>(pivot_result.cols, col_path);
    if (col_leaf == static_cast<std::size_t>(-1)) {
      return Value::error(kPivotRefError);
    }
  }

  if (row_leaf >= pivot_result.values.size() || col_leaf >= pivot_result.values[row_leaf].size() ||
      df_idx >= pivot_result.values[row_leaf][col_leaf].size()) {
    return Value::error(kPivotRefError);
  }
  return reify_in_arena(pivot_result.values[row_leaf][col_leaf][df_idx], arena);
}

}  // namespace eval
}  // namespace formulon
