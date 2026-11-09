//
// Implementation of `resolve_range_arg`. See `range_args.h` for the
// public contract.

#include "eval/range_args.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/array_context.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/special_forms_lazy.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// `propagate_errors = false` is what lets `COUNT` / `IS*` inspect an error
// argument instead of inheriting it, and it has a consequence worth stating
// where the filter applies it: an error that reached the value list because
// an *expansion* failed is indistinguishable here from one that was in the
// data, and both are simply dropped as "not a number". A function with this
// flag therefore converts a failed argument expansion into a plausible wrong
// count rather than into a visible error, so a gap in the expanders surfaces
// as silence on exactly these functions.
bool filter_range_sourced_value(const FunctionDef& def, const Value& value, Value* out, bool* keep, Value* out_err) {
  if (def.propagate_errors && value.is_error()) {
    *out_err = value;
    return false;
  }
  if (def.range_filter_numeric_only && value.kind() != ValueKind::Number) {
    *keep = false;
    return true;
  }
  if (def.range_filter_bool_coercible && value.kind() != ValueKind::Number && value.kind() != ValueKind::Bool) {
    *keep = false;
    return true;
  }
  if (def.range_filter_a_coerce) {
    if (value.kind() == ValueKind::Blank) {
      *keep = false;
      return true;
    }
    if (value.kind() == ValueKind::Bool) {
      *out = Value::number(value.as_boolean() ? 1.0 : 0.0);
      *keep = true;
      return true;
    }
    if (value.kind() == ValueKind::Text) {
      *out = Value::number(0.0);
      *keep = true;
      return true;
    }
  }
  *out = value;
  *keep = true;
  return true;
}

/// Range-shaped Call-name dispatch. `OFFSET` / `CHOOSE` produce
/// rectangles directly; `IF` preserves shape through the picked branch;
/// `ROW` / `COLUMN` spill to a 1-D index array. Each kind routes to a
/// dedicated expansion path below — the lookup table replaces a 5-way
/// `if (case_insensitive_eq(name, ...))` chain so the case-insensitive
/// compare runs at most twice (early-out on first hit) instead of five
/// times in the worst case.
enum class RangeShapedKind : std::uint8_t { Offset, Choose, If, Row, Column };

constexpr std::array<std::pair<std::string_view, RangeShapedKind>, 5> kRangeShapedNames = {{
    {"OFFSET", RangeShapedKind::Offset},
    {"CHOOSE", RangeShapedKind::Choose},
    {"IF", RangeShapedKind::If},
    {"ROW", RangeShapedKind::Row},
    {"COLUMN", RangeShapedKind::Column},
}};

bool lookup_range_shaped_kind(std::string_view name, RangeShapedKind* out) {
  for (const auto& entry : kRangeShapedNames) {
    if (strings::case_insensitive_eq(name, entry.first)) {
      *out = entry.second;
      return true;
    }
  }
  return false;
}

// Expands `[lhs : rhs]` into `out_cells` and records the rectangle's shape.
bool assign_expanded(const EvalContext& ctx, const parser::Reference& lhs, const parser::Reference& rhs, Arena& arena,
                     const FunctionRegistry& registry, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                     std::uint32_t& out_rows, std::uint32_t& out_cols) {
  auto expanded = ctx.expand_range(lhs, rhs, arena, registry, &out_rows, &out_cols);
  if (!expanded) {
    *out_err_code = expanded.error();
    return false;
  }
  *out_cells = std::move(expanded.value());
  return true;
}

// `bool + out_param` core shared with the `expand_*_call` helpers; the public
// `resolve_range_arg` below wraps it in an Expected.
//
// `out_scalar` — when non-null — reports whether the result came from the
// bare-scalar collapse at the bottom of the function rather than from a
// genuine rectangle. Families whose argument slot documents a reference
// or array need that distinction to keep rejecting `=IRR(5)`.
bool resolve_range_arg_into(const parser::AstNode& raw_arg, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                            std::uint32_t& out_rows, std::uint32_t& out_cols, bool* out_scalar = nullptr) {
  if (out_scalar != nullptr) {
    *out_scalar = false;
  }
  // LET-binding passthrough: when a caller wrote `VLOOKUP(key, t, 2, FALSE)`
  // with `t` bound to a RangeOp / OFFSET-call / ArrayLiteral via LET, the
  // shape decisions below need the original AST, not the NameRef. Single-
  // cell Refs and scalar bindings are intentionally left as-is so the
  // existing 1-cell / scalar-fallback semantics are preserved.
  const parser::AstNode& arg_node = resolve_range_binding(raw_arg, ctx.name_env(), /*accept_ref=*/false);
  // 3-D reference: the cells of the span, flattened (no rectangle shape).
  if (arg_node.kind() == parser::NodeKind::Ref3D) {
    if (is_grouped_ref3d(arg_node)) {
      *out_err_code = ErrorCode::Name;
      return false;
    }
    out_cells->clear();
    if (!expand_ref3d_cells(arg_node, arena, registry, ctx, out_cells, out_err_code)) {
      return false;
    }
    out_rows = static_cast<std::uint32_t>(out_cells->size());
    out_cols = 1U;
    return true;
  }
  // OFFSET / CHOOSE / IF / ROW / COLUMN all need range-shaped expansion
  // glue (see per-branch comments below). Dispatch via a single
  // case-insensitive name lookup against `kRangeShapedNames` so the hot
  // path runs at most one full string compare per matching prefix —
  // dramatically cheaper than the previous five-way `if`-chain when none
  // of the names match (the common case for plain `RangeOp` / `Ref`
  // arguments). Any other Call (INDIRECT, a user-defined function, etc.)
  // falls through to the scalar-evaluation branch at the bottom of the
  // function because dynamic range construction requires a `Value::Array`
  // runtime we do not yet have.
  if (arg_node.kind() == parser::NodeKind::Call) {
    RangeShapedKind kind = RangeShapedKind::Offset;
    if (lookup_range_shaped_kind(arg_node.as_call_name(), &kind)) {
      switch (kind) {
        case RangeShapedKind::Offset:
          return expand_offset_call(arg_node, arena, registry, ctx, out_cells, out_err_code, &out_rows, &out_cols);
        case RangeShapedKind::Choose:
          return expand_choose_call(arg_node, arena, registry, ctx, out_cells, out_err_code, &out_rows, &out_cols);
        case RangeShapedKind::Row:
          // ROW(range) / COLUMN(range) spill in Excel 365 to a vertical /
          // horizontal array of 1-based indices. Without a `Value::Array`
          // runtime the scalar path collapses to the rectangle's first
          // row / column, so the seam here unpacks the indices directly
          // into the aggregator's range buffer.
          return expand_row_call(arg_node, arena, registry, ctx, out_cells, out_err_code, &out_rows, &out_cols);
        case RangeShapedKind::Column:
          return expand_column_call(arg_node, arena, registry, ctx, out_cells, out_err_code, &out_rows, &out_cols);
        case RangeShapedKind::If: {
          // `IF(cond, then, [else])` preserves reference-shape through the
          // picked branch in Mac Excel, so `=LET(r, IF(TRUE, A1:A3, B1:B3),
          // SUM(r))` aggregates the 3-cell range rather than collapsing `r`
          // to a scalar. Short-circuit the condition exactly like
          // `eval_if_lazy`, then recurse into the chosen branch so nested
          // CHOOSE / OFFSET / RangeOp / Ref keep their existing expansion
          // paths. Errors propagate left-to-right (cond first, then the
          // chosen branch), matching Excel and `expand_choose_call`. For
          // the `IF(FALSE, then)` two-arity case Excel's scalar path
          // returns boolean FALSE — not a reference — so we surface
          // `#VALUE!` and let the caller fall back to the scalar branch
          // (mirrors the `IF` block in `resolve_reference_call`).
          const SelectedBranch picked = select_if_branch(arg_node, arena, registry, ctx);
          if (picked.kind == SelectedBranch::Kind::kArraySubject) {
            // An array condition picks per cell instead of short-circuiting, so
            // `SUMPRODUCT(IF(A1:A5<=3, A1:A5, 0))` aggregates the masked column.
            const Value result = eval_if_array_cond_lazy(arg_node, picked.value, arena, registry, ctx);
            return expand_array_result(result, out_cells, out_err_code, &out_rows, &out_cols);
          }
          if (picked.kind == SelectedBranch::Kind::kResult) {
            // `IF(FALSE, then)` is the scalar FALSE, not a reference.
            *out_err_code = picked.value.is_error() ? picked.value.as_error() : ErrorCode::Value;
            return false;
          }
          const parser::AstNode& chosen = arg_node.as_call_arg(picked.arg);
          return resolve_range_arg_into(chosen, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols,
                                        out_scalar);
        }
      }
    }
  }
  if (arg_node.kind() == parser::NodeKind::IntersectOp) {
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    bool disjoint = false;
    ErrorCode intersect_err = ErrorCode::Value;
    if (!compute_intersect_rect(arg_node.as_intersect_lhs(), arg_node.as_intersect_rhs(), arena, registry, ctx, &sheet,
                                &top, &left, &bottom, &right, &disjoint, &intersect_err)) {
      *out_err_code = intersect_err;
      return false;
    }
    if (disjoint) {
      *out_err_code = ErrorCode::Null;
      return false;
    }
    return expand_resolved_rect_cells(sheet, top, left, bottom, right, arena, registry, ctx, out_cells, out_err_code,
                                      &out_rows, &out_cols);
  }
  if (arg_node.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = arg_node.as_range_operand(0);
    const parser::AstNode& rhs_ast = arg_node.as_range_operand(1);
    // Multi-column (`A:C`) / multi-row (`1:3`) whole references parse as a
    // RangeOp over two whole-column / whole-row Refs. `resolve_range_endpoint`
    // deliberately rejects whole references (they have no bounded rectangle
    // on their own), so route the pair straight to `expand_range`, which
    // clamps the unbounded axis to the sheet's used range and reports the
    // concrete shape.
    if (arg_node.as_range_arity() == 2U && lhs_ast.kind() == parser::NodeKind::Ref &&
        rhs_ast.kind() == parser::NodeKind::Ref) {
      const parser::Reference& lhs_ref = lhs_ast.as_ref();
      const parser::Reference& rhs_ref = rhs_ast.as_ref();
      const bool lhs_full = lhs_ref.is_full_col || lhs_ref.is_full_row;
      const bool rhs_full = rhs_ref.is_full_col || rhs_ref.is_full_row;
      if (lhs_full || rhs_full) {
        return assign_expanded(ctx, lhs_ref, rhs_ref, arena, registry, out_cells, out_err_code, out_rows, out_cols);
      }
    } else {
      // A `:` chain of three or more endpoints, or one with a non-`Ref`
      // endpoint, is no bare pair and the test above does not see it. The chain
      // names one rectangle -- the bounding box of every endpoint -- and
      // the shared reduction spells it as the endpoint pair that
      // `expand_range` already knows how to enumerate. A chain containing
      // a reference-returning call reduces to nothing and falls through to
      // the endpoint union below.
      parser::Reference chain_lhs{};
      parser::Reference chain_rhs{};
      if (declared_rect_endpoint_pair(arg_node, &chain_lhs, &chain_rhs, ctx.workbook())) {
        return assign_expanded(ctx, chain_lhs, chain_rhs, arena, registry, out_cells, out_err_code, out_rows, out_cols);
      }
    }
    // Endpoints may be plain Refs, reference-producing calls
    // (`OFFSET(...)` / `INDIRECT(...)`) or names standing for either.
    // `resolve_range_endpoint` normalises every shape to a rectangle so we
    // can union them and feed `expand_range` two synthetic Refs on the
    // sheet `merge_range_endpoint_sheets` settles.
    ErrorCode endpoint_err = ErrorCode::Ref;
    std::string_view union_sheet;
    std::uint32_t union_top = 0;
    std::uint32_t union_left = 0;
    std::uint32_t union_bottom = 0;
    std::uint32_t union_right = 0;
    if (!resolve_range_endpoints(arg_node, arena, registry, ctx, &union_sheet, &union_top, &union_left, &union_bottom,
                                 &union_right, &endpoint_err)) {
      *out_err_code = endpoint_err;
      return false;
    }
    return expand_resolved_rect_cells(union_sheet, union_top, union_left, union_bottom, union_right, arena, registry,
                                      ctx, out_cells, out_err_code, &out_rows, &out_cols);
  }
  if (arg_node.kind() == parser::NodeKind::Ref) {
    const parser::Reference& ref = arg_node.as_ref();
    if (ref.is_full_col || ref.is_full_row) {
      // Whole-column (`A:A`) / whole-row (`1:1`) reference: expand against
      // the sheet's used range so range-aware consumers (SUM / COUNTIF /
      // lookup / dynamic-array) see the populated cells. `expand_range`
      // clamps the unbounded axis and reports the concrete shape.
      return assign_expanded(ctx, ref, ref, arena, registry, out_cells, out_err_code, out_rows, out_cols);
    }
    // Single-cell Ref: treat as a 1-element range so COUNTIF(A1, ">0") is
    // well-defined. Error / blank surface via `resolve_ref` as a Value and
    // are forwarded unchanged; the matcher handles them correctly.
    out_cells->clear();
    const Value cell = ctx.resolve_ref(ref, arena, registry);
    out_cells->push_back(cell.is_blank() ? Value::blank(BlankGridProjection::kReferenceGridZero) : cell);
    out_rows = 1U;
    out_cols = 1U;
    return true;
  }
  if (arg_node.kind() == parser::NodeKind::SpillRef) {
    // Spilled-range `A1#`: resolve the spill region anchored at the
    // reference and copy its row-major cells into the output buffer.
    // Mirrors the dispatcher's SpillRef branch in `tree_walker.cpp` so any
    // range-aware consumer (lookup, conditional aggregator, regression,
    // workdays, …) accepts a SpillRef passed through a LET binding without
    // collapsing to its anchor scalar.
    std::string_view anchor_sheet;
    std::uint32_t anchor_row = 0;
    std::uint32_t anchor_col = 0;
    if (!resolve_spill_anchor_node(arg_node, arena, registry, ctx, &anchor_sheet, &anchor_row, &anchor_col,
                                   out_err_code)) {
      return false;
    }
    const Sheet* current = ctx.current_sheet();
    if (current == nullptr) {
      *out_err_code = ErrorCode::Name;
      return false;
    }
    const Sheet* target = current;
    if (!anchor_sheet.empty()) {
      const Workbook* wb = ctx.workbook();
      if (wb == nullptr) {
        *out_err_code = ErrorCode::Ref;
        return false;
      }
      target = wb->sheet_by_name(anchor_sheet);
      if (target == nullptr) {
        *out_err_code = ErrorCode::Ref;
        return false;
      }
    }
    if (anchor_row >= Sheet::kMaxRows || anchor_col >= Sheet::kMaxCols) {
      *out_err_code = ErrorCode::Ref;
      return false;
    }
    // The copied cells outlive the region, so the read re-homes any Text
    // payload into `arena` while the sheet lock is held.
    out_cells->clear();
    if (!ctx.read_spill_region(*target, anchor_row, anchor_col, arena, *out_cells, &out_rows, &out_cols)) {
      *out_err_code = ErrorCode::Ref;
      return false;
    }
    return true;
  }
  if (arg_node.kind() == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = arg_node.as_array_rows();
    const std::uint32_t cols = arg_node.as_array_cols();
    out_cells->clear();
    out_cells->reserve(static_cast<std::size_t>(rows) * cols);
    for (std::uint32_t row = 0; row < rows; ++row) {
      for (std::uint32_t col = 0; col < cols; ++col) {
        out_cells->push_back(eval_node(arg_node.as_array_element(row, col), arena, registry, ctx));
      }
    }
    out_rows = rows;
    out_cols = cols;
    return true;
  }
  // Generic fallback: evaluate the expression and inspect the resulting
  // `Value`. Dynamic-array producers (MUNIT, SEQUENCE, RANDARRAY, MAP,
  // REDUCE, BYROW, BYCOL, MAKEARRAY, LAMBDA invocations, ...) return a
  // `Value::Array` here, which we unpack row-major into `out_cells` so
  // INDEX / SUMPRODUCT / MATCH / aggregators navigate the rectangle as if
  // it had been written as a literal range. Errors propagate; bare scalars
  // (Number / Bool / Text / Blank) collapse to a 1x1 range, which fixes
  // `=SUM(<scalar>)`-style formulas that previously surfaced #VALUE!.
  // Reference-shaped nodes (`RangeOp` / `Ref` / `SpillRef` / `ArrayLiteral`
  // / `OFFSET` / `CHOOSE` / `IF` / `ROW` / `COLUMN`) never reach this branch — their
  // dedicated expansion paths above handle them without re-evaluation.
  const Value result = eval_node(arg_node, arena, registry, ctx);
  if (result.is_error()) {
    *out_err_code = result.as_error();
    return false;
  }
  if (result.is_array()) {
    const std::uint32_t rows = result.as_array_rows();
    const std::uint32_t cols = result.as_array_cols();
    const Value* src = result.as_array_cells();
    const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
    out_cells->assign(src, src + total);
    out_rows = rows;
    out_cols = cols;
    return true;
  }
  // Scalar value (Number / Bool / Text / Blank / Lambda): treat as a 1x1
  // range so single-argument aggregators (`=SUM(7)`, `=AVERAGE(A1+1)`)
  // behave as Excel does instead of failing.
  if (out_scalar != nullptr) {
    *out_scalar = true;
  }
  out_cells->clear();
  out_cells->push_back(result);
  out_rows = 1U;
  out_cols = 1U;
  return true;
}

}  // namespace

bool is_range_shaped_call_name(std::string_view name) {
  RangeShapedKind kind{};
  return lookup_range_shaped_kind(name, &kind);
}

bool expand_resolved_rect_cells(std::string_view sheet, std::uint32_t top, std::uint32_t left, std::uint32_t bottom,
                                std::uint32_t right, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                                std::uint32_t* out_rows, std::uint32_t* out_cols) {
  parser::Reference lhs{};
  parser::Reference rhs{};
  set_rect_corner_refs(sheet, top, left, bottom, right, &lhs, &rhs);
  std::uint32_t expanded_rows = 0;
  std::uint32_t expanded_cols = 0;
  if (!assign_expanded(ctx, lhs, rhs, arena, registry, out_cells, out_err_code, expanded_rows, expanded_cols)) {
    return false;
  }
  if (out_rows != nullptr) {
    *out_rows = bottom - top + 1U;
  }
  if (out_cols != nullptr) {
    *out_cols = right - left + 1U;
  }
  return true;
}

bool append_range_sourced_value(const FunctionDef& def, const Value& value, std::vector<Value>* values,
                                Value* out_err) {
  Value filtered = Value::blank();
  bool keep = false;
  if (!filter_range_sourced_value(def, value, &filtered, &keep, out_err)) {
    return false;
  }
  if (keep) {
    values->push_back(filtered);
  }
  return true;
}

bool append_range_sourced_values(const FunctionDef& def, const Value* cells, std::size_t count,
                                 std::vector<Value>* values, Value* out_err) {
  for (std::size_t i = 0; i < count; ++i) {
    if (!append_range_sourced_value(def, cells[i], values, out_err)) {
      return false;
    }
  }
  return true;
}

bool append_array_sourced_value(const FunctionDef& def, const Value& value, std::vector<Value>* values,
                                Value* out_err) {
  if (def.range_filter_a_coerce && !value.is_error() && !value.is_number()) {
    // Logical elements are skipped; text elements are skipped too, except
    // that AVERAGEA still counts a text element as zero.
    if (value.is_text() && def.canonical_name == "AVERAGEA") {
      values->push_back(Value::number(0.0));
    }
    return true;
  }
  return append_range_sourced_value(def, value, values, out_err);
}

bool append_array_sourced_values(const FunctionDef& def, const Value* cells, std::size_t count,
                                 std::vector<Value>* values, Value* out_err) {
  for (std::size_t i = 0; i < count; ++i) {
    if (!append_array_sourced_value(def, cells[i], values, out_err)) {
      return false;
    }
  }
  return true;
}

bool expand_ref3d_cells(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out, ErrorCode* out_err) {
  *out_err = ErrorCode::Ref;
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr) {
    return false;
  }
  const parser::Reference& cell = node.as_ref3d_cell();
  const bool is_range = node.as_ref3d_is_range();
  const parser::Reference& cell_end = node.as_ref3d_cell_end();
  const std::size_t begin_idx = wb->sheet_index_by_name(node.as_ref3d_sheet_begin());
  const std::size_t end_idx = wb->sheet_index_by_name(node.as_ref3d_sheet_end());
  const bool full_col = cell.is_full_col || (is_range && cell_end.is_full_col);
  const bool full_row = cell.is_full_row || (is_range && cell_end.is_full_row);
  const bool incompatible_whole_shape =
      (full_col && full_row) ||
      (is_range && (cell.is_full_col != cell_end.is_full_col || cell.is_full_row != cell_end.is_full_row));
  const bool corners_out_of_bounds =
      (full_col && (cell.col >= Sheet::kMaxCols || (is_range && cell_end.col >= Sheet::kMaxCols))) ||
      (full_row && (cell.row >= Sheet::kMaxRows || (is_range && cell_end.row >= Sheet::kMaxRows))) ||
      (!full_col && !full_row &&
       (cell.row >= Sheet::kMaxRows || cell.col >= Sheet::kMaxCols ||
        (is_range && (cell_end.row >= Sheet::kMaxRows || cell_end.col >= Sheet::kMaxCols))));
  if (begin_idx != Workbook::kNoSheet && end_idx == Workbook::kNoSheet) {
    *out_err = ErrorCode::Name;
    return false;
  }
  if (begin_idx == Workbook::kNoSheet || end_idx == Workbook::kNoSheet || incompatible_whole_shape ||
      corners_out_of_bounds) {
    return false;
  }
  const std::size_t lo = std::min(begin_idx, end_idx);
  const std::size_t hi = std::max(begin_idx, end_idx);
  // Tail rectangle: a single cell, or the normalised `cell:cell_end`
  // area. Excel aggregates the same rectangle from every sheet in the
  // span, so the cross-product (sheets * area cells) flows into the
  // range-aware function.
  for (std::size_t s = lo; s <= hi; ++s) {
    const std::string_view sheet_name = wb->sheet(s).name();
    // Each sheet walks through EvalContext, so whole-axis bounds follow the 2-D extent and overlay rules.
    parser::Reference lhs = cell;
    parser::Reference rhs = is_range ? cell_end : cell;
    lhs.sheet = sheet_name;
    rhs.sheet = sheet_name;
    const auto walked = ctx.walked_range_rect(lhs, rhs);
    if (!walked) {
      return false;
    }
    if (!walked.value().has_value()) {
      continue;
    }
    const DeclaredRect& rect = *walked.value();
    const std::size_t needed = out->size() + static_cast<std::size_t>(rect.rows()) * rect.cols();
    if (needed > out->capacity()) {
      out->reserve(std::max(needed, out->capacity() * 2U));
    }
    for (std::uint32_t r = rect.row_first; r <= rect.row_last; ++r) {
      for (std::uint32_t c = rect.col_first; c <= rect.col_last; ++c) {
        parser::Reference per_sheet{};
        per_sheet.sheet = sheet_name;
        per_sheet.row = r;
        per_sheet.col = c;
        out->push_back(ctx.resolve_ref(per_sheet, arena, registry));
      }
    }
  }
  return true;
}

Expected<RangeResult, ErrorCode> resolve_range_arg(const parser::AstNode& arg_node, Arena& arena,
                                                   const FunctionRegistry& registry, const EvalContext& ctx) {
  RangeResult result;
  ErrorCode err = ErrorCode::Value;
  if (!resolve_range_arg_into(arg_node, arena, registry, ctx, &result.cells, &err, result.rows, result.cols,
                              &result.from_scalar)) {
    return err;
  }
  return result;
}

bool resolve_flat_cells(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, Value* out_err) {
  auto resolved = resolve_range_arg_no_scalar(arg_node, arena, registry, ctx, ErrorCode::Value);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  *out_cells = std::move(resolved.value().cells);
  return true;
}

const Value* first_error(const std::vector<Value>& cells) {
  for (const Value& v : cells) {
    if (v.is_error()) {
      return &v;
    }
  }
  return nullptr;
}

Expected<RangeResult, ErrorCode> resolve_range_arg_no_scalar(const parser::AstNode& arg_node, Arena& arena,
                                                             const FunctionRegistry& registry, const EvalContext& ctx,
                                                             ErrorCode scalar_error) {
  auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
  if (!resolved) {
    return std::move(resolved.error());
  }
  if (resolved.value().from_scalar) {
    return scalar_error;
  }
  return std::move(resolved.value());
}

bool resolve_array_value(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, const ArrayValue** out, Value* out_err) {
  const Value v = eval_node_as_array(arg, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (!v.is_array()) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = v.as_array();
  return true;
}

Expected<RangeResult, ErrorCode> resolve_array_arg_na(const parser::AstNode& arg_node, Arena& arena,
                                                      const FunctionRegistry& registry, const EvalContext& ctx) {
  auto resolved = resolve_range_arg_no_scalar(arg_node, arena, registry, ctx, ErrorCode::NA);
  if (!resolved) {
    // `resolve_range_arg` reports `#VALUE!` for a subtree it cannot turn
    // into a rectangle and `#REF!` for expansion failures. The regression
    // and hypothesis-test families use `#N/A` for shape errors, so remap
    // the shape-rejection case while letting every other code (`#REF!`,
    // a propagated `#DIV/0!`, …) pass through.
    const ErrorCode err_code = resolved.error();
    return err_code == ErrorCode::Value ? ErrorCode::NA : err_code;
  }
  return std::move(resolved.value());
}

namespace {

// A static reference among a call's arguments whose declared rectangle spans
// a whole grid axis.
struct AxisSpan {
  parser::Reference lhs;
  parser::Reference rhs;
  bool column = false;
  bool whole_axis = false;
};

// Collects the axis-spanning static references `node` reads without a call
// dispatch of its own: operands of operators, LET-bound names, and the
// reference-shaped calls `resolve_range_arg` expands in place. Any other call
// is dispatched, and so computes its own shared extent.
void collect_axis_spans(const parser::AstNode& node, const EvalContext& ctx, std::uint32_t depth,
                        std::vector<AxisSpan>* out) {
  constexpr std::uint32_t kMaxDepth = 32U;
  if (depth > kMaxDepth) {
    return;
  }
  if (node.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& bound = resolve_name_ast(node, ctx.name_env());
    if (&bound != &node) {
      collect_axis_spans(bound, ctx, depth + 1U, out);
    }
    return;
  }
  AxisSpan span;
  if (declared_rect_endpoint_pair(node, &span.lhs, &span.rhs, ctx.workbook())) {
    const auto rect = declared_rect(span.lhs, span.rhs);
    if (!rect) {
      return;
    }
    span.whole_axis = rect.value().whole_axis;
    if (rect.value().rows() == Sheet::kMaxRows) {
      span.column = true;
      out->push_back(span);
    } else if (rect.value().cols() == Sheet::kMaxCols) {
      out->push_back(span);
    }
    return;
  }
  switch (node.kind()) {
    case parser::NodeKind::BinaryOp:
      for (std::uint32_t i = 0; i < node.as_binary_arity(); ++i) {
        collect_axis_spans(node.as_binary_operand(i), ctx, depth + 1U, out);
      }
      return;
    case parser::NodeKind::UnaryOp:
      collect_axis_spans(node.as_unary_operand(), ctx, depth + 1U, out);
      return;
    case parser::NodeKind::Call:
      if (is_range_shaped_ast(node)) {
        for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
          collect_axis_spans(node.as_call_arg(i), ctx, depth + 1U, out);
        }
      }
      return;
    default:
      return;
  }
}

// Walks the spans on one axis to their longest length, recording in
// `*shared` where that length was measured. A bounded span already covers
// the whole axis. Whole-axis spans are measured once per sheet, over the
// bounding band of that sheet's spans: the band's populated extent is the
// longest of theirs, or longer only by cells blank in all of them.
void share_longest_walk(const std::vector<AxisSpan>& spans, bool column, const EvalContext& ctx,
                        WholeAxisExtent* shared) {
  std::uint32_t& length = column ? shared->rows : shared->cols;
  std::vector<AxisSpan> bands;
  for (const AxisSpan& span : spans) {
    if (span.column != column) {
      continue;
    }
    if (!span.whole_axis) {
      length = column ? Sheet::kMaxRows : Sheet::kMaxCols;
      return;
    }
    const std::string_view sheet = span.lhs.sheet.empty() ? span.rhs.sheet : span.lhs.sheet;
    const auto same_sheet = [sheet](const AxisSpan& band) { return band.lhs.sheet == sheet; };
    const auto band = std::find_if(bands.begin(), bands.end(), same_sheet);
    const std::uint32_t first = column ? std::min(span.lhs.col, span.rhs.col) : std::min(span.lhs.row, span.rhs.row);
    const std::uint32_t last = column ? std::max(span.lhs.col, span.rhs.col) : std::max(span.lhs.row, span.rhs.row);
    if (band == bands.end()) {
      AxisSpan fresh = span;
      fresh.lhs.sheet = sheet;
      fresh.rhs.sheet = sheet;
      (column ? fresh.lhs.col : fresh.lhs.row) = first;
      (column ? fresh.rhs.col : fresh.rhs.row) = last;
      bands.push_back(fresh);
      continue;
    }
    std::uint32_t& band_first = column ? band->lhs.col : band->lhs.row;
    std::uint32_t& band_last = column ? band->rhs.col : band->rhs.row;
    band_first = std::min(band_first, first);
    band_last = std::max(band_last, last);
  }
  for (const AxisSpan& band : bands) {
    // A reference that cannot be walked reports its own error when the
    // callee resolves it.
    const Sheet* sheet = nullptr;
    const auto walked = ctx.walked_range_rect(band.lhs, band.rhs, &sheet);
    if (!walked) {
      continue;
    }
    const std::uint32_t band_length =
        !walked.value().has_value() ? 0U : (column ? walked.value()->rows() : walked.value()->cols());
    length = std::max(length, band_length);
    if (bands.size() == 1U) {
      (column ? shared->rows_sheet : shared->cols_sheet) = sheet;
      (column ? shared->rows_band_first : shared->cols_band_first) = column ? band.lhs.col : band.lhs.row;
      (column ? shared->rows_band_last : shared->cols_band_last) = column ? band.rhs.col : band.rhs.row;
    }
  }
}

}  // namespace

bool static_reference_shape(const parser::AstNode& arg_node, const EvalContext& ctx, std::uint32_t* out_rows,
                            std::uint32_t* out_cols) {
  parser::Reference lhs{};
  parser::Reference rhs{};
  if (!declared_rect_endpoint_pair(resolve_name_ast(arg_node, ctx.name_env()), &lhs, &rhs, ctx.workbook())) {
    return false;
  }
  const auto rect = ctx.declared_range_rect(lhs, rhs);
  if (!rect) {
    return false;
  }
  *out_rows = rect.value().rows();
  *out_cols = rect.value().cols();
  return true;
}

WholeAxisExtent shared_whole_axis_extent(const parser::AstNode& call, const EvalContext& ctx) {
  std::vector<AxisSpan> spans;
  for (std::uint32_t i = 0; i < call.as_call_arity(); ++i) {
    collect_axis_spans(call.as_call_arg(i), ctx, 0U, &spans);
  }
  const auto columns = std::count_if(spans.begin(), spans.end(), [](const AxisSpan& span) { return span.column; });
  const auto rows = static_cast<std::ptrdiff_t>(spans.size()) - columns;
  // A lone span has nothing to agree with; its own extent is walked as is.
  WholeAxisExtent shared;
  if (columns >= 2) {
    share_longest_walk(spans, /*column=*/true, ctx, &shared);
  }
  if (rows >= 2) {
    share_longest_walk(spans, /*column=*/false, ctx, &shared);
  }
  return shared;
}

}  // namespace eval
}  // namespace formulon
