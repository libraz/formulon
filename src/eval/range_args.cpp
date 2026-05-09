// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `resolve_range_arg`. See `range_args.h` for the
// public contract.

#include "eval/range_args.h"

#include <algorithm>
#include <array>
#include <cstdint>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/shape_ops_lazy.h"
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

// Internal counterpart that still uses the legacy `bool + out_param`
// shape so it can call (and be called by) the cluster of `expand_*_call`
// helpers — which are not yet migrated. The public `resolve_range_arg`
// below is a thin Expected-returning wrapper around this. Once the
// rest of the family migrates, this helper folds away.
bool resolve_range_arg_into(const parser::AstNode& raw_arg, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                            std::uint32_t* out_rows, std::uint32_t* out_cols) {
  // LET-binding passthrough: when a caller wrote `VLOOKUP(key, t, 2, FALSE)`
  // with `t` bound to a RangeOp / OFFSET-call / ArrayLiteral via LET, the
  // shape decisions below need the original AST, not the NameRef. Single-
  // cell Refs and scalar bindings are intentionally left as-is so the
  // existing 1-cell / scalar-fallback semantics are preserved.
  const parser::AstNode* effective = &raw_arg;
  if (raw_arg.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(raw_arg, ctx.name_env());
    if (&resolved != &raw_arg && is_range_shaped_ast(resolved)) {
      effective = &resolved;
    }
  }
  const parser::AstNode& arg_node = *effective;
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
          return expand_offset_call(arg_node, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
        case RangeShapedKind::Choose:
          return expand_choose_call(arg_node, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
        case RangeShapedKind::Row:
          // ROW(range) / COLUMN(range) spill in Excel 365 to a vertical /
          // horizontal array of 1-based indices. Without a `Value::Array`
          // runtime the scalar path collapses to the rectangle's first
          // row / column, so the seam here unpacks the indices directly
          // into the aggregator's range buffer.
          return expand_row_call(arg_node, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
        case RangeShapedKind::Column:
          return expand_column_call(arg_node, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
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
          const std::uint32_t arity = arg_node.as_call_arity();
          if (arity != 2U && arity != 3U) {
            *out_err_code = ErrorCode::Value;
            return false;
          }
          const Value cond = eval_node(arg_node.as_call_arg(0), arena, registry, ctx);
          if (cond.is_error()) {
            *out_err_code = cond.as_error();
            return false;
          }
          auto coerced = coerce_to_bool(cond);
          if (!coerced) {
            *out_err_code = coerced.error();
            return false;
          }
          if (!coerced.value() && arity == 2U) {
            *out_err_code = ErrorCode::Value;
            return false;
          }
          const std::uint32_t pick = coerced.value() ? 1U : 2U;
          const parser::AstNode& chosen = arg_node.as_call_arg(pick);
          return resolve_range_arg_into(chosen, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
        }
      }
    }
  }
  if (arg_node.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = arg_node.as_range_lhs();
    const parser::AstNode& rhs_ast = arg_node.as_range_rhs();
    // Endpoints may be plain Refs or reference-producing calls
    // (`OFFSET(...)` / `INDIRECT(...)`). `resolve_range_endpoint`
    // normalises both shapes to a rectangle so we can union them and
    // feed `expand_range` two synthetic Refs. Sheet-qualifier
    // validation (mismatched qualifiers -> #REF!) is delegated to
    // `expand_range` itself.
    std::string_view lhs_sheet;
    std::string_view rhs_sheet;
    std::uint32_t lhs_top = 0;
    std::uint32_t lhs_left = 0;
    std::uint32_t lhs_bottom = 0;
    std::uint32_t lhs_right = 0;
    std::uint32_t rhs_top = 0;
    std::uint32_t rhs_left = 0;
    std::uint32_t rhs_bottom = 0;
    std::uint32_t rhs_right = 0;
    ErrorCode endpoint_err = ErrorCode::Ref;
    if (!resolve_range_endpoint(lhs_ast, arena, registry, ctx, &lhs_sheet, &lhs_top, &lhs_left, &lhs_bottom, &lhs_right,
                                &endpoint_err) ||
        !resolve_range_endpoint(rhs_ast, arena, registry, ctx, &rhs_sheet, &rhs_top, &rhs_left, &rhs_bottom, &rhs_right,
                                &endpoint_err)) {
      *out_err_code = endpoint_err;
      return false;
    }
    parser::Reference union_lhs{};
    parser::Reference union_rhs{};
    union_lhs.sheet = lhs_sheet;
    union_lhs.row = std::min(lhs_top, rhs_top);
    union_lhs.col = std::min(lhs_left, rhs_left);
    union_rhs.sheet = rhs_sheet;
    union_rhs.row = std::max(lhs_bottom, rhs_bottom);
    union_rhs.col = std::max(lhs_right, rhs_right);
    auto expanded = ctx.expand_range(union_lhs, union_rhs, arena, registry);
    if (!expanded) {
      *out_err_code = expanded.error();
      return false;
    }
    *out_cells = std::move(expanded.value());
    // Defensive normalisation: while the union is constructed with min /
    // max above (so union_rhs >= union_lhs is the documented invariant),
    // a future refactor that stops using min / max — or a malformed
    // endpoint resolver returning a degenerate rectangle — would silently
    // wrap the unsigned subtraction below into a multi-billion shape.
    // Recompute both axes through min/max so the dimension is always
    // positive regardless of which endpoint became the union top-left.
    if (out_rows != nullptr) {
      const std::uint32_t r_lo = std::min(union_lhs.row, union_rhs.row);
      const std::uint32_t r_hi = std::max(union_lhs.row, union_rhs.row);
      *out_rows = r_hi - r_lo + 1U;
    }
    if (out_cols != nullptr) {
      const std::uint32_t c_lo = std::min(union_lhs.col, union_rhs.col);
      const std::uint32_t c_hi = std::max(union_lhs.col, union_rhs.col);
      *out_cols = c_hi - c_lo + 1U;
    }
    return true;
  }
  if (arg_node.kind() == parser::NodeKind::Ref) {
    // Single-cell Ref: treat as a 1-element range so COUNTIF(A1, ">0") is
    // well-defined. Error / blank surface via `resolve_ref` as a Value and
    // are forwarded unchanged; the matcher handles them correctly.
    out_cells->clear();
    out_cells->push_back(ctx.resolve_ref(arg_node.as_ref(), arena, registry));
    if (out_rows != nullptr) {
      *out_rows = 1U;
    }
    if (out_cols != nullptr) {
      *out_cols = 1U;
    }
    return true;
  }
  if (arg_node.kind() == parser::NodeKind::SpillRef) {
    // Spilled-range `A1#`: resolve the spill region anchored at the
    // reference and copy its row-major cells into the output buffer.
    // Mirrors the dispatcher's SpillRef branch in `tree_walker.cpp` so any
    // range-aware consumer (lookup, conditional aggregator, regression,
    // workdays, …) accepts a SpillRef passed through a LET binding without
    // collapsing to its anchor scalar.
    const parser::Reference& sr = arg_node.as_spill_ref();
    const Sheet* current = ctx.current_sheet();
    if (current == nullptr) {
      *out_err_code = ErrorCode::Name;
      return false;
    }
    const Sheet* target = current;
    if (!sr.sheet.empty()) {
      const Workbook* wb = ctx.workbook();
      if (wb == nullptr) {
        *out_err_code = ErrorCode::Ref;
        return false;
      }
      target = wb->sheet_by_name(sr.sheet);
      if (target == nullptr) {
        *out_err_code = ErrorCode::Ref;
        return false;
      }
    }
    if (sr.row >= Sheet::kMaxRows || sr.col >= Sheet::kMaxCols) {
      *out_err_code = ErrorCode::Ref;
      return false;
    }
    const SpillRegion* region = target->spill_region_at_anchor(sr.row, sr.col);
    if (region == nullptr) {
      *out_err_code = ErrorCode::Ref;
      return false;
    }
    out_cells->clear();
    out_cells->reserve(region->cells.size());
    for (const Value& v : region->cells) {
      out_cells->push_back(v);
    }
    if (out_rows != nullptr) {
      *out_rows = region->rows;
    }
    if (out_cols != nullptr) {
      *out_cols = region->cols;
    }
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
  // Reference-shaped nodes (`RangeOp` / `Ref` / `SpillRef` / `OFFSET` /
  // `CHOOSE` / `IF` / `ROW` / `COLUMN`) never reach this branch — their
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
    out_cells->clear();
    out_cells->reserve(total);
    for (std::size_t i = 0; i < total; ++i) {
      out_cells->push_back(src[i]);
    }
    if (out_rows != nullptr) {
      *out_rows = rows;
    }
    if (out_cols != nullptr) {
      *out_cols = cols;
    }
    return true;
  }
  // Scalar value (Number / Bool / Text / Blank / Lambda): treat as a 1x1
  // range so single-argument aggregators (`=SUM(7)`, `=AVERAGE(A1+1)`)
  // behave as Excel does instead of failing.
  out_cells->clear();
  out_cells->push_back(result);
  if (out_rows != nullptr) {
    *out_rows = 1U;
  }
  if (out_cols != nullptr) {
    *out_cols = 1U;
  }
  return true;
}

}  // namespace

Expected<RangeResult, ErrorCode> resolve_range_arg(const parser::AstNode& arg_node, Arena& arena,
                                                   const FunctionRegistry& registry, const EvalContext& ctx) {
  RangeResult result;
  ErrorCode err = ErrorCode::Value;
  if (!resolve_range_arg_into(arg_node, arena, registry, ctx, &result.cells, &err, &result.rows, &result.cols)) {
    return err;
  }
  return result;
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
  const parser::NodeKind k = arg_node.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
    if (!resolved) {
      // `resolve_range_arg` reports `#VALUE!` for non-Ref / non-RangeOp
      // shapes and `#REF!` for expansion failures. The regression and
      // hypothesis-test families use `#N/A` for shape errors, so remap
      // the shape-rejection case while letting `#REF!` pass through.
      const ErrorCode err_code = resolved.error();
      return err_code == ErrorCode::Value ? ErrorCode::NA : err_code;
    }
    return std::move(resolved.value());
  }
  if (k == parser::NodeKind::ArrayLiteral) {
    RangeResult out;
    out.rows = arg_node.as_array_rows();
    out.cols = arg_node.as_array_cols();
    const std::size_t total = static_cast<std::size_t>(out.rows) * out.cols;
    out.cells.reserve(total);
    for (std::uint32_t r = 0; r < out.rows; ++r) {
      for (std::uint32_t c = 0; c < out.cols; ++c) {
        out.cells.push_back(eval_node(arg_node.as_array_element(r, c), arena, registry, ctx));
      }
    }
    return out;
  }
  // Scalar / arithmetic / Call subtree. Evaluate so any pre-existing
  // error propagates with its real code; otherwise reject with `#N/A`.
  const Value v = eval_node(arg_node, arena, registry, ctx);
  if (v.is_error()) {
    return v.as_error();
  }
  return ErrorCode::NA;
}

}  // namespace eval
}  // namespace formulon
