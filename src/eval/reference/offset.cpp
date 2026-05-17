// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the OFFSET lazy impl plus the related range-shape
// expanders (`expand_offset_call`, `expand_choose_call`, `expand_if_call`,
// `expand_row_call`, `expand_column_call`).
//
// The rectangle-construction core (`compute_offset_rect`, `OffsetBase`)
// lives in `reference/common.cpp` because the intersection resolver
// (`reference/intersection.cpp`) reaches it too. This TU only owns the
// OFFSET evaluator and the range-expander dispatch on top.

#include "eval/reference/offset.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/reference/common.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

Value eval_offset_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  refs_internal::OffsetBase base{};
  std::uint32_t top_row = 0;
  std::uint32_t left_col = 0;
  std::uint32_t height = 0;
  std::uint32_t width = 0;
  ErrorCode err = ErrorCode::Value;
  if (!refs_internal::compute_offset_rect(call, arena, registry, ctx, &base, &top_row, &left_col, &height, &width,
                                          &err)) {
    return Value::error(err);
  }
  (void)height;
  (void)width;
  // Scalar context for a multi-cell OFFSET: Excel 365 dynamic-array
  // semantics spill the rectangle, and a reader that samples only the
  // anchor cell (as xlwings does in the oracle pipeline) sees the
  // top-left value. Aggregators hit a different path (they expand the
  // rectangle via `expand_offset_call` wired into `resolve_range_arg`),
  // so returning the top-left here only matters for direct scalar
  // consumption, where it reproduces Mac Excel 365's observable output.
  parser::Reference target{};
  target.sheet = base.sheet;
  target.row = top_row;
  target.col = left_col;
  return ctx.resolve_ref(target, arena, registry);
}

bool expand_offset_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols) {
  refs_internal::OffsetBase base{};
  std::uint32_t top_row = 0;
  std::uint32_t left_col = 0;
  std::uint32_t height = 0;
  std::uint32_t width = 0;
  ErrorCode err = ErrorCode::Value;
  if (!refs_internal::compute_offset_rect(call, arena, registry, ctx, &base, &top_row, &left_col, &height, &width,
                                          &err)) {
    *out_err_code = err;
    return false;
  }
  // Build two synthetic endpoint references delimiting the rectangle and
  // hand them off to `EvalContext::expand_range`, which already handles
  // cross-sheet routing, cycle detection, and per-cell recursion.
  parser::Reference lhs{};
  parser::Reference rhs{};
  lhs.sheet = base.sheet;
  lhs.row = top_row;
  lhs.col = left_col;
  rhs.sheet = base.sheet;
  rhs.row = top_row + height - 1U;
  rhs.col = left_col + width - 1U;
  auto expanded = ctx.expand_range(lhs, rhs, arena, registry);
  if (!expanded) {
    *out_err_code = expanded.error();
    return false;
  }
  *out_cells = std::move(expanded.value());
  if (out_rows != nullptr) {
    *out_rows = height;
  }
  if (out_cols != nullptr) {
    *out_cols = width;
  }
  return true;
}

bool expand_choose_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols) {
  const std::uint32_t arity = call.as_call_arity();
  // Need at least the index plus one value, matching `eval_choose_lazy`.
  if (arity < 2U) {
    *out_err_code = ErrorCode::Value;
    return false;
  }
  // Evaluate the index argument; CHOOSE expects a 1-based integer
  // selector. Errors propagate with their original code, mirroring
  // `eval_choose_lazy`.
  const Value idx_val = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (idx_val.is_error()) {
    *out_err_code = idx_val.as_error();
    return false;
  }
  auto idx_num = coerce_to_number(idx_val);
  if (!idx_num) {
    *out_err_code = idx_num.error();
    return false;
  }
  // Excel truncates (toward zero) rather than rounds: CHOOSE(2.9, ...)
  // selects the 2nd value, not the 3rd. For valid `[1, arity-1]` indices
  // these are non-negative, so `std::floor` matches `eval_choose_lazy`.
  const double raw = std::floor(idx_num.value());
  if (!(raw >= 1.0 && raw <= static_cast<double>(arity - 1U))) {
    *out_err_code = ErrorCode::Value;
    return false;
  }
  const auto picked_slot = static_cast<std::uint32_t>(raw);
  const parser::AstNode& chosen = call.as_call_arg(picked_slot);
  // Recurse so nested OFFSET / CHOOSE chains also flatten cleanly. Any
  // other shape (Ref, RangeOp, …) falls through to `resolve_range_arg`,
  // which already knows how to handle them — including the
  // "anything else -> #VALUE!" fallthrough for scalar children.
  if (chosen.kind() == parser::NodeKind::Call) {
    const std::string_view name = chosen.as_call_name();
    if (strings::case_insensitive_eq(name, "OFFSET")) {
      return expand_offset_call(chosen, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
    }
    if (strings::case_insensitive_eq(name, "CHOOSE")) {
      return expand_choose_call(chosen, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
    }
  }
  auto resolved = resolve_range_arg(chosen, arena, registry, ctx);
  if (!resolved) {
    *out_err_code = resolved.error();
    return false;
  }
  auto& rr = resolved.value();
  if (out_rows != nullptr) {
    *out_rows = rr.rows;
  }
  if (out_cols != nullptr) {
    *out_cols = rr.cols;
  }
  *out_cells = std::move(rr.cells);
  return true;
}

bool expand_if_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    std::vector<Value>* out_cells, ErrorCode* out_err_code, std::uint32_t* out_rows,
                    std::uint32_t* out_cols) {
  const std::uint32_t arity = call.as_call_arity();
  // `IF(cond, then, [else])` — Mac Excel preserves reference-shape through
  // the picked branch, so `=LET(r, IF(TRUE, A1:A3, B1:B3), SUM(r))` should
  // aggregate the 3-cell range rather than collapse to a scalar. Mirror
  // `eval_if_lazy`'s short-circuit semantics: evaluate cond first (errors
  // propagate left-to-right matching Excel), then recurse into the picked
  // branch. The `IF(FALSE, then)` two-arity case returns boolean FALSE in
  // Excel's scalar path — not a reference — so we surface `#VALUE!` and
  // let the caller fall back to its scalar branch (see commit `e068a7f`'s
  // `resolve_reference_call` IF block for the matching reasoning).
  if (arity != 2U && arity != 3U) {
    *out_err_code = ErrorCode::Value;
    return false;
  }
  const Value cond = eval_node(call.as_call_arg(0), arena, registry, ctx);
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
  const parser::AstNode& chosen = call.as_call_arg(pick);
  // Recurse so nested OFFSET / CHOOSE / IF chains also flatten cleanly.
  // Any other shape (Ref, RangeOp, …) falls through to `resolve_range_arg`,
  // which already knows how to handle them — including the
  // "anything else -> #VALUE!" fallthrough for scalar children.
  if (chosen.kind() == parser::NodeKind::Call) {
    const std::string_view name = chosen.as_call_name();
    if (strings::case_insensitive_eq(name, "OFFSET")) {
      return expand_offset_call(chosen, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
    }
    if (strings::case_insensitive_eq(name, "CHOOSE")) {
      return expand_choose_call(chosen, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
    }
    if (strings::case_insensitive_eq(name, "IF")) {
      return expand_if_call(chosen, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
    }
  }
  auto resolved = resolve_range_arg(chosen, arena, registry, ctx);
  if (!resolved) {
    *out_err_code = resolved.error();
    return false;
  }
  auto& rr = resolved.value();
  if (out_rows != nullptr) {
    *out_rows = rr.rows;
  }
  if (out_cols != nullptr) {
    *out_cols = rr.cols;
  }
  *out_cells = std::move(rr.cells);
  return true;
}

namespace {

// Shared body of `expand_row_call` / `expand_column_call`. `want_row`
// picks the axis: when true, fills `out_cells` with 1-based row indices
// drawn from `[top..bottom]` of the resolved rectangle and reports
// `(rows = bottom-top+1, cols = 1)`; when false, fills with column
// indices from `[left..right]` and reports `(rows = 1, cols = right-left+1)`.
//
// The shape inspection mirrors `eval_row_or_column` in `shape_ops_lazy.cpp`:
// LET-bound NameRefs are looked through, single-cell `Ref` becomes a 1x1,
// `RangeOp(Ref, Ref)` covers the full row / column span, and a nested
// reference-returning `Call` (INDIRECT / OFFSET / IF / CHOOSE) routes
// through `resolve_reference_call`. Anything else evaluates the subtree
// to surface errors and otherwise reports `#VALUE!`. The 0-arity branch
// (bare `=ROW()` / `=COLUMN()`) emits a single 1x1 indexed by the
// formula cell, or `#VALUE!` if no formula cell is bound.
bool expand_row_or_column_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, bool want_row, std::vector<Value>* out_cells,
                               ErrorCode* out_err_code, std::uint32_t* out_rows, std::uint32_t* out_cols) {
  const std::uint32_t arity = call.as_call_arity();
  out_cells->clear();
  if (arity == 0U) {
    if (!ctx.has_formula_cell()) {
      *out_err_code = ErrorCode::Value;
      return false;
    }
    const std::uint32_t idx = want_row ? ctx.formula_row() : ctx.formula_col();
    out_cells->push_back(Value::number(static_cast<double>(idx + 1U)));
    if (out_rows != nullptr) {
      *out_rows = 1U;
    }
    if (out_cols != nullptr) {
      *out_cols = 1U;
    }
    return true;
  }
  if (arity != 1U) {
    *out_err_code = ErrorCode::Value;
    return false;
  }

  // LET-binding passthrough: `=LET(r, A1:A3, SUM(ROW(r)))` parses `r` as
  // a NameRef. Mirror `eval_row_or_column`'s rule: accept the broader
  // "Ref OR range-shaped" set so a single-cell binding still yields a
  // meaningful row / column index.
  const parser::AstNode& raw_arg = call.as_call_arg(0);
  const parser::AstNode* effective = &raw_arg;
  if (raw_arg.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(raw_arg, ctx.name_env());
    if (&resolved != &raw_arg && (resolved.kind() == parser::NodeKind::Ref || is_range_shaped_ast(resolved))) {
      effective = &resolved;
    }
  }
  const parser::AstNode& arg = *effective;

  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  bool resolved_rect = false;

  const parser::NodeKind k = arg.kind();
  if (k == parser::NodeKind::Ref) {
    const parser::Reference& r = arg.as_ref();
    top = bottom = r.row;
    left = right = r.col;
    resolved_rect = true;
  } else if (k == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = arg.as_range_lhs();
    const parser::AstNode& rhs_ast = arg.as_range_rhs();
    if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
      *out_err_code = ErrorCode::Value;
      return false;
    }
    const parser::Reference& lhs = lhs_ast.as_ref();
    const parser::Reference& rhs = rhs_ast.as_ref();
    top = std::min(lhs.row, rhs.row);
    bottom = std::max(lhs.row, rhs.row);
    left = std::min(lhs.col, rhs.col);
    right = std::max(lhs.col, rhs.col);
    resolved_rect = true;
  } else if (k == parser::NodeKind::Call) {
    std::string_view sheet;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    if (resolve_reference_call(arg, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range, &err)) {
      resolved_rect = true;
    } else {
      // Fall through to the scalar-evaluate branch so subtree errors propagate.
    }
  }

  if (resolved_rect) {
    if (want_row) {
      const std::uint32_t height = bottom - top + 1U;
      out_cells->reserve(height);
      for (std::uint32_t r = top; r <= bottom; ++r) {
        out_cells->push_back(Value::number(static_cast<double>(r + 1U)));
      }
      if (out_rows != nullptr) {
        *out_rows = height;
      }
      if (out_cols != nullptr) {
        *out_cols = 1U;
      }
    } else {
      const std::uint32_t width = right - left + 1U;
      out_cells->reserve(width);
      for (std::uint32_t c = left; c <= right; ++c) {
        out_cells->push_back(Value::number(static_cast<double>(c + 1U)));
      }
      if (out_rows != nullptr) {
        *out_rows = 1U;
      }
      if (out_cols != nullptr) {
        *out_cols = width;
      }
    }
    return true;
  }

  // Evaluate the subtree to surface any errors verbatim (e.g. `ROW(1/0)`),
  // otherwise reject non-references with `#VALUE!`.
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    *out_err_code = v.as_error();
    return false;
  }
  *out_err_code = ErrorCode::Value;
  return false;
}

}  // namespace

bool expand_row_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                     std::uint32_t* out_rows, std::uint32_t* out_cols) {
  return expand_row_or_column_call(call, arena, registry, ctx, /*want_row=*/true, out_cells, out_err_code, out_rows,
                                   out_cols);
}

bool expand_column_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols) {
  return expand_row_or_column_call(call, arena, registry, ctx, /*want_row=*/false, out_cells, out_err_code, out_rows,
                                   out_cols);
}

}  // namespace eval
}  // namespace formulon
