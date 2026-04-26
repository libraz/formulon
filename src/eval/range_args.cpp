// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `resolve_range_arg`. See `range_args.h` for the
// public contract.

#include "eval/range_args.h"

#include <algorithm>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/reference_lazy.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

bool resolve_range_arg(const parser::AstNode& raw_arg, Arena& arena, const FunctionRegistry& registry,
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
  // OFFSET(...) and CHOOSE(...) both produce a rectangle that aggregator-
  // family callers (SUM, AVERAGE, COUNTIF, …) must be able to iterate as if
  // it were a literal `RangeOp`. Forward to the dedicated expansion helpers
  // so the expansion shares the same `EvalContext::expand_range` path and
  // cross-sheet / cycle guarantees. Any other Call (INDIRECT, a user-
  // defined function, etc.) falls through to the scalar `#VALUE!` branch
  // below because dynamic range construction requires a `Value::Array`
  // runtime we do not yet have.
  if (arg_node.kind() == parser::NodeKind::Call && strings::case_insensitive_eq(arg_node.as_call_name(), "OFFSET")) {
    return expand_offset_call(arg_node, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
  }
  if (arg_node.kind() == parser::NodeKind::Call && strings::case_insensitive_eq(arg_node.as_call_name(), "CHOOSE")) {
    return expand_choose_call(arg_node, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
  }
  if (arg_node.kind() == parser::NodeKind::Call && strings::case_insensitive_eq(arg_node.as_call_name(), "IF")) {
    // `IF(cond, then, [else])` preserves reference-shape through the picked
    // branch in Mac Excel, so `=LET(r, IF(TRUE, A1:A3, B1:B3), SUM(r))`
    // aggregates the 3-cell range rather than collapsing `r` to a scalar.
    // Short-circuit the condition exactly like `eval_if_lazy`, then recurse
    // into the chosen branch so nested CHOOSE / OFFSET / RangeOp / Ref keep
    // their existing expansion paths. Errors propagate left-to-right (cond
    // first, then the chosen branch), matching Excel and `expand_choose_call`.
    // For the `IF(FALSE, then)` two-arity case Excel's scalar path returns
    // boolean FALSE — not a reference — so we surface `#VALUE!` and let the
    // caller fall back to the scalar branch (mirrors the `IF` block in
    // `resolve_reference_call` added in commit `e068a7f`).
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
    return resolve_range_arg(chosen, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
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
    if (out_rows != nullptr) {
      *out_rows = union_rhs.row - union_lhs.row + 1U;
    }
    if (out_cols != nullptr) {
      *out_cols = union_rhs.col - union_lhs.col + 1U;
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
  // Any other shape — literal, call, arithmetic, array — is not a range
  // and Excel rejects it with #VALUE!.
  *out_err_code = ErrorCode::Value;
  return false;
}

}  // namespace eval
}  // namespace formulon
