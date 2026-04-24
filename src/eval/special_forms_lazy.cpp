// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the short-circuit "special form" family: `IF`,
// `IFERROR`, and `IFNA`. These are routed through the dispatch table in
// `tree_walker.cpp` via the `lazy_impls.h` extern declarations. See that
// header for the dispatch-table contract and the motivation for the
// split.

#include "eval/special_forms_lazy.h"

#include <cstdint>
#include <string_view>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

// IF(cond, then, else?) - then is evaluated iff cond coerces to true; else
// is evaluated iff cond coerces to false. When the third argument is
// omitted Excel returns the boolean `FALSE` for the falsey path.
Value eval_if_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                   const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }
  const Value cond = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (cond.is_error()) {
    return cond;
  }
  auto coerced = coerce_to_bool(cond);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  if (coerced.value()) {
    return eval_node(call.as_call_arg(1), arena, registry, ctx);
  }
  if (arity == 3) {
    return eval_node(call.as_call_arg(2), arena, registry, ctx);
  }
  return Value::boolean(false);
}

// IFERROR(value, fallback) - returns `value` unchanged unless it is any
// error, in which case `fallback` is evaluated and returned. The fallback
// subtree is NOT evaluated when `value` is non-error (true short-circuit).
// If `fallback` itself raises an error it is propagated as-is.
Value eval_iferror_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  if (call.as_call_arity() != 2) {
    return Value::error(ErrorCode::Value);
  }
  const Value primary = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (!primary.is_error()) {
    return primary;
  }
  return eval_node(call.as_call_arg(1), arena, registry, ctx);
}

// IFNA(value, fallback) - returns `value` unchanged unless it is exactly
// `#N/A`, in which case `fallback` is evaluated and returned. All other
// errors (including `#DIV/0!`, `#REF!`, `#VALUE!`, `#NAME?`) propagate as
// `value`. The fallback subtree is NOT evaluated unless the trigger fires.
Value eval_ifna_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 2) {
    return Value::error(ErrorCode::Value);
  }
  // Excel's IFNA coerces a Blank result to number 0 on both the pass-through
  // and the fallback path, matching the implicit Blank->0 promotion that
  // applies when a formula cell's ultimate value is returned to the grid.
  const Value primary = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (!(primary.is_error() && primary.as_error() == ErrorCode::NA)) {
    if (primary.is_blank()) {
      return Value::number(0.0);
    }
    return primary;
  }
  const Value fallback = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (fallback.is_blank()) {
    return Value::number(0.0);
  }
  return fallback;
}

// COUNT(value, ...) - Excel's rule is provenance-sensitive:
//   * A direct scalar argument that evaluates to a Number OR Bool counts.
//   * A cell referenced inside a range contributes only if it holds a
//     Number; range-sourced Bools and text are skipped.
// Implementing this correctly requires per-arg AST inspection, so COUNT is
// routed through the lazy dispatch table. Text, blanks, and errors never
// count in either shape. A direct-arg error is silently skipped (COUNT is
// registered with `propagate_errors = false` in the eager path; this lazy
// impl mirrors that behaviour).
Value eval_count_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1) {
    return Value::error(ErrorCode::Value);
  }
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& arg_node = call.as_call_arg(i);
    if (arg_node.kind() == parser::NodeKind::RangeOp) {
      // Range argument: count only Number cells; Bool / Text / Blank /
      // Error are skipped. Only literal A1:B2 rectangles are expanded, to
      // mirror the eager dispatcher; anything else degrades to #REF! here.
      const parser::AstNode& lhs_ast = arg_node.as_range_lhs();
      const parser::AstNode& rhs_ast = arg_node.as_range_rhs();
      if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
        continue;
      }
      auto expanded = ctx.expand_range(lhs_ast.as_ref(), rhs_ast.as_ref(), arena, registry);
      if (!expanded) {
        continue;
      }
      for (const Value& v : expanded.value()) {
        if (v.is_number()) {
          total += 1.0;
        }
      }
      continue;
    }
    // Direct argument: count Number AND Bool. Text / Blank / Error skip.
    const Value v = eval_node(arg_node, arena, registry, ctx);
    if (v.is_number() || v.is_boolean()) {
      total += 1.0;
    }
  }
  return Value::number(total);
}

// IFS(cond1, val1, ...) - Excel's multi-branch short-circuit: the first
// TRUE condition wins and its paired value is evaluated and returned; all
// remaining branches (both conditions AND values) are skipped. Errors in
// an evaluated condition propagate. When no condition matches (including
// the degenerate odd-arity case) the result is #N/A, matching Excel's
// documented "if none match" behaviour.
Value eval_ifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2) {
    return Value::error(ErrorCode::Value);
  }
  // Iterate in (cond, value) pairs. If the count is odd, the trailing
  // condition has no paired value; we still evaluate it for error
  // propagation, then fall through to #N/A.
  for (std::uint32_t i = 0; i + 1 < arity; i += 2) {
    const Value cond = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (cond.is_error()) {
      return cond;
    }
    auto coerced = coerce_to_bool(cond);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value()) {
      return eval_node(call.as_call_arg(i + 1), arena, registry, ctx);
    }
  }
  if ((arity % 2) == 1) {
    // Trailing unpaired condition: evaluate for error propagation only,
    // then fall through to #N/A regardless of its truth value.
    const Value trailing = eval_node(call.as_call_arg(arity - 1), arena, registry, ctx);
    if (trailing.is_error()) {
      return trailing;
    }
  }
  return Value::error(ErrorCode::NA);
}

// Equality test for SWITCH: matches the `=` operator's semantics for the
// scalar types SWITCH actually consumes. Text comparison is ASCII
// case-insensitive (Excel-canonical). Cross-type pairs never match -- this
// is NOT an error, just "not equal", so the caller keeps walking cases.
bool switch_equal(const Value& lhs, const Value& rhs) {
  if (lhs.kind() != rhs.kind()) {
    return false;
  }
  switch (lhs.kind()) {
    case ValueKind::Number:
      return lhs.as_number() == rhs.as_number();
    case ValueKind::Bool:
      return lhs.as_boolean() == rhs.as_boolean();
    case ValueKind::Text:
      return strings::case_insensitive_eq(lhs.as_text(), rhs.as_text());
    case ValueKind::Blank:
      return true;
    default:
      return false;
  }
}

// SWITCH(expr, case1, val1, ..., [default]) - first case that equals
// `expr` wins; only that branch's value subtree is evaluated. An extra
// trailing argument (odd arity after expr) is the default. No match and
// no default -> #N/A. Errors in `expr` or in any evaluated case expression
// propagate.
Value eval_switch_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Minimum useful form is SWITCH(expr, case, val): 3 args. A bare
  // SWITCH(expr) or SWITCH(expr, default) is rejected as an arity
  // violation (matches Excel's "You've entered too few arguments").
  if (arity < 3) {
    return Value::error(ErrorCode::Value);
  }
  const Value expr = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (expr.is_error()) {
    return expr;
  }
  // Walk (case, value) pairs starting at index 1. If a trailing single
  // argument remains at the end it is the default.
  std::uint32_t i = 1;
  while (i + 1 < arity) {
    const Value case_val = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (case_val.is_error()) {
      return case_val;
    }
    if (switch_equal(expr, case_val)) {
      return eval_node(call.as_call_arg(i + 1), arena, registry, ctx);
    }
    i += 2;
  }
  if (i < arity) {
    // Trailing default argument.
    return eval_node(call.as_call_arg(i), arena, registry, ctx);
  }
  return Value::error(ErrorCode::NA);
}

}  // namespace eval
}  // namespace formulon
