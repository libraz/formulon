// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the short-circuit "special form" family: `IF`,
// `IFERROR`, and `IFNA`. These are routed through the dispatch table in
// `tree_walker.cpp` via the `lazy_impls.h` extern declarations. See that
// header for the dispatch-table contract and the motivation for the
// split.

#include "eval/special_forms_lazy.h"

#include <cstdint>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "utils/arena.h"
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
  const Value primary = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (!(primary.is_error() && primary.as_error() == ErrorCode::NA)) {
    return primary;
  }
  return eval_node(call.as_call_arg(1), arena, registry, ctx);
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

}  // namespace eval
}  // namespace formulon
