//
// LET / LAMBDA binding-source resolution: decides whether a binding records a
// reference (so the bound name reads it where it is used) or an evaluated
// value. The public contract is declared in `tree_walker/dispatch.h`.

#include <cstdint>
#include <string_view>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/range_resolvers.h"
#include "eval/tree_walker/dispatch.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// True when `node` reads no lexical binding, so it evaluates the same in any
// scope. Conservative: a shape it does not know is not scope-free.
bool is_scope_free(const parser::AstNode& node, const NameEnv* env) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
    case parser::NodeKind::Ref:
    case parser::NodeKind::Ref3D:
    case parser::NodeKind::ArrayLiteral:
      return true;
    case parser::NodeKind::SpillRef:
      return node.as_spill_ref_anchor_expr() == nullptr;
    case parser::NodeKind::UnaryOp:
      return is_scope_free(node.as_unary_operand(), env);
    case parser::NodeKind::BinaryOp:
      return is_scope_free(node.as_binary_lhs(), env) && is_scope_free(node.as_binary_rhs(), env);
    case parser::NodeKind::RangeOp:
      return is_scope_free(node.as_range_lhs(), env) && is_scope_free(node.as_range_rhs(), env);
    case parser::NodeKind::IntersectOp:
      return is_scope_free(node.as_intersect_lhs(), env) && is_scope_free(node.as_intersect_rhs(), env);
    case parser::NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (!is_scope_free(node.as_union_child(i), env)) {
          return false;
        }
      }
      return true;
    case parser::NodeKind::Call:
      if (env != nullptr && env->lookup(node.as_call_name()) != nullptr) {
        return false;
      }
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        if (!is_scope_free(node.as_call_arg(i), env)) {
          return false;
        }
      }
      return true;
    default:
      return false;
  }
}

// Resolves a binding source that returns a reference (a reference-returning
// call, a `:` over one, or an intersection) to the static `RangeOp` it names,
// so the binding carries that rectangle rather than the call. Returns nullptr
// for any other shape or when the call picks a non-reference (`INDEX({1,2},1)`).
const parser::AstNode* resolve_computed_reference(const parser::AstNode& expr, Arena& arena,
                                                  const FunctionRegistry& registry, const EvalContext& ctx) {
  if (expr.kind() == parser::NodeKind::Call) {
    const NameEnv* env = ctx.name_env();
    if (!is_reference_call_name(expr.as_call_name()) ||
        (env != nullptr && env->lookup(expr.as_call_name()) != nullptr)) {
      return nullptr;
    }
  } else if (expr.kind() == parser::NodeKind::RangeOp) {
    parser::Reference lhs{};
    parser::Reference rhs{};
    if (declared_rect_endpoint_pair(expr, &lhs, &rhs)) {
      return nullptr;
    }
  } else if (expr.kind() != parser::NodeKind::IntersectOp) {
    return nullptr;
  }
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  ErrorCode err = ErrorCode::Value;
  if (!resolve_reference_rect(expr, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &err)) {
    return nullptr;
  }
  parser::Reference first{};
  parser::Reference last{};
  first.sheet = sheet;
  last.sheet = sheet;
  if (top == 0U && bottom == Sheet::kMaxRows - 1U) {
    first.is_full_col = true;
    last.is_full_col = true;
    first.col = left;
    last.col = right;
  } else if (left == 0U && right == Sheet::kMaxCols - 1U) {
    first.is_full_row = true;
    last.is_full_row = true;
    first.row = top;
    last.row = bottom;
  } else {
    first.row = top;
    first.col = left;
    last.row = bottom;
    last.col = right;
  }
  // Always a `RangeOp`, even for one cell (`D1:D1`): a bare `Ref` binding is
  // the scalar-provenance shape, and a reference-returning call is not.
  parser::AstNode* lhs = parser::make_ref(arena, first);
  parser::AstNode* rhs = parser::make_ref(arena, last);
  if (lhs == nullptr || rhs == nullptr) {
    return nullptr;
  }
  return parser::make_range_op(arena, lhs, rhs);
}

}  // namespace

const parser::AstNode* resolve_binding_reference(const parser::AstNode& expr, Arena& arena,
                                                 const FunctionRegistry& registry, const EvalContext& ctx) {
  parser::Reference lhs{};
  parser::Reference rhs{};
  switch (expr.kind()) {
    case parser::NodeKind::Ref:
      return &expr;
    case parser::NodeKind::RangeOp:
      if (declared_rect_endpoint_pair(expr, &lhs, &rhs)) {
        return &expr;
      }
      break;
    case parser::NodeKind::NameRef: {
      const NameEnv* env = ctx.name_env();
      const parser::AstNode* bound =
          (env != nullptr && expr.as_name_sheet().empty()) ? env->lookup_ast(expr.as_name()) : nullptr;
      if (bound != nullptr && (bound->kind() == parser::NodeKind::Ref || bound->kind() == parser::NodeKind::RangeOp)) {
        return bound;
      }
      return nullptr;
    }
    default:
      break;
  }
  return resolve_computed_reference(expr, arena, registry, ctx);
}

Value eval_binding_source(const parser::AstNode& expr, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, const parser::AstNode** out_ast) {
  const parser::AstNode* ref = resolve_binding_reference(expr, arena, registry, ctx);
  *out_ast = ref;
  // A union, or a range-shaped call naming no single rectangle (one picking a
  // union), is recorded as written when it reads the same in any scope.
  const bool unresolved_range =
      ref == nullptr && (expr.kind() == parser::NodeKind::UnionOp ||
                         (expr.kind() == parser::NodeKind::Call && is_range_shaped_ast(expr)));
  if (expr.kind() == parser::NodeKind::ArrayLiteral ||
      (expr.kind() == parser::NodeKind::SpillRef && expr.as_spill_ref_anchor_expr() == nullptr) ||
      (unresolved_range && is_scope_free(expr, ctx.name_env()))) {
    *out_ast = &expr;
  } else if (expr.kind() == parser::NodeKind::NameRef && expr.as_name_sheet().empty() && ctx.name_env() != nullptr) {
    *out_ast = ctx.name_env()->lookup_ast(expr.as_name());
  }
  // A reference is read where the bound name is used, not here.
  if (ref != nullptr) {
    return Value::blank();
  }
  return eval_node(expr, arena, registry, ctx);
}

}  // namespace eval
}  // namespace formulon
