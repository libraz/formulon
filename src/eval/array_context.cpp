//
// Array-context evaluation seam: `eval_node_as_array` and the cellwise binary
// broadcaster. See `eval/array_context.h`.

#include "eval/array_context.h"

#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/range_resolvers.h"
#include "eval/scalar_ops.h"
#include "eval/tree_walker/broadcast.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace detail {

// Builds a freshly arena-allocated `ArrayValue` from a flat row-major cell
// vector and shape. The cells are copied into a fresh arena buffer so the
// caller's vector can go out of scope safely; both the buffer and the
// `ArrayValue` header live in `arena` for the same lifetime contract as
// `Value::Text`. Returns `nullptr` when the shape is rejected or the arena
// is exhausted; every caller maps that to `#NUM!`.
const ArrayValue* make_array_value(Arena& arena, std::uint32_t rows, std::uint32_t cols,
                                   const std::vector<Value>& cells) {
  return array_from_values(rows, cols, cells.data(), cells.size(), arena);
}

// `Value::array(nullptr)` is not a legal value, so every `make_array_value`
// call site funnels its result through this guard rather than repeating the
// null check.
Value array_value_or_num_error(const ArrayValue* arr) {
  return arr == nullptr ? Value::error(ErrorCode::Num) : Value::array(arr);
}

// True for the reference-shaped `Call` names that `resolve_range_arg`
// already knows how to expand into a flat row-major rectangle. Kept
// alongside the array-context dispatch so the two seams stay in sync
// when a new reference-producing builtin (XLOOKUP-style spilling, etc.)
// is added.
bool is_range_producing_call(const parser::AstNode& call_node) {
  if (call_node.kind() != parser::NodeKind::Call) {
    return false;
  }
  return is_range_shaped_call_name(call_node.as_call_name());
}

}  // namespace detail

namespace {
using detail::array_value_or_num_error;
using detail::is_range_producing_call;
using detail::make_array_value;
}  // namespace

Value eval_node_as_array(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  // LET-binding NameRef passthrough: mirror the pattern used by the other
  // shape-aware seams so `=LET(r, A1:A3, r+1)` sees the bound RangeOp / Call
  // AST rather than the NameRef wrapper.
  const parser::AstNode& target = resolve_range_binding(node, ctx.name_env(), /*accept_ref=*/true);
  const parser::NodeKind k = target.kind();

  // BinaryOp -> recurse into the cellwise broadcaster.
  if (k == parser::NodeKind::BinaryOp) {
    return eval_binop_array_ctx(target, arena, registry, ctx);
  }

  // UnaryOp -> evaluate operand as array, then apply cellwise.
  if (k == parser::NodeKind::UnaryOp) {
    const Value inner = eval_node_as_array(target.as_unary_operand(), arena, registry, ctx);
    if (inner.is_error()) {
      return inner;
    }
    // `eval_node_as_array` is contracted to return either an Array or a
    // scalar Error; any other shape would be a bug in this seam.
    if (!inner.is_array()) {
      return Value::error(ErrorCode::Value);
    }
    const ArrayValue* in_arr = inner.as_array();
    const std::size_t n = static_cast<std::size_t>(in_arr->rows) * static_cast<std::size_t>(in_arr->cols);
    Value* buf = nullptr;
    ArrayValue* out = allocate_derived_array_value(in_arr->rows, in_arr->cols, arena, buf);
    if (out == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    for (std::size_t i = 0; i < n; ++i) {
      buf[i] = apply_unary(target.as_unary_op(), in_arr->cells[i]);
    }
    return Value::array(out);
  }

  // Range-shaped AST: Ref / RangeOp / one of the reference-producing Calls.
  // `resolve_range_arg` already knows how to expand each of these into a
  // flat row-major vector + shape, so reuse it to keep the expansion path
  // identical to the conditional-aggregator and lookup families. An
  // ArrayLiteral keeps its per-cell errors instead of short-circuiting.
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp || is_range_producing_call(target) ||
      k == parser::NodeKind::ArrayLiteral) {
    auto resolved = resolve_range_arg(target, arena, registry, ctx);
    if (!resolved) {
      return Value::error(resolved.error());
    }
    auto& rr = resolved.value();
    return array_value_or_num_error(make_array_value(arena, rr.rows, rr.cols, rr.cells));
  }

  // Scalar fallback. Evaluate normally, then wrap into a 1x1 array. Errors
  // stay scalar so they can short-circuit the broadcaster in the caller.
  // A pre-existing Array (some future array-producing builtin) is forwarded
  // unchanged so eval_node_as_array remains idempotent on its outputs.
  const Value v = eval_node(target, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  if (v.is_array()) {
    return v;
  }
  Value* buf = nullptr;
  ArrayValue* out = allocate_derived_array_value(1U, 1U, arena, buf);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  buf[0] = v;
  return Value::array(out);
}

Value eval_binop_array_ctx(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  // Folded left to right; the left-most error wins between scalars, and a
  // scalar error against an array is broadcast error per position.
  Value lhs = eval_node_as_array(node.as_binary_operand(0), arena, registry, ctx);
  for (std::uint32_t i = 1; i < node.as_binary_arity(); ++i) {
    const Value rhs = eval_node_as_array(node.as_binary_operand(i), arena, registry, ctx);
    if (lhs.is_error() && !rhs.is_array()) {
      return lhs;
    }
    if (rhs.is_error() && !lhs.is_array()) {
      return rhs;
    }
    // Both operands are guaranteed Array post the array-context contract.
    // Delegate to the single Excel-broadcasting implementation shared with the
    // top-level BinaryOp dispatch (`tree_walker/broadcast.cpp`), which handles
    // size-1 broadcast, the outer product, RxC/Rx1/1xC combinations, and #N/A
    // padding for mismatched non-1 axes.
    lhs = broadcast_binop(node.as_binary_op(i - 1U), lhs, rhs, arena);
  }
  return lhs;
}

}  // namespace eval
}  // namespace formulon
