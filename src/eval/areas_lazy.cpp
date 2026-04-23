// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the `AREAS` lazy impl. The structural count of
// rectangles in a reference argument is computed by recursing over the
// AST: `UnionOp` children sum, `RangeOp` / `Ref` contribute 1, anything
// else is `#VALUE!`. See `eval/areas_lazy.h` for the dispatch contract.

#include "eval/areas_lazy.h"

#include <cstdint>

#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Recursively sums the leaf rectangles in a reference-shaped AST subtree.
// Returns a negative value on structural mismatch so callers can
// translate it into `#VALUE!` without an extra out-parameter. The walk is
// pure AST inspection; no `eval_node` call is required because AREAS
// answers from shape alone.
std::int64_t count_areas(const parser::AstNode& n) noexcept {
  const parser::NodeKind k = n.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    return 1;
  }
  if (k == parser::NodeKind::UnionOp) {
    std::int64_t total = 0;
    const std::uint32_t arity = n.as_union_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      const std::int64_t c = count_areas(n.as_union_child(i));
      if (c < 0) {
        return -1;
      }
      total += c;
    }
    return total;
  }
  return -1;
}

}  // namespace

Value eval_areas_lazy(const parser::AstNode& call, Arena& /*arena*/, const FunctionRegistry& /*registry*/,
                      const EvalContext& /*ctx*/) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const std::int64_t n = count_areas(call.as_call_arg(0));
  if (n < 0) {
    return Value::error(ErrorCode::Value);
  }
  return Value::number(static_cast<double>(n));
}

}  // namespace eval
}  // namespace formulon
