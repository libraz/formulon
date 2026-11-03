#include "eval/scalar_args.h"

#include <cmath>

#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "utils/arena.h"

namespace formulon {
namespace eval {

bool eval_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, double& out, Value& error_out) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    error_out = Value::error(coerced.error());
    return false;
  }
  out = coerced.value();
  return true;
}

bool eval_bool_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   bool& out, Value& error_out) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  auto coerced = coerce_to_bool(v);
  if (!coerced) {
    error_out = Value::error(coerced.error());
    return false;
  }
  out = coerced.value();
  return true;
}

bool eval_truncated_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, double& out, Value& error_out) {
  double n = 0.0;
  if (!eval_number_arg(node, arena, registry, ctx, n, error_out)) {
    return false;
  }
  out = std::trunc(n);
  return true;
}

}  // namespace eval
}  // namespace formulon
