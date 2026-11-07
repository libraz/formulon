#include "eval/scalar_args.h"

#include <cmath>

#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "parser/ast.h"
#include "utils/arena.h"

namespace formulon {
namespace eval {

bool number_from_value(const Value& v, double& out, Value& error_out) {
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

bool bool_from_value(const Value& v, bool& out, Value& error_out) {
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

bool eval_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, double& out, Value& error_out) {
  return number_from_value(eval_node(node, arena, registry, ctx), out, error_out);
}

bool eval_bool_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   bool& out, Value& error_out) {
  return bool_from_value(eval_node(node, arena, registry, ctx), out, error_out);
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

Value eval_scalar_call_args(const parser::AstNode& call, std::uint32_t min_arity, std::uint32_t max_arity,
                            bool atp_rule, std::uint32_t logical_args, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::vector<Value>& out) {
  const std::uint32_t evaluated = atp_rule ? atp_evaluated_arity(call, min_arity, max_arity) : call.as_call_arity();
  out.clear();
  out.reserve(evaluated);
  for (std::uint32_t i = 0; i < evaluated; ++i) {
    Value v = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    if (atp_rule) {
      const bool logical = ((logical_args >> i) & 1U) != 0U;
      const Value atp = atp_arg_error(call.as_call_arg(i), v, atp_slot_required(i, min_arity, max_arity), logical);
      if (atp.is_error()) {
        return atp;
      }
    }
    out.push_back(v);
  }
  return Value::blank();
}

}  // namespace eval
}  // namespace formulon
