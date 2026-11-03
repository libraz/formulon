// Scalar argument readers shared by the lazy function families.

#ifndef FORMULON_EVAL_SCALAR_ARGS_H_
#define FORMULON_EVAL_SCALAR_ARGS_H_

#include <cstdint>
#include <vector>

#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

class Arena;

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Evaluate `node` and coerce it to a number. On error or coercion failure
/// writes the caller-visible error to `error_out` and returns `false`.
bool eval_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, double& out, Value& error_out);

/// Evaluate `node` and coerce it through `coerce_to_bool`. On error or
/// coercion failure writes the caller-visible error to `error_out` and
/// returns `false`.
bool eval_bool_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   bool& out, Value& error_out);

/// Evaluate `node`, coerce to a finite number, and truncate toward zero.
/// On error or coercion failure writes the caller-visible error to
/// `error_out` and returns `false`. Shared by every helper that takes a
/// numeric scalar argument (counts, indices, ignore masks, etc.).
bool eval_truncated_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, double& out, Value& error_out);

/// Evaluates the scalar arguments of a call with `min_arity`..`max_arity`
/// left to right into `out`. Returns the left-most argument error, or blank
/// on success. With `atp_rule` the Analysis-ToolPak rule applies: trailing
/// omitted optional slots are dropped, and each slot rejects a boolean unless
/// bit `i` of `logical_args` is set (see `atp_arg_error`).
Value eval_scalar_call_args(const parser::AstNode& call, std::uint32_t min_arity, std::uint32_t max_arity,
                            bool atp_rule, std::uint32_t logical_args, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::vector<Value>& out);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SCALAR_ARGS_H_
