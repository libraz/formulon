// Scalar argument readers shared by the lazy function families.

#ifndef FORMULON_EVAL_SCALAR_ARGS_H_
#define FORMULON_EVAL_SCALAR_ARGS_H_

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

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SCALAR_ARGS_H_
