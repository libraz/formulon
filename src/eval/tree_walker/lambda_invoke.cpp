//
// Runtime invocation of `LambdaValue`s: parameter binding, the arity and
// lambda-depth rules, callable resolution for the lazy lambda helpers, and
// reference resolution through a lambda body. The public contract is declared
// in `tree_walker/dispatch.h`.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/omitted_arg.h"
#include "eval/range_resolvers.h"
#include "eval/tree_walker/depth_guard.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Builds the eta-expansion of the built-in `name` at `arity`:
// `LAMBDA(p1, ..., pn, name(p1, ..., pn))`. The parameter names cannot be
// spelled in a formula, so the body can only ever see its own arguments.
// Returns nullptr on arena exhaustion.
const LambdaValue* eta_expand(std::string_view name, std::uint32_t arity, Arena& arena) {
  std::string_view* params = nullptr;
  const parser::AstNode** body_args = nullptr;
  if (arity > 0U) {
    params = arena.create_array<std::string_view>(arity);
    body_args = arena.create_array<const parser::AstNode*>(arity);
    if (params == nullptr || body_args == nullptr) {
      return nullptr;
    }
  }
  for (std::uint32_t i = 0; i < arity; ++i) {
    params[i] = arena.intern("@" + std::to_string(i + 1U));
    body_args[i] = parser::make_name_ref(arena, params[i]);
    if (params[i].empty() || body_args[i] == nullptr) {
      return nullptr;
    }
  }
  const parser::AstNode* body = parser::make_call(arena, name, body_args, arity);
  auto* lv = arena.create<LambdaValue>();
  if (body == nullptr || lv == nullptr) {
    return nullptr;
  }
  lv->params = params;
  lv->param_count = arity;
  lv->optional_count = 0U;
  lv->body = body;
  lv->captured_env = nullptr;
  return lv;
}

// The lambda a call of `lv` with `arity` arguments runs: `lv` itself, or the
// eta-expansion of the built-in a function value names. Null on arena
// exhaustion.
const LambdaValue* at_arity(const LambdaValue* lv, std::uint32_t arity, Arena& arena) {
  return lv->builtin.empty() ? lv : eta_expand(lv->builtin, arity, arena);
}

// `syntax_args` is the call site's argument AST, consulted for one thing
// only: telling a syntactically omitted slot (`f(1, , 3)`) from a supplied
// one. It is deliberately separate from `ast_args`, the per-argument AST
// `eval_binding_source` chose to record on the binding.
// Binds `lv`'s parameters for a call into `*env`; the error a call surfaces
// instead when the arity does not fit or the lambda has no body.
std::optional<ErrorCode> bind_lambda_params(const LambdaValue* lv, std::uint32_t arity, const Value* args,
                                            const parser::AstNode* const* ast_args,
                                            const parser::AstNode* const* syntax_args, Arena& arena, NameEnv* out) {
  const std::uint32_t required = lv->param_count - lv->optional_count;
  if (arity < required || arity > lv->param_count) {
    return ErrorCode::Value;
  }
  if (lv->body == nullptr) {
    return ErrorCode::Name;
  }
  NameEnv& env = *out;
  if (lv->captured_env != nullptr) {
    env = *lv->captured_env;
  }
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode* expr = (ast_args != nullptr) ? ast_args[i] : nullptr;
    // `f(1, , 3)` supplies the slot syntactically and omits its value, and
    // Excel treats that as omitted wherever it appears -- leading, middle
    // or trailing -- not only past the end of the argument list. Binding
    // it as an ordinary blank would give the body the right number (an
    // omitted slot arithmetic-coerces to 0 either way) while telling
    // ISOMITTED the wrong thing, which is the whole of what that function
    // is for. Only the AST can answer this: by the time the argument is a
    // Value, an omitted slot and a blank cell are the same blank.
    const parser::AstNode* syntax = (syntax_args != nullptr) ? syntax_args[i] : expr;
    if (syntax != nullptr && is_omitted_arg(*syntax)) {
      env = env.extend_omitted(lv->params[i], arena);
      continue;
    }
    env = env.extend(lv->params[i], args[i], expr, arena);
  }
  for (std::uint32_t i = arity; i < lv->param_count; ++i) {
    env = env.extend_omitted(lv->params[i], arena);
  }
  return std::nullopt;
}

EvalContext lambda_body_context(const LambdaValue* lv, const NameEnv& env, const EvalContext& ctx) {
  EvalContext body_ctx = ctx.with_name_env(&env);
  if (lv->name_scope_sheet >= 0) {
    body_ctx = body_ctx.with_name_scope_sheet(lv->name_scope_sheet);
  }
  return body_ctx;
}

Value invoke_lambda_values_impl(const LambdaValue* lv, std::uint32_t arity, const Value* args,
                                const parser::AstNode* const* ast_args, const parser::AstNode* const* syntax_args,
                                Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  NameEnv env;
  if (const std::optional<ErrorCode> err = bind_lambda_params(lv, arity, args, ast_args, syntax_args, arena, &env)) {
    return Value::error(*err);
  }
  return eval_node(*lv->body, arena, registry, lambda_body_context(lv, env, ctx));
}

}  // namespace

Value invoke_lambda_values(const LambdaValue* lv, std::uint32_t arity, const Value* args, Arena& arena,
                           const FunctionRegistry& registry, const EvalContext& ctx) {
  return invoke_lambda_values_with_ast(lv, arity, args, /*ast_args=*/nullptr, arena, registry, ctx);
}

Value invoke_lambda_values_with_ast(const LambdaValue* lv, std::uint32_t arity, const Value* args,
                                    const parser::AstNode* const* ast_args, Arena& arena,
                                    const FunctionRegistry& registry, const EvalContext& ctx) {
  if (lv == nullptr || (arity != 0U && args == nullptr)) {
    return Value::error(ErrorCode::Value);
  }
  // Lambda-depth cap fires before the arity check so a runaway
  // self-recursion (e.g. `LAMBDA(n, f(n+1))`) cannot keep extending the
  // call stack on its own dime. See `kMaxLambdaDepth` for rationale.
  EvalDepthGuard lambda_guard(ctx.lambda_depth_counter(), kMaxLambdaDepth);
  if (lambda_guard.exceeded()) {
    return Value::error(ErrorCode::Calc);
  }
  lv = at_arity(lv, arity, arena);
  if (lv == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return invoke_lambda_values_impl(lv, arity, args, ast_args, /*syntax_args=*/nullptr, arena, registry, ctx);
}

const parser::AstNode* array_literal_ast(const ArrayValue* arr, Arena& arena) {
  // `make_array_literal` requires rows / cols >= 1.
  if (arr == nullptr || arr->rows == 0U || arr->cols == 0U) {
    return nullptr;
  }
  const std::size_t total = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
  const parser::AstNode** children = arena.create_array<const parser::AstNode*>(total);
  if (children == nullptr) {
    return nullptr;
  }
  for (std::size_t i = 0; i < total; ++i) {
    parser::AstNode* lit = parser::make_literal(arena, arr->cells[i]);
    if (lit == nullptr) {
      return nullptr;
    }
    children[i] = lit;
  }
  return parser::make_array_literal(arena, arr->rows, arr->cols, children);
}

namespace {

// A lambda is callable by a helper when it accepts `call_arity` arguments
// (trailing `[optional]` params may stay unbound) and carries a body. A
// function value is callable at any arity; the built-in judges it when run.
const LambdaValue* check_callable(const LambdaValue* lv, std::uint32_t call_arity, Arena& arena, Value* out_err) {
  if (!lv->builtin.empty()) {
    const LambdaValue* expanded = at_arity(lv, call_arity, arena);
    if (expanded == nullptr) {
      *out_err = Value::error(ErrorCode::Num);
    }
    return expanded;
  }
  const std::uint32_t required = lv->param_count - lv->optional_count;
  if (call_arity < required || call_arity > lv->param_count) {
    *out_err = Value::error(ErrorCode::Value);
    return nullptr;
  }
  if (lv->body == nullptr) {
    *out_err = Value::error(ErrorCode::Name);
    return nullptr;
  }
  return lv;
}

}  // namespace

const LambdaValue* resolve_callable(const parser::AstNode& arg, std::uint32_t call_arity, Arena& arena,
                                    const FunctionRegistry& registry, const EvalContext& ctx, Value* out_err) {
  // A bare built-in name evaluates to a function value like any LAMBDA.
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return nullptr;
  }
  if (!v.is_lambda()) {
    *out_err = Value::error(ErrorCode::Value);
    return nullptr;
  }
  return check_callable(v.as_lambda(), call_arity, arena, out_err);
}

Value invoke_lambda(const LambdaValue* lv, std::uint32_t arity, const parser::AstNode* const* call_args, Arena& arena,
                    const FunctionRegistry& registry, const EvalContext& ctx) {
  if (lv == nullptr || (arity != 0U && call_args == nullptr)) {
    return Value::error(ErrorCode::Value);
  }
  // Lambda-depth cap fires before the arity check so a runaway
  // self-recursion (e.g. `LAMBDA(n, f(n+1))`) cannot keep extending the
  // call stack on its own dime. See `kMaxLambdaDepth` for rationale.
  EvalDepthGuard lambda_guard(ctx.lambda_depth_counter(), kMaxLambdaDepth);
  if (lambda_guard.exceeded()) {
    return Value::error(ErrorCode::Calc);
  }
  lv = at_arity(lv, arity, arena);
  if (lv == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::uint32_t required = lv->param_count - lv->optional_count;
  if (arity < required || arity > lv->param_count) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> args;
  std::vector<const parser::AstNode*> bound_asts;
  args.reserve(arity);
  bound_asts.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode* bound = nullptr;
    args.push_back(eval_binding_source(*call_args[i], arena, registry, ctx, &bound));
    bound_asts.push_back(bound);
  }
  return invoke_lambda_values_impl(lv, arity, args.empty() ? nullptr : args.data(),
                                   bound_asts.empty() ? nullptr : bound_asts.data(), call_args, arena, registry, ctx);
}

bool resolve_lambda_reference(const LambdaValue* lv, std::uint32_t arity, const parser::AstNode* const* call_args,
                              Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                              std::string_view* out_sheet, std::uint32_t* out_top_row, std::uint32_t* out_left_col,
                              std::uint32_t* out_bottom_row, std::uint32_t* out_right_col, ErrorCode* out_err) {
  if (lv == nullptr || (arity != 0U && call_args == nullptr)) {
    *out_err = ErrorCode::Value;
    return false;
  }
  EvalDepthGuard lambda_guard(ctx.lambda_depth_counter(), kMaxLambdaDepth);
  if (lambda_guard.exceeded()) {
    *out_err = ErrorCode::Calc;
    return false;
  }
  lv = at_arity(lv, arity, arena);
  if (lv == nullptr) {
    *out_err = ErrorCode::Num;
    return false;
  }
  std::vector<Value> args;
  std::vector<const parser::AstNode*> bound_asts;
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode* bound = nullptr;
    args.push_back(eval_binding_source(*call_args[i], arena, registry, ctx, &bound));
    bound_asts.push_back(bound);
  }
  NameEnv env;
  if (const std::optional<ErrorCode> err =
          bind_lambda_params(lv, arity, args.empty() ? nullptr : args.data(),
                             bound_asts.empty() ? nullptr : bound_asts.data(), call_args, arena, &env)) {
    *out_err = *err;
    return false;
  }
  return resolve_reference_rect(*lv->body, arena, registry, lambda_body_context(lv, env, ctx), out_sheet, out_top_row,
                                out_left_col, out_bottom_row, out_right_col, out_err);
}

}  // namespace eval
}  // namespace formulon
