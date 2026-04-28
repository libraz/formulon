// Copyright 2026 libraz. Licensed under the MIT License.
//
// Evaluator tests for the `LAMBDA` form (`parser::NodeKind::Lambda`) and
// immediately-invoked lambda call (`parser::NodeKind::LambdaCall`).
//
// The Pratt parser does not yet recognise the `LAMBDA(...)` keyword form,
// so these tests build the AST directly via the `make_lambda` /
// `make_lambda_call` factories. That exercises exactly the evaluator
// surface the brief targets.
//
// Coverage:
//   * IIFE (immediately-invoked function expression).
//   * Multi-param IIFE.
//   * Zero-param IIFE.
//   * `LET`-bound lambda dispatched through a `NameRef` callee.
//   * Closure capture: lambda body sees an outer `LET` binding.
//   * Nested lambdas / currying.
//   * Arity mismatch -> #VALUE!.
//   * Argument-error propagation.
//   * Non-lambda callee -> #VALUE!.
//   * Top-level `Lambda` value surface contract: `evaluate()` projects it
//     onto `#CALC!` to match Mac Excel 365's cell renderer.

#include <string>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using parser::AstNode;
using parser::BinOp;
using parser::make_binary_op;
using parser::make_lambda;
using parser::make_lambda_call;
using parser::make_let_binding;
using parser::make_literal;
using parser::make_name_ref;

// Convenience: small wrapper that owns an arena and the ad-hoc AST built for
// each test, so tests stay focused on the assertions.
struct AstBuilder {
  Arena arena;

  AstNode* num(double v) { return make_literal(arena, Value::number(v)); }
  AstNode* nameref(std::string_view name) { return make_name_ref(arena, name); }
  AstNode* mul(AstNode* a, AstNode* b) { return make_binary_op(arena, BinOp::Mul, a, b); }
  AstNode* add(AstNode* a, AstNode* b) { return make_binary_op(arena, BinOp::Add, a, b); }
  AstNode* div(AstNode* a, AstNode* b) { return make_binary_op(arena, BinOp::Div, a, b); }

  AstNode* lambda1(std::string_view p0, AstNode* body) {
    std::vector<std::string_view> params{p0};
    return make_lambda(arena, params.data(), 1, body);
  }
  AstNode* lambda2(std::string_view p0, std::string_view p1, AstNode* body) {
    std::vector<std::string_view> params{p0, p1};
    return make_lambda(arena, params.data(), 2, body);
  }
  AstNode* lambda0(AstNode* body) { return make_lambda(arena, /*params=*/nullptr, 0, body); }

  AstNode* call(AstNode* callee, std::vector<AstNode*> args) {
    std::vector<const AstNode*> as;
    as.reserve(args.size());
    for (AstNode* a : args) {
      as.push_back(a);
    }
    return make_lambda_call(arena, callee, as.empty() ? nullptr : as.data(), static_cast<std::uint32_t>(as.size()));
  }
};

Value Eval(AstBuilder& b, AstNode* root) {
  return evaluate(*root, b.arena, default_registry(), EvalContext{});
}

// ---------------------------------------------------------------------------
// IIFE shapes
// ---------------------------------------------------------------------------

TEST(EvalLambda, SingleParamIIFE) {
  // =LAMBDA(x, x*2)(5) -> 10
  AstBuilder b;
  AstNode* body = b.mul(b.nameref("x"), b.num(2.0));
  AstNode* lam = b.lambda1("x", body);
  AstNode* call = b.call(lam, {b.num(5.0)});
  const Value v = Eval(b, call);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(EvalLambda, MultiParamIIFE) {
  // =LAMBDA(x, y, x+y)(3, 4) -> 7
  AstBuilder b;
  AstNode* body = b.add(b.nameref("x"), b.nameref("y"));
  AstNode* lam = b.lambda2("x", "y", body);
  AstNode* call = b.call(lam, {b.num(3.0), b.num(4.0)});
  const Value v = Eval(b, call);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(EvalLambda, ZeroParamIIFE) {
  // =LAMBDA(42)() -> 42 (the body is the constant 42; no params).
  AstBuilder b;
  AstNode* body = b.num(42.0);
  AstNode* lam = b.lambda0(body);
  AstNode* call = b.call(lam, {});
  const Value v = Eval(b, call);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 42.0);
}

// ---------------------------------------------------------------------------
// LET-bound dispatch
// ---------------------------------------------------------------------------

TEST(EvalLambda, LetBoundLambdaDispatch) {
  // =LET(f, LAMBDA(x, x*2), f(5) + f(10)) -> 30.
  AstBuilder b;
  AstNode* body = b.mul(b.nameref("x"), b.num(2.0));
  AstNode* lam = b.lambda1("x", body);

  // f(5) + f(10)
  AstNode* call_a = b.call(b.nameref("f"), {b.num(5.0)});
  AstNode* call_b = b.call(b.nameref("f"), {b.num(10.0)});
  AstNode* sum = b.add(call_a, call_b);

  std::vector<std::string_view> names{"f"};
  std::vector<const AstNode*> exprs{lam};
  AstNode* let = make_let_binding(b.arena, names.data(), exprs.data(), 1, sum);

  const Value v = Eval(b, let);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 30.0);
}

TEST(EvalLambda, ClosureCapturesLetBinding) {
  // =LET(y, 100, LAMBDA(x, x+y))(5) -> 105.
  AstBuilder b;
  AstNode* body = b.add(b.nameref("x"), b.nameref("y"));
  AstNode* lam = b.lambda1("x", body);

  std::vector<std::string_view> names{"y"};
  std::vector<const AstNode*> exprs{b.num(100.0)};
  AstNode* let = make_let_binding(b.arena, names.data(), exprs.data(), 1, lam);

  // Apply the LET-result lambda to 5 at the outermost level.
  AstNode* call = b.call(let, {b.num(5.0)});
  const Value v = Eval(b, call);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 105.0);
}

TEST(EvalLambda, NestedLambdaCurrying) {
  // =LAMBDA(x, LAMBDA(y, x+y))(3)(4) -> 7.
  AstBuilder b;
  AstNode* inner_body = b.add(b.nameref("x"), b.nameref("y"));
  AstNode* inner = b.lambda1("y", inner_body);
  AstNode* outer = b.lambda1("x", inner);
  AstNode* call_outer = b.call(outer, {b.num(3.0)});
  AstNode* call_inner = b.call(call_outer, {b.num(4.0)});
  const Value v = Eval(b, call_inner);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 7.0);
}

// ---------------------------------------------------------------------------
// Error paths
// ---------------------------------------------------------------------------

TEST(EvalLambda, ArityMismatchSurfacesValueError) {
  // =LAMBDA(x, y, x+y)(1) -> #VALUE!
  AstBuilder b;
  AstNode* body = b.add(b.nameref("x"), b.nameref("y"));
  AstNode* lam = b.lambda2("x", "y", body);
  AstNode* call = b.call(lam, {b.num(1.0)});
  const Value v = Eval(b, call);
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(EvalLambda, ArgumentErrorPropagates) {
  // =LAMBDA(x, x+1)(1/0) -> #DIV/0!
  AstBuilder b;
  AstNode* body = b.add(b.nameref("x"), b.num(1.0));
  AstNode* lam = b.lambda1("x", body);
  AstNode* div_by_zero = b.div(b.num(1.0), b.num(0.0));
  AstNode* call = b.call(lam, {div_by_zero});
  const Value v = Eval(b, call);
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(EvalLambda, NonLambdaCalleeSurfacesValueError) {
  // The parser does not currently emit `LambdaCall` nodes from source, but
  // the evaluator must still reject a non-lambda callee defensively when
  // future parser changes admit such forms. Build the AST directly.
  // Calls a literal number — i.e. `(42)(7)` — which must yield #VALUE!.
  AstBuilder b;
  AstNode* call = b.call(b.num(42.0), {b.num(7.0)});
  const Value v = Eval(b, call);
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Top-level surface contract
// ---------------------------------------------------------------------------

TEST(EvalLambda, BareLambdaAtTopLevelYieldsCalcError) {
  // Mac Excel 365 displays a non-IIFE LAMBDA expression in a cell as
  // `#CALC!` because the renderer cannot project a closure onto a scalar.
  // The internal evaluator still produces Lambda values for IIFE / LET
  // dispatch — only the public `evaluate()` boundary projects onto #CALC!.
  AstBuilder b;
  AstNode* body = b.add(b.nameref("x"), b.num(1.0));
  AstNode* lam = b.lambda1("x", body);
  const Value v = Eval(b, lam);
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(EvalLambda, DebugStringShowsParamCount) {
  // The internal Lambda value is observable via `debug_to_string` for
  // diagnostics. Build a Lambda value through the IIFE-less path that
  // does NOT go through `evaluate()` (which would project onto #CALC!),
  // by using `eval_node` directly via a closure-capture round-trip.
  // Simpler: construct a LambdaValue manually and verify the renderer.
  Arena arena;
  auto* lv = arena.create<LambdaValue>();
  ASSERT_NE(lv, nullptr);
  lv->params = nullptr;
  lv->param_count = 3;
  lv->body = nullptr;  // Not dereferenced by debug_to_string.
  lv->captured_env = nullptr;
  const Value v = Value::lambda(lv);
  EXPECT_EQ(v.debug_to_string(), "Lambda(3 params)");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
