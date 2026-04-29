// Copyright 2026 libraz. Licensed under the MIT License.
//
// Evaluator tests for the `LAMBDA` form (`parser::NodeKind::Lambda`) and
// immediately-invoked lambda call (`parser::NodeKind::LambdaCall`).
//
// Two layers of coverage:
//
//   * AST-driven (lower layer): tests that build the AST directly via the
//     `make_lambda` / `make_lambda_call` factories. These pin down the
//     evaluator surface independently of the parser front-end.
//
//   * Parser-driven (upper layer): tests that go through `Parser::Parse` so
//     that end-user formula source like `=LAMBDA(x, x*2)(5)` exercises the
//     full pipeline. These also cover the name-bound dispatch path where a
//     `LET`-bound LAMBDA is invoked through an ordinary `f(x)` Call AST
//     node (no `LambdaCall` node ever materialises).
//
// Coverage (AST layer):
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
//
// Coverage (parser layer):
//   * IIFE / multi-param / zero-param / LET-dispatch / closure capture /
//     currying — all expressed as formula source.
//   * Arity-mismatch and argument-error paths.
//   * Parser-level rejection of empty LAMBDA, cell-ref-shaped param,
//     duplicate param names.
//   * Calling a non-lambda LET binding via `name(args)` -> #VALUE!.

#include <string>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
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
    return make_lambda(arena, params.data(), 1, /*optional_count=*/0, body);
  }
  AstNode* lambda2(std::string_view p0, std::string_view p1, AstNode* body) {
    std::vector<std::string_view> params{p0, p1};
    return make_lambda(arena, params.data(), 2, /*optional_count=*/0, body);
  }
  AstNode* lambda0(AstNode* body) {
    return make_lambda(arena, /*params=*/nullptr, 0, /*optional_count=*/0, body);
  }

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
  lv->optional_count = 0;
  lv->body = nullptr;  // Not dereferenced by debug_to_string.
  lv->captured_env = nullptr;
  const Value v = Value::lambda(lv);
  EXPECT_EQ(v.debug_to_string(), "Lambda(3 params)");
}

// ---------------------------------------------------------------------------
// Parser-driven tests
// ---------------------------------------------------------------------------
//
// These build the AST through `Parser::Parse` so that end-user formula
// source exercises the full LAMBDA / IIFE pipeline (parameter parsing,
// postfix `(` for IIFE / curry, name-bound dispatch via NameEnv).

// Owns the parse and eval arenas plus the parser so caller assertions can
// inspect the produced diagnostics. Re-creating the parser per test is
// cheap and keeps each case fully isolated.
struct ParseAndEval {
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p;
  parser::AstNode* root = nullptr;

  explicit ParseAndEval(std::string_view src) : p(src, parse_arena) { root = p.parse(); }

  Value run() {
    if (root == nullptr) {
      return Value::error(ErrorCode::Name);
    }
    return evaluate(*root, eval_arena, default_registry(), EvalContext{});
  }
};

// Convenience wrapper for the common "parse, expect no errors, evaluate"
// flow. Tests that assert parse diagnostics use `ParseAndEval` directly.
Value EvalSource(std::string_view src) {
  ParseAndEval pe(src);
  EXPECT_TRUE(pe.p.errors().empty()) << "unexpected parse errors for: " << src;
  return pe.run();
}

TEST(EvalLambda, ParserSingleParamIIFE) {
  // =LAMBDA(x, x*2)(5) -> 10.
  const Value v = EvalSource("=LAMBDA(x, x*2)(5)");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(EvalLambda, ParserMultiParamIIFE) {
  // =LAMBDA(x, y, x+y)(3, 4) -> 7.
  const Value v = EvalSource("=LAMBDA(x, y, x+y)(3, 4)");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(EvalLambda, ParserZeroParamIIFE) {
  // =LAMBDA(42)() -> 42 (single slot is body; zero parameters).
  const Value v = EvalSource("=LAMBDA(42)()");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(EvalLambda, ParserLetBoundLambdaDispatch) {
  // =LET(f, LAMBDA(x, x*2), f(5)+f(10)) -> 30.
  // The parser produces a `Call` node for `f(...)`; the evaluator's
  // name-bound dispatch path resolves `f` to the Lambda value and
  // invokes it as if the user had written an explicit IIFE.
  const Value v = EvalSource("=LET(f, LAMBDA(x, x*2), f(5)+f(10))");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 30.0);
}

TEST(EvalLambda, ParserClosureCapturesLetBinding) {
  // =LET(y, 100, LAMBDA(x, x+y)(5)) -> 105. The inner LAMBDA closes over
  // the LET-bound `y`; the IIFE applies it at the body site.
  const Value v = EvalSource("=LET(y, 100, LAMBDA(x, x+y)(5))");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 105.0);
}

TEST(EvalLambda, ParserNestedLambdaCurryingFromSource) {
  // =LAMBDA(x, LAMBDA(y, x+y))(3)(4) -> 7. The postfix `(` rule chains
  // the second IIFE onto the first invocation's result.
  const Value v = EvalSource("=LAMBDA(x, LAMBDA(y, x+y))(3)(4)");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(EvalLambda, ParserArityMismatchSurfacesValueError) {
  // =LAMBDA(x, y, x+y)(1) -> #VALUE!.
  const Value v = EvalSource("=LAMBDA(x, y, x+y)(1)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(EvalLambda, ParserOptionalParamOmittedTakesGuardedBranch) {
  // =LAMBDA(x, [y], IF(ISOMITTED(y), x, x+y))(5) -> 5.
  // The optional `y` is omitted; ISOMITTED detects the omitted slot and
  // the IF picks the unguarded `x` branch.
  const Value v = EvalSource("=LAMBDA(x, [y], IF(ISOMITTED(y), x, x+y))(5)");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(EvalLambda, ParserOptionalParamSuppliedTakesPresentBranch) {
  // =LAMBDA(x, [y], IF(ISOMITTED(y), x, x+y))(5,10) -> 15.
  const Value v = EvalSource("=LAMBDA(x, [y], IF(ISOMITTED(y), x, x+y))(5,10)");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 15.0);
}

TEST(EvalLambda, ParserOptionalParamMinArityStillEnforced) {
  // Required slots = `param_count - optional_count` = 1; calling with
  // zero args still surfaces #VALUE!.
  const Value v = EvalSource("=LAMBDA(x, [y], x)()");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(EvalLambda, ParserIsomittedOnLiteralReturnsFalse) {
  // ISOMITTED outside any LAMBDA call (or applied to a non-name-ref
  // arg) returns FALSE.
  const Value v = EvalSource("=ISOMITTED(5)");
  ASSERT_TRUE(v.is_boolean()) << v.debug_to_string();
  EXPECT_FALSE(v.as_boolean());
}

TEST(EvalLambda, ParserIsomittedOnPresentLambdaParamReturnsFalse) {
  // =LET(f, LAMBDA(x, ISOMITTED(x)), f(5)) -> FALSE. `x` is a regular
  // (non-optional) param and is bound, so the omitted flag is FALSE.
  const Value v = EvalSource("=LET(f, LAMBDA(x, ISOMITTED(x)), f(5))");
  ASSERT_TRUE(v.is_boolean()) << v.debug_to_string();
  EXPECT_FALSE(v.as_boolean());
}

TEST(EvalLambda, ParserEmptyLambdaIsParseError) {
  // =LAMBDA() -> the parser must surface the dedicated `LambdaEmpty`
  // diagnostic; the evaluator surfaces `#NAME?` from the
  // `ErrorPlaceholder` left in the AST. The diagnostic is the
  // load-bearing assertion here.
  ParseAndEval pe("=LAMBDA()");
  ASSERT_FALSE(pe.p.errors().empty());
  bool saw = false;
  for (const auto& e : pe.p.errors()) {
    if (e.code == parser::ParseErrorCode::LambdaEmpty) {
      saw = true;
      break;
    }
  }
  EXPECT_TRUE(saw) << "expected LambdaEmpty diagnostic";
}

TEST(EvalLambda, ParserCellRefShapedParamIsParseError) {
  // =LAMBDA(A1, x) -> the param slot `A1` is forbidden because it
  // collides with the A1 cell reference. The tokenizer routes the
  // shape into `CellRef` which `parse_lambda_call` recognises and
  // diagnoses with `LambdaInvalidParam`.
  ParseAndEval pe("=LAMBDA(A1, x)");
  ASSERT_FALSE(pe.p.errors().empty());
  bool saw = false;
  for (const auto& e : pe.p.errors()) {
    if (e.code == parser::ParseErrorCode::LambdaInvalidParam) {
      saw = true;
      break;
    }
  }
  EXPECT_TRUE(saw) << "expected LambdaInvalidParam diagnostic";
}

TEST(EvalLambda, ParserDuplicateParamIsParseError) {
  // =LAMBDA(x, x, body) -> the second `x` shadows the first; the
  // parser rejects with `LambdaDuplicateParam`.
  ParseAndEval pe("=LAMBDA(x, x, body)");
  ASSERT_FALSE(pe.p.errors().empty());
  bool saw = false;
  for (const auto& e : pe.p.errors()) {
    if (e.code == parser::ParseErrorCode::LambdaDuplicateParam) {
      saw = true;
      break;
    }
  }
  EXPECT_TRUE(saw) << "expected LambdaDuplicateParam diagnostic";
}

TEST(EvalLambda, ParserArgErrorPropagatesThroughNameDispatch) {
  // =LET(f, LAMBDA(x, x+1), f(1/0)) -> #DIV/0!. Argument-evaluation
  // errors must propagate through the name-bound dispatch path, the
  // same way they do through an explicit IIFE.
  const Value v = EvalSource("=LET(f, LAMBDA(x, x+1), f(1/0))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(EvalLambda, ParserCallingNonLambdaBindingIsValueError) {
  // =LET(n, 5, n(1)) -> #VALUE!. The bound value is a number; calling
  // a non-callable resolves to #VALUE! at the dispatch site (matches
  // Excel's behaviour for the analogous `(1)(2)` IIFE shape).
  const Value v = EvalSource("=LET(n, 5, n(1))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
