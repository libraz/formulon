//
// Recursion-depth caps for the tree-walk evaluator. The parser already
// caps formula nesting at 128 layers, but two orthogonal vectors bypass
// that bound:
//
//   * `eval_node` is re-entered through `EvalContext::resolve_ref` when a
//     referenced cell is itself a formula. A linear chain of references
//     (`A1=A2, A2=A3, ..., A1000=1`) is not a cycle, so the evaluator's
//     cycle detector lets it through; only `kMaxEvalDepth` bounds the
//     resulting native-stack consumption.
//
//   * `invoke_lambda` is re-entered every time a user-defined LAMBDA
//     applies itself recursively. The body AST stays small, so the eval
//     cap does not fire; `kMaxLambdaDepth` is the relevant guard.
//
// These tests pin both caps. The eval-depth test builds the AST directly
// (the parser would reject such a deeply nested formula), the lambda
// tests go through the parser to mirror real workloads.

#include <cstdint>
#include <string>
#include <vector>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using parser::AstNode;
using parser::make_literal;
using parser::make_unary_op;
using parser::UnaryOp;

// Builds a chain of `depth` nested unary `+` operators wrapping a single
// literal `1`. The evaluator descends into the operand on each layer, so
// the resulting AST exercises `eval_node` exactly `depth + 1` times in a
// strict ladder.
AstNode* BuildUnaryLadder(Arena& arena, std::uint32_t depth) {
  AstNode* node = make_literal(arena, Value::number(1.0));
  for (std::uint32_t i = 0; i < depth; ++i) {
    node = make_unary_op(arena, UnaryOp::Plus, node);
  }
  return node;
}

// 600 layers comfortably exceeds the 512 cap; below the cap the AST
// evaluates back to the literal value untouched.
TEST(EvalRecursionDepth, EvalNodeRejectsDeeplyNestedFormula) {
  Arena arena;
  AstNode* root = BuildUnaryLadder(arena, /*depth=*/600);
  ASSERT_NE(root, nullptr);

  const Value v = evaluate(*root, arena, default_registry(), test::mac_context());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(EvalRecursionDepth, EvalNodeAcceptsDepthBelowCap) {
  Arena arena;
  // 200 layers is well under the 512 cap: the result must be the literal.
  AstNode* root = BuildUnaryLadder(arena, /*depth=*/200);
  ASSERT_NE(root, nullptr);

  const Value v = evaluate(*root, arena, default_registry(), test::mac_context());
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

// Source-driven helper. Mirrors the `ParseAndEval` shape used by the
// neighbouring lambda tests: the parse and eval arenas are kept separate
// so the eval-side allocations cannot accidentally outlive the AST. The
// formula source MUST start with `=` (the parser's expected entry shape).
struct ParseAndEval {
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p;
  AstNode* root = nullptr;

  explicit ParseAndEval(std::string_view src) : p(src, parse_arena) { root = p.parse(); }

  Value run() {
    if (root == nullptr) {
      return Value::error(ErrorCode::Name);
    }
    return evaluate(*root, eval_arena, default_registry(), test::mac_context());
  }
};

// Excel LAMBDA closures capture their environment at definition time, so
// `LET(f, LAMBDA(n, f(n-1)), f(0))` does NOT see `f` inside its body —
// the binding is established AFTER the LAMBDA expression is captured.
// The idiomatic workaround is the Y-combinator pattern: have the lambda
// take itself as a parameter (`self`), and the caller passes the lambda
// to itself. `self(self, n-1)` then bottoms out at `invoke_lambda` once
// per recursion, so a deep `n` reliably grows the lambda-depth counter.
//
// `g(g, 300)` recurses 301 times before terminating. The cap is 256, so
// the inner-most invocations short-circuit to `#CALC!`, and the IF
// cascade propagates that error all the way back to the caller.
TEST(EvalRecursionDepth, LambdaSelfRecursionRejectedAtDepthCap) {
  ParseAndEval pe("=LET(g, LAMBDA(self, n, IF(n>0, self(self, n-1), 0)), g(g, 300))");
  ASSERT_TRUE(pe.p.errors().empty()) << "unexpected parse errors";
  const Value v = pe.run();
  ASSERT_TRUE(v.is_error()) << "expected #CALC!, got value of kind " << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

// `g(g, 0)` would never terminate (the body always recurses) without a
// cap; the lambda guard catches it and surfaces `#CALC!`.
TEST(EvalRecursionDepth, LambdaInfiniteRecursionRejected) {
  ParseAndEval pe("=LET(g, LAMBDA(self, n, self(self, n+1)), g(g, 0))");
  ASSERT_TRUE(pe.p.errors().empty()) << "unexpected parse errors";
  const Value v = pe.run();
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

// `g(g, 100)` recurses 101 times and is well below the 256 cap; it must
// terminate with the recursion's base case (0).
TEST(EvalRecursionDepth, BoundedLambdaCompletes) {
  ParseAndEval pe("=LET(g, LAMBDA(self, n, IF(n>0, self(self, n-1), 0)), g(g, 100))");
  ASSERT_TRUE(pe.p.errors().empty()) << "unexpected parse errors";
  const Value v = pe.run();
  ASSERT_TRUE(v.is_number()) << "expected number(0), got kind " << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 0.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
