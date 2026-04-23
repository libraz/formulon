// Copyright 2026 libraz. Licensed under the MIT License.
//
// Evaluator tests for the `LET` special form and bare `NameRef` lookup.
// Covers sequential binding, shadowing, error flow, nested LETs, case
// insensitivity, and the base-case `#NAME?` when a name is unbound.

#include <string_view>

#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// ---------------------------------------------------------------------------
// Scalar binding shapes
// ---------------------------------------------------------------------------

TEST(EvalLet, SingleBindingArithmetic) {
  const Value v = EvalSource("=LET(x, 10, x+5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 15.0);
}

TEST(EvalLet, SequentialBindingsUsePriorScope) {
  // y is defined in terms of x, which is already in scope.
  const Value v = EvalSource("=LET(x, 10, y, x*2, y+1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 21.0);
}

TEST(EvalLet, BodyReferencesFirstBinding) {
  const Value v = EvalSource("=LET(x, 7, x)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(EvalLet, NumericConstantBinding) {
  const Value v = EvalSource("=LET(x, 0, x)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(EvalLet, TextValueBinding) {
  const Value v = EvalSource("=LET(g, \"hello\", UPPER(g))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "HELLO");
}

TEST(EvalLet, BooleanValueBinding) {
  const Value v = EvalSource("=LET(b, TRUE, IF(b, 1, 0))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(EvalLet, FunctionCallOnBoundValue) {
  const Value v = EvalSource("=LET(x, 3, MAX(x, 5))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

// ---------------------------------------------------------------------------
// Nested LETs
// ---------------------------------------------------------------------------

TEST(EvalLet, NestedLetInBindingExpression) {
  // Inner LET produces 25; outer body adds 1.
  const Value v = EvalSource("=LET(x, LET(y, 5, y*y), x+1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 26.0);
}

TEST(EvalLet, NestedLetInBody) {
  const Value v = EvalSource("=LET(x, 2, LET(y, x+1, y*y))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

// ---------------------------------------------------------------------------
// Shadowing
// ---------------------------------------------------------------------------

TEST(EvalLet, LaterBindingShadowsEarlier) {
  // Second `x` binds (outer x) + 10 = 11. Body reads the shadowed value.
  const Value v = EvalSource("=LET(x, 1, x, x+10, x)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 11.0);
}

TEST(EvalLet, InnerLetDoesNotLeakOuterScope) {
  // Outer x is shadowed inside the inner LET's body but recovers after.
  const Value v = EvalSource("=LET(x, 1, LET(x, 99, x) + x)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 100.0);
}

// ---------------------------------------------------------------------------
// Case sensitivity
// ---------------------------------------------------------------------------

TEST(EvalLet, NameLookupIsCaseInsensitive) {
  const Value v = EvalSource("=LET(Foo, 3, FOO+1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(EvalLet, ShadowingIsCaseInsensitive) {
  // FOO and foo are the same identifier; the second binding shadows.
  const Value v = EvalSource("=LET(FOO, 1, foo, 9, Foo)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

// ---------------------------------------------------------------------------
// Error flow
// ---------------------------------------------------------------------------

TEST(EvalLet, BindingErrorIsCatchableInBody) {
  // LET does not short-circuit on a binding initialiser that evaluates to
  // an error -- the error becomes the binding's value and IFERROR in the
  // body can catch it.
  const Value v = EvalSource("=LET(x, 1/0, IFERROR(x, 99))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 99.0);
}

TEST(EvalLet, BindingErrorPropagatesWhenUncaught) {
  const Value v = EvalSource("=LET(x, 1/0, x+1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(EvalLet, ErrorInBodyPropagates) {
  const Value v = EvalSource("=LET(x, 5, x/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Name-resolution boundary
// ---------------------------------------------------------------------------

TEST(EvalLet, UnboundNameRefIsNameError) {
  // Bare identifier at the top level (no LET in scope, no workbook-level
  // name resolution wired) resolves to #NAME?.
  const Value v = EvalSource("=foo");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(EvalLet, NameOutOfLetScopeIsNameError) {
  // `y` is never bound; the inner reference fails.
  const Value v = EvalSource("=LET(x, 1, y)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(EvalLet, ForwardReferenceBeforeBindingIsNameError) {
  // The body of `x`'s initialiser runs BEFORE y is bound, so `y` is
  // unresolved and the binding value becomes #NAME?. Body then uses x
  // which carries that error forward.
  const Value v = EvalSource("=LET(x, y, y, 2, x)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
