// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for the Cube-category built-in stubs. All seven CUBE* functions
// return #NAME? (Formulon has no OLAP connection infrastructure) once
// argument evaluation completes. An error in any argument short-circuits
// to that error via the dispatcher's default propagate_errors behaviour.

#include <string_view>

#include "eval/function_registry.h"
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
// Registry pin
// ---------------------------------------------------------------------------

TEST(BuiltinsCubeRegistry, NamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("CUBEKPIMEMBER"), nullptr);
  EXPECT_NE(reg.lookup("CUBEMEMBER"), nullptr);
  EXPECT_NE(reg.lookup("CUBEMEMBERPROPERTY"), nullptr);
  EXPECT_NE(reg.lookup("CUBERANKEDMEMBER"), nullptr);
  EXPECT_NE(reg.lookup("CUBESET"), nullptr);
  EXPECT_NE(reg.lookup("CUBESETCOUNT"), nullptr);
  EXPECT_NE(reg.lookup("CUBEVALUE"), nullptr);
}

// ---------------------------------------------------------------------------
// Stub returns #NAME?
// ---------------------------------------------------------------------------

TEST(BuiltinsCube, CubeValueMinArityReturnsName) {
  const Value v = EvalSource("=CUBEVALUE(\"conn\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeValueVariadicReturnsName) {
  const Value v = EvalSource("=CUBEVALUE(\"conn\",\"a\",\"b\",\"c\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeMemberReturnsName) {
  const Value v = EvalSource("=CUBEMEMBER(\"conn\",\"[Sales].[All]\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeMemberWithCaptionReturnsName) {
  const Value v = EvalSource("=CUBEMEMBER(\"conn\",\"[Sales].[All]\",\"caption\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeMemberPropertyReturnsName) {
  const Value v = EvalSource("=CUBEMEMBERPROPERTY(\"conn\",\"member\",\"prop\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeRankedMemberReturnsName) {
  const Value v = EvalSource("=CUBERANKEDMEMBER(\"conn\",\"set\",1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeKpiMemberReturnsName) {
  const Value v = EvalSource("=CUBEKPIMEMBER(\"conn\",\"kpi\",1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeSetReturnsName) {
  const Value v = EvalSource("=CUBESET(\"conn\",\"set\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeSetFullArityReturnsName) {
  const Value v = EvalSource("=CUBESET(\"conn\",\"set\",\"caption\",1,\"sort\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsCube, CubeSetCountReturnsName) {
  const Value v = EvalSource("=CUBESETCOUNT(\"set\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// Error propagation (dispatcher short-circuit)
// ---------------------------------------------------------------------------

TEST(BuiltinsCube, CubeValuePropagatesArgumentError) {
  const Value v = EvalSource("=CUBEVALUE(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsCube, CubeMemberPropagatesArgumentError) {
  const Value v = EvalSource("=CUBEMEMBER(\"conn\",1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsCube, CubeSetCountPropagatesArgumentError) {
  const Value v = EvalSource("=CUBESETCOUNT(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Arity enforcement (min/max bounds)
// ---------------------------------------------------------------------------

TEST(BuiltinsCube, CubeMemberMinArityViolation) {
  const Value v = EvalSource("=CUBEMEMBER(\"conn\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCube, CubeSetCountArityViolation) {
  const Value v = EvalSource("=CUBESETCOUNT(\"a\",\"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
