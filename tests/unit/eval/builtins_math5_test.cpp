// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the PERMUT / PERMUTATIONA combinatorial builtins.
// Both are scalar-only, float-truncation inputs, with `#NUM!` on the
// documented domain violations.

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

// Parses `src` and evaluates it via the default function registry. Arenas
// are thread-local and reset per call.
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

TEST(BuiltinsMath5Registry, AllNamesRegistered) {
  for (const char* name : {"PERMUT", "PERMUTATIONA"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

// ---------------------------------------------------------------------------
// PERMUT
// ---------------------------------------------------------------------------

TEST(BuiltinsMath5Permut, Basic) {
  // 5 * 4 = 20.
  const Value v = EvalSource("=PERMUT(5, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsMath5Permut, KZeroIsOne) {
  const Value v = EvalSource("=PERMUT(5, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath5Permut, KEqualsNIsFactorial) {
  // 5! = 120.
  const Value v = EvalSource("=PERMUT(5, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 120.0);
}

TEST(BuiltinsMath5Permut, KGreaterThanNIsNum) {
  const Value v = EvalSource("=PERMUT(3, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath5Permut, NegativeNIsNum) {
  const Value v = EvalSource("=PERMUT(-1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// PERMUTATIONA
// ---------------------------------------------------------------------------

TEST(BuiltinsMath5PermutationA, Basic) {
  // 3^2 = 9.
  const Value v = EvalSource("=PERMUTATIONA(3, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

TEST(BuiltinsMath5PermutationA, ZeroZeroIsOne) {
  const Value v = EvalSource("=PERMUTATIONA(0, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath5PermutationA, TenPowerMatches) {
  // 2^10 = 1024.
  const Value v = EvalSource("=PERMUTATIONA(2, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1024.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
