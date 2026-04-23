// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for the INFO(type_text) environment-string function. Formulon
// intentionally returns host-agnostic, build-time-stable strings (see the
// divergence note in `src/eval/builtins/info.cpp`), so there is no oracle
// corpus for INFO - these unit tests are the only contract.

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

TEST(BuiltinsInfoRegistry, InfoIsRegistered) {
  EXPECT_NE(default_registry().lookup("INFO"), nullptr);
}

TEST(BuiltinsInfoExtra, Directory) {
  const Value v = EvalSource("=INFO(\"directory\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "/");
}

TEST(BuiltinsInfoExtra, NumFileIsOne) {
  const Value v = EvalSource("=INFO(\"numfile\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsInfoExtra, OsVersion) {
  const Value v = EvalSource("=INFO(\"osversion\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Formulon");
}

TEST(BuiltinsInfoExtra, Recalc) {
  const Value v = EvalSource("=INFO(\"recalc\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Automatic");
}

TEST(BuiltinsInfoExtra, Release) {
  const Value v = EvalSource("=INFO(\"release\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Formulon 0.1");
}

TEST(BuiltinsInfoExtra, SystemIsPcdos) {
  // Formulon always returns "pcdos"; Mac Excel ja-JP would return "mac".
  // See the divergence note in `src/eval/builtins/info.cpp`.
  const Value v = EvalSource("=INFO(\"system\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "pcdos");
}

TEST(BuiltinsInfoExtra, CaseInsensitiveKey) {
  const Value v = EvalSource("=INFO(\"DIRECTORY\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "/");
}

TEST(BuiltinsInfoExtra, UnknownKeyIsValueError) {
  const Value v = EvalSource("=INFO(\"unknown\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsInfoExtra, NumericArgIsValueError) {
  // Numbers coerce to their shortest-text form (e.g. "5") which is not a
  // valid info_type.
  const Value v = EvalSource("=INFO(5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
