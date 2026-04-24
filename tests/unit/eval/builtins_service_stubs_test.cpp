// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for the service-stub family: IMAGE, RTD, TRANSLATE,
// DETECTLANGUAGE, COPILOT. Each returns a deterministic Excel-visible
// error (see `src/eval/builtins/service_stubs.cpp` for the rationale);
// all five ride the eager dispatch path so an argument error short-
// circuits before the fixed return fires. The tests here pin:
//
//   * the nominal registry entries exist (registry look-up succeeds);
//   * each stub returns its documented fixed error on valid inputs;
//   * an error argument propagates instead of the stub's nominal error.

#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry.
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

TEST(BuiltinsServiceStubRegistry, NamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("IMAGE"), nullptr);
  EXPECT_NE(reg.lookup("RTD"), nullptr);
  EXPECT_NE(reg.lookup("TRANSLATE"), nullptr);
  EXPECT_NE(reg.lookup("DETECTLANGUAGE"), nullptr);
  EXPECT_NE(reg.lookup("COPILOT"), nullptr);
}

// ---------------------------------------------------------------------------
// IMAGE - #VALUE!
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubImage, ReturnsValue) {
  const Value v = EvalSource("=IMAGE(\"http://x\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsServiceStubImage, ArgumentErrorPropagates) {
  // 1/0 -> #DIV/0! wins over the nominal #VALUE! stub return.
  const Value v = EvalSource("=IMAGE(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// RTD - #N/A
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubRtd, ReturnsNA) {
  const Value v = EvalSource("=RTD(\"prog\", \"srv\", \"topic\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsServiceStubRtd, ArgumentErrorPropagates) {
  const Value v = EvalSource("=RTD(1/0, \"srv\", \"topic\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// TRANSLATE - #NAME?
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubTranslate, ReturnsName) {
  const Value v = EvalSource("=TRANSLATE(\"hello\", \"en\", \"ja\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsServiceStubTranslate, ArityTwoAlsoWorks) {
  // Two-arg form is accepted (min_arity = 2 for auto-detect source).
  const Value v = EvalSource("=TRANSLATE(\"hello\", \"ja\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsServiceStubTranslate, ArgumentErrorPropagates) {
  const Value v = EvalSource("=TRANSLATE(\"hello\", 1/0, \"ja\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// DETECTLANGUAGE - #NAME?
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubDetectLanguage, ReturnsName) {
  // Literal Japanese input still produces #NAME? because the stub body
  // never consults its argument.
  const Value v = EvalSource("=DETECTLANGUAGE(\"\xE3\x81\x93\xE3\x82\x93\xE3\x81\xAB\xE3\x81\xA1\xE3\x81\xAF\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsServiceStubDetectLanguage, ArgumentErrorPropagates) {
  const Value v = EvalSource("=DETECTLANGUAGE(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// COPILOT - #NAME?
// ---------------------------------------------------------------------------

TEST(BuiltinsServiceStubCopilot, ReturnsName) {
  const Value v = EvalSource("=COPILOT(\"summarize\", 1, 2, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsServiceStubCopilot, ArgumentErrorPropagates) {
  const Value v = EvalSource("=COPILOT(\"summarize\", 1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
