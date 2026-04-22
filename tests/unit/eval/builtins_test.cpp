// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the registered built-in functions and the special-
// cased `IF` short-circuit. Each test parses a formula source, evaluates
// the AST through the default registry, and asserts the resulting Value.

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

// Parses `src` and evaluates it via the default function registry. The
// thread-local arenas keep text payloads readable for the immediately
// following EXPECT_*. Each call resets the arenas to avoid cross-test
// contamination.
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
// SUM
// ---------------------------------------------------------------------------

TEST(BuiltinsSum, SingleArgument) {
  const Value v = EvalSource("=SUM(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsSum, ThreeIntegers) {
  const Value v = EvalSource("=SUM(1,2,3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsSum, FractionalArguments) {
  const Value v = EvalSource("=SUM(1.5, 2.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsSum, BoolsCoerceToNumbers) {
  const Value v = EvalSource("=SUM(TRUE, FALSE, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsSum, NumericTextCoerces) {
  const Value v = EvalSource("=SUM(\"2\", 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsSum, NonNumericTextYieldsValue) {
  const Value v = EvalSource("=SUM(\"abc\", 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSum, ErrorPropagates) {
  const Value v = EvalSource("=SUM(1, #REF!, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsSum, LeftMostErrorWins) {
  const Value v = EvalSource("=SUM(1, #DIV/0!, #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSum, OverflowYieldsNum) {
  const Value v = EvalSource("=SUM(1e308, 1e308)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsSum, EmptyArgListIsArityViolation) {
  const Value v = EvalSource("=SUM()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// CONCAT / CONCATENATE
// ---------------------------------------------------------------------------

TEST(BuiltinsConcat, SingleString) {
  const Value v = EvalSource("=CONCAT(\"a\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a");
}

TEST(BuiltinsConcat, ThreeStrings) {
  const Value v = EvalSource("=CONCAT(\"a\",\"b\",\"c\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "abc");
}

TEST(BuiltinsConcat, NumbersStringify) {
  const Value v = EvalSource("=CONCAT(1,2,3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123");
}

TEST(BuiltinsConcat, BoolsStringify) {
  const Value v = EvalSource("=CONCAT(TRUE,\"-\",FALSE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "TRUE-FALSE");
}

TEST(BuiltinsConcat, ErrorPropagates) {
  const Value v = EvalSource("=CONCAT(\"x\", #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsConcatenate, AliasMatchesConcat) {
  const Value v = EvalSource("=CONCATENATE(\"a\",\"b\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(BuiltinsConcatenate, EmptyArgListIsArityViolation) {
  const Value v = EvalSource("=CONCATENATE()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// LEN
// ---------------------------------------------------------------------------

TEST(BuiltinsLen, EmptyString) {
  const Value v = EvalSource("=LEN(\"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsLen, AsciiString) {
  const Value v = EvalSource("=LEN(\"hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsLen, NumberCoercedToString) {
  const Value v = EvalSource("=LEN(123)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsLen, BoolCoercedToString) {
  const Value v = EvalSource("=LEN(TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsLen, BmpCharactersCountAsOneEach) {
  // "あいう" = 3 BMP codepoints, 3 UTF-16 units.
  const Value v = EvalSource("=LEN(\"\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsLen, SupplementaryPlaneCountsAsTwo) {
  // "🎉" U+1F389 -> surrogate pair -> 2 UTF-16 units.
  const Value v = EvalSource("=LEN(\"\xF0\x9F\x8E\x89\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsLen, ErrorPropagates) {
  const Value v = EvalSource("=LEN(#N/A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsLen, TooManyArgsIsArityViolation) {
  const Value v = EvalSource("=LEN(\"a\",\"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// IF (short-circuit, special-cased in the evaluator)
// ---------------------------------------------------------------------------

TEST(BuiltinsIf, TrueBranchSelected) {
  const Value v = EvalSource("=IF(TRUE, \"yes\", \"no\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "yes");
}

TEST(BuiltinsIf, FalseBranchSelected) {
  const Value v = EvalSource("=IF(FALSE, \"yes\", \"no\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "no");
}

TEST(BuiltinsIf, TruthyNumberSelectsTrueBranch) {
  const Value v = EvalSource("=IF(1, \"yes\", \"no\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "yes");
}

TEST(BuiltinsIf, ZeroSelectsFalseBranch) {
  const Value v = EvalSource("=IF(0, \"yes\", \"no\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "no");
}

TEST(BuiltinsIf, OmittedFalseBranchTrueCase) {
  const Value v = EvalSource("=IF(TRUE, \"yes\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "yes");
}

TEST(BuiltinsIf, OmittedFalseBranchFalseCaseReturnsBooleanFalse) {
  const Value v = EvalSource("=IF(FALSE, \"yes\")");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsIf, ConditionErrorPropagates) {
  const Value v = EvalSource("=IF(#REF!, 1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// The two short-circuit tests below are the load-bearing ones: if either
// fails, the IF branch is no longer being skipped at evaluation time.
TEST(BuiltinsIf, TrueBranchShortCircuitsDivByZero) {
  const Value v = EvalSource("=IF(TRUE, 1, 1/0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsIf, FalseBranchShortCircuitsDivByZero) {
  const Value v = EvalSource("=IF(FALSE, 1/0, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsIf, OneArgIsArityViolation) {
  const Value v = EvalSource("=IF(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsIf, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=IF()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Unknown function
// ---------------------------------------------------------------------------

TEST(BuiltinsDispatch, UnknownNameYieldsName) {
  const Value v = EvalSource("=FOOBAR(1,2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
