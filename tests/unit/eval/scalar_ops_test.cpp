// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the stateless scalar-operator primitives in
// `eval/scalar_ops.h`. The focus is `apply_unary`'s Excel-365 contract:
// unary `+` is a type-preserving identity (does NOT coerce), unary `-`
// and `%` coerce to number, and any error operand is propagated
// verbatim. The tests bypass the parser to call `apply_unary` directly,
// which is the same seam that `tree_walker.cpp` and
// `shape_ops_lazy.cpp`'s array-context broadcaster invoke.

#include "eval/scalar_ops.h"

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Unary `+` — Excel 365 identity contract
// ---------------------------------------------------------------------------

TEST(ApplyUnary, PlusOnNumberReturnsSameNumber) {
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::number(5.0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(ApplyUnary, PlusOnNegativeNumberPreservesSign) {
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::number(-3.5));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.5);
}

TEST(ApplyUnary, PlusOnTextReturnsSameTextWithoutCoercion) {
  // Excel 365: `=+"text"` evaluates to the text "text" (NOT #VALUE!).
  // This is the canary for the dead-code regression — if the early
  // return is ever removed, this test surfaces a #VALUE! instead.
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::text("text"));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "text");
}

TEST(ApplyUnary, PlusOnEmptyTextReturnsEmptyText) {
  // `=+""` -> "" (empty text), the most cited divergence between Excel's
  // unary `+` identity and a hypothetical numeric coercion.
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::text(""));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(ApplyUnary, PlusOnTrueReturnsTrueWithoutCoercion) {
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::boolean(true));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(ApplyUnary, PlusOnFalseReturnsFalseWithoutCoercion) {
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::boolean(false));
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(ApplyUnary, PlusOnBlankReturnsBlank) {
  // `=+A1` where A1 is blank stays blank; the identity does not coerce.
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::blank());
  EXPECT_TRUE(v.is_blank());
}

TEST(ApplyUnary, PlusOnRefErrorPropagatesError) {
  // Errors short-circuit before the identity branch, but the resulting
  // Value still carries the original code unchanged.
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::error(ErrorCode::Ref));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(ApplyUnary, PlusOnDivZeroErrorPropagatesError) {
  const Value v = apply_unary(parser::UnaryOp::Plus, Value::error(ErrorCode::Div0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Unary `-` — coerces to number; #VALUE! on non-numeric text
// ---------------------------------------------------------------------------

TEST(ApplyUnary, MinusOnNumberNegates) {
  const Value v = apply_unary(parser::UnaryOp::Minus, Value::number(7.5));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -7.5);
}

TEST(ApplyUnary, MinusOnNonNumericTextReturnsValueError) {
  const Value v = apply_unary(parser::UnaryOp::Minus, Value::text("abc"));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ApplyUnary, MinusOnNumericTextCoercesAndNegates) {
  const Value v = apply_unary(parser::UnaryOp::Minus, Value::text("12"));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -12.0);
}

// ---------------------------------------------------------------------------
// Unary `%` — coerces to number, divides by 100
// ---------------------------------------------------------------------------

TEST(ApplyUnary, PercentOnNumberDividesBy100) {
  const Value v = apply_unary(parser::UnaryOp::Percent, Value::number(50.0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.5);
}

TEST(ApplyUnary, PercentOnTrueCoercesToOne) {
  // TRUE coerces to 1; 1% -> 0.01.
  const Value v = apply_unary(parser::UnaryOp::Percent, Value::boolean(true));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.01);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
