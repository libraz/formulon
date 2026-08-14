//
// Unit tests for the stateless scalar-operator primitives in
// `eval/scalar_ops.h`. The focus is `apply_unary`'s Excel-365 contract:
// unary `+` is a type-preserving identity (does NOT coerce), unary `-`
// and `%` coerce to number, and any error operand is propagated
// verbatim. The tests bypass the parser to call `apply_unary` directly,
// which is the same seam that `tree_walker.cpp` and
// `shape_ops_lazy.cpp`'s array-context broadcaster invoke.

#include "eval/scalar_ops.h"

#include <limits>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

void ExpectAllComparisonOperators(const Value& lhs, const Value& rhs, int expected_cmp,
                                  bool expected_unordered = false) {
  bool unordered = false;
  EXPECT_EQ(compare_values(lhs, rhs, &unordered), expected_cmp);
  EXPECT_EQ(unordered, expected_unordered);

  const parser::BinOp operators[] = {
      parser::BinOp::Eq,   parser::BinOp::NotEq, parser::BinOp::Lt,
      parser::BinOp::LtEq, parser::BinOp::Gt,    parser::BinOp::GtEq,
  };
  const bool expected[] = {
      !expected_unordered && expected_cmp == 0, expected_unordered || expected_cmp != 0,
      !expected_unordered && expected_cmp < 0,  !expected_unordered && expected_cmp <= 0,
      !expected_unordered && expected_cmp > 0,  !expected_unordered && expected_cmp >= 0,
  };
  for (std::size_t i = 0; i < sizeof(operators) / sizeof(operators[0]); ++i) {
    const Value actual = apply_comparison(operators[i], lhs, rhs);
    ASSERT_TRUE(actual.is_boolean()) << "comparison operator index " << i;
    EXPECT_EQ(actual.as_boolean(), expected[i]) << "comparison operator index " << i;
  }
}

// ---------------------------------------------------------------------------
// Numeric comparison equality buckets and raw-order fallbacks
// ---------------------------------------------------------------------------

TEST(CompareValues, AdjacentMroundAndRoundValuesShareEqualityBucket) {
  // Mac Excel's MROUND and ROUND paths can produce adjacent doubles for the
  // same displayed value (7.1 versus 7.1000000000000005). Formula comparison
  // operators must agree on the 15-significant-digit equality bucket.
  ExpectAllComparisonOperators(Value::number(7.1), Value::number(7.1000000000000005), 0);
  ExpectAllComparisonOperators(Value::number(-7.1), Value::number(-7.1000000000000005), 0);
}

TEST(CompareValues, FifteenDigitBucketsOnlyRelaxEqualityWhenKeysMatch) {
  ExpectAllComparisonOperators(Value::number(1.23456789012344), Value::number(1.2345678901234402), 0);
  ExpectAllComparisonOperators(Value::number(1.23456789012344), Value::number(1.23456789012346), -1);
  ExpectAllComparisonOperators(Value::number(-1.23456789012344), Value::number(-1.23456789012346), 1);
}

TEST(CompareValues, HalfUpBoundaryAtOneQuadrillionUsesFifteenDigits) {
  const double base = 1'000'000'000'000'000.0;
  const double plus_four = 1'000'000'000'000'004.0;
  const double plus_five = 1'000'000'000'000'005.0;
  const double plus_ten = 1'000'000'000'000'010.0;
  ExpectAllComparisonOperators(Value::number(base), Value::number(plus_four), 0);
  ExpectAllComparisonOperators(Value::number(base), Value::number(plus_five), -1);
  ExpectAllComparisonOperators(Value::number(base), Value::number(plus_ten), -1);
  ExpectAllComparisonOperators(Value::number(plus_five), Value::number(plus_ten), 0);

  const double negative_base = -base;
  const double minus_four = -1'000'000'000'000'004.0;
  const double minus_five = -1'000'000'000'000'005.0;
  const double minus_ten = -1'000'000'000'000'010.0;
  ExpectAllComparisonOperators(Value::number(negative_base), Value::number(minus_four), 0);
  ExpectAllComparisonOperators(Value::number(negative_base), Value::number(minus_five), 1);
  ExpectAllComparisonOperators(Value::number(negative_base), Value::number(minus_ten), 1);
  ExpectAllComparisonOperators(Value::number(minus_five), Value::number(minus_ten), 0);
}

TEST(CompareValues, ZeroSubnormalInfinityAndSignMismatchRetainRawOrder) {
  const double subnormal = std::numeric_limits<double>::denorm_min();
  ExpectAllComparisonOperators(Value::number(-0.0), Value::number(0.0), 0);
  ExpectAllComparisonOperators(Value::number(0.0), Value::number(subnormal), -1);
  ExpectAllComparisonOperators(Value::number(subnormal), Value::number(2.0 * subnormal), -1);
  ExpectAllComparisonOperators(Value::number(-subnormal), Value::number(subnormal), -1);
  ExpectAllComparisonOperators(Value::number(-7.1), Value::number(7.1), -1);

  const double infinity = std::numeric_limits<double>::infinity();
  const double max_finite = std::numeric_limits<double>::max();
  ExpectAllComparisonOperators(Value::number(infinity), Value::number(infinity), 0);
  ExpectAllComparisonOperators(Value::number(max_finite), Value::number(infinity), -1);
  ExpectAllComparisonOperators(Value::number(-infinity), Value::number(-max_finite), -1);
}

TEST(CompareValues, NaNRemainsUnorderedForEveryComparisonOperator) {
  ExpectAllComparisonOperators(Value::number(std::numeric_limits<double>::quiet_NaN()), Value::number(1.0), 0, true);
}

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
