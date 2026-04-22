// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the tree-walk evaluator. Tests parse a formula source and
// evaluate the AST end-to-end, except where the tested NodeKind is not easy
// to emit through the parser (in which case the AST is built directly via
// the `make_*` factories).

#include "eval/tree_walker.h"

#include <cmath>
#include <limits>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Convenience: parse `src` (with `=` allowed but optional), evaluate the AST
// in a fresh arena, and return the resulting Value. The arena is a static
// thread-local so the returned text payloads remain readable for the
// EXPECT_EQ that immediately follows. Each call clears the arena first to
// avoid cross-test contamination.
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
// Literals
// ---------------------------------------------------------------------------

TEST(TreeWalkerLiterals, NumberLiteral) {
  const Value v = EvalSource("=42");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(TreeWalkerLiterals, BoolLiteral) {
  const Value v = EvalSource("=TRUE");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(TreeWalkerLiterals, TextLiteral) {
  const Value v = EvalSource("=\"hello\"");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello");
}

TEST(TreeWalkerLiterals, ErrorLiteral) {
  const Value v = EvalSource("=#DIV/0!");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(TreeWalkerLiterals, BlankFromFactory) {
  Arena ast_arena;
  Arena out_arena;
  parser::AstNode* node = parser::make_literal(ast_arena, Value::blank());
  ASSERT_NE(node, nullptr);
  const Value v = evaluate(*node, out_arena);
  EXPECT_TRUE(v.is_blank());
}

// ---------------------------------------------------------------------------
// Unary operators
// ---------------------------------------------------------------------------

TEST(TreeWalkerUnary, UnaryPlus) {
  const Value v = EvalSource("=+5");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(TreeWalkerUnary, UnaryMinus) {
  const Value v = EvalSource("=-5");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -5.0);
}

TEST(TreeWalkerUnary, Percent) {
  const Value v = EvalSource("=50%");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.5);
}

TEST(TreeWalkerUnary, MinusOnBoolCoercesToOne) {
  const Value v = EvalSource("=-TRUE");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -1.0);
}

TEST(TreeWalkerUnary, MinusOnNumericText) {
  const Value v = EvalSource("=-\"10\"");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -10.0);
}

TEST(TreeWalkerUnary, OperandErrorPropagates) {
  const Value v = EvalSource("=-#DIV/0!");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Arithmetic
// ---------------------------------------------------------------------------

TEST(TreeWalkerArith, AddIntegers) {
  const Value v = EvalSource("=1+2");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(TreeWalkerArith, SubMulDivPow) {
  EXPECT_EQ(EvalSource("=7-3").as_number(), 4.0);
  EXPECT_EQ(EvalSource("=2*3").as_number(), 6.0);
  EXPECT_EQ(EvalSource("=10/4").as_number(), 2.5);
  EXPECT_EQ(EvalSource("=2^10").as_number(), 1024.0);
}

TEST(TreeWalkerArith, PrecedenceMulOverAdd) {
  const Value v = EvalSource("=1+2*3");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(TreeWalkerArith, ParensOverridePrecedence) {
  const Value v = EvalSource("=(1+2)*3");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

// ---------------------------------------------------------------------------
// Coercion in arithmetic
// ---------------------------------------------------------------------------

TEST(TreeWalkerCoerce, NumericTextLhs) {
  const Value v = EvalSource("=\"2\"+3");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(TreeWalkerCoerce, BoolPlusNumber) {
  EXPECT_EQ(EvalSource("=TRUE+1").as_number(), 2.0);
  EXPECT_EQ(EvalSource("=FALSE+1").as_number(), 1.0);
}

TEST(TreeWalkerCoerce, EmptyStringIsZeroInArith) {
  const Value v = EvalSource("=\"\"+1");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(TreeWalkerCoerce, NonNumericTextIsValueError) {
  const Value v = EvalSource("=\"abc\"+1");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Concatenation
// ---------------------------------------------------------------------------

TEST(TreeWalkerConcat, TextText) {
  const Value v = EvalSource("=\"a\"&\"b\"");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(TreeWalkerConcat, NumberNumberProducesIntegerForm) {
  // Critical: the integer-fast-path in format_double must give "12", never
  // "1.0000002.000000".
  const Value v = EvalSource("=1&2");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "12");
}

TEST(TreeWalkerConcat, FractionalNumberWithText) {
  const Value v = EvalSource("=1.5&\"x\"");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1.5x");
}

TEST(TreeWalkerConcat, BoolWithText) {
  const Value v = EvalSource("=TRUE&\" yes\"");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "TRUE yes");
}

TEST(TreeWalkerConcat, BlankLeftActsAsEmpty) {
  Arena ast_arena;
  Arena out_arena;
  parser::AstNode* lhs = parser::make_literal(ast_arena, Value::blank());
  parser::AstNode* rhs = parser::make_literal(ast_arena, Value::text("x"));
  parser::AstNode* node = parser::make_binary_op(ast_arena, parser::BinOp::Concat, lhs, rhs);
  ASSERT_NE(node, nullptr);
  const Value v = evaluate(*node, out_arena);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "x");
}

// ---------------------------------------------------------------------------
// Numeric comparison
// ---------------------------------------------------------------------------

TEST(TreeWalkerCompare, NumericLess) {
  const Value v = EvalSource("=1<2");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(TreeWalkerCompare, NumericEq) {
  const Value v = EvalSource("=1=1");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(TreeWalkerCompare, NumericNotEq) {
  const Value v = EvalSource("=1<>2");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(TreeWalkerCompare, NumericGreaterEqual) {
  const Value v = EvalSource("=2>=2");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// Cross-type comparison
// ---------------------------------------------------------------------------

TEST(TreeWalkerCompare, NumberLessThanText) {
  const Value v = EvalSource("=1<\"a\"");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(TreeWalkerCompare, TextLessThanBool) {
  const Value v = EvalSource("=\"z\"<FALSE");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(TreeWalkerCompare, TextEqualityIsCaseInsensitive) {
  const Value v = EvalSource("=\"abc\"=\"ABC\"");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

// ---------------------------------------------------------------------------
// Error propagation
// ---------------------------------------------------------------------------

TEST(TreeWalkerErrors, DivisionByZero) {
  const Value v = EvalSource("=1/0");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(TreeWalkerErrors, ZeroDivZeroIsDiv0) {
  const Value v = EvalSource("=0/0");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(TreeWalkerErrors, LeftMostWinsLhsError) {
  // Lhs error is returned even though rhs is also an error.
  const Value v = EvalSource("=#REF!+1");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(TreeWalkerErrors, RhsErrorPropagatesWhenLhsClean) {
  const Value v = EvalSource("=1+#REF!");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(TreeWalkerErrors, BothErrorsLeftMostWins) {
  const Value v = EvalSource("=#DIV/0!+#REF!");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(TreeWalkerErrors, ConcatPropagatesError) {
  const Value v = EvalSource("=#NAME?&\"x\"");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// Numeric overflow / domain errors
// ---------------------------------------------------------------------------

TEST(TreeWalkerNumeric, OverflowToNum) {
  const Value v = EvalSource("=1E308*10");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(TreeWalkerNumeric, NegativeBaseFractionalExponent) {
  // (-1)^0.5 -> NaN -> #NUM!
  const Value v = EvalSource("=(-1)^0.5");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(TreeWalkerNumeric, ZeroPowZeroIsOne) {
  // Excel matches IEEE pow(0, 0) == 1.
  const Value v = EvalSource("=0^0");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// Error-placeholder pathway
// ---------------------------------------------------------------------------

TEST(TreeWalkerPlaceholder, MalformedFormulaYieldsError) {
  // `=1++` triggers panic-mode recovery; the resulting AST contains an
  // ErrorPlaceholder that the evaluator must surface as #NAME?.
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=1++", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  // We do not assert on errors().empty(): the parser is expected to record
  // diagnostics here. We only care that evaluation surfaces an error Value.
  const Value v = evaluate(*root, eval_arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// Unsupported NodeKinds return the correct sentinel
// ---------------------------------------------------------------------------

TEST(TreeWalkerUnsupported, RefReturnsName) {
  const Value v = EvalSource("=A1");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(TreeWalkerUnsupported, UnknownCallReturnsName) {
  const Value v = EvalSource("=NOT_A_REAL_FUNCTION(1,2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(TreeWalkerUnsupported, ArrayLiteralReturnsValue) {
  const Value v = EvalSource("={1,2}");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TreeWalkerUnsupported, RangeOpReturnsValue) {
  Arena ast_arena;
  Arena out_arena;
  parser::Parser p("=A1:B2", ast_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const Value v = evaluate(*root, out_arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Implicit intersection
// ---------------------------------------------------------------------------

TEST(TreeWalkerIntersection, AtIsIdentityForScalar) {
  const Value v = EvalSource("=@1");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
