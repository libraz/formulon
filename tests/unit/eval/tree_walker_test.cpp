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
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
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

TEST(TreeWalkerNumeric, ZeroPowZeroIsNum) {
  // Excel reports #NUM! for 0^0 rather than the IEEE pow convention of 1.
  // This applies to both the `^` operator and the POWER() builtin via the
  // shared apply_pow helper.
  const Value v = EvalSource("=0^0");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
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

// ---------------------------------------------------------------------------
// Reference resolution via `EvalContext`
// ---------------------------------------------------------------------------

// Mirrors `EvalSource` but binds an `EvalContext` to `sheet` so local A1
// references in `src` resolve against the supplied sheet.
Value EvalInSheet(const Sheet& sheet, std::string_view src) {
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
  EvalContext ctx(sheet);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

TEST(TreeWalkerRefs, LiteralCellPlusOne) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(41.0));
  const Value v = EvalInSheet(sheet, "=A1+1");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(TreeWalkerRefs, EmptyCellPlusOne) {
  Sheet sheet("Sheet1");
  const Value v = EvalInSheet(sheet, "=A1+1");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(TreeWalkerRefs, TwoCellSum) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(10.0));
  sheet.set_cell_value(0, 1, Value::number(32.0));
  const Value v = EvalInSheet(sheet, "=A1+B1");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(TreeWalkerRefs, TextCellCoercedInArithmetic) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::text("5"));
  const Value v = EvalInSheet(sheet, "=A1+1");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(TreeWalkerRefs, TextCellConcatenation) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::text("hi"));
  const Value v = EvalInSheet(sheet, "=A1&\"!\"");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hi!");
}

TEST(TreeWalkerRefs, BoolCellInArithmetic) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::boolean(true));
  const Value v = EvalInSheet(sheet, "=A1+1");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(TreeWalkerRefs, ErrorCellPropagates) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::error(ErrorCode::Div0));
  const Value v = EvalInSheet(sheet, "=A1+1");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(TreeWalkerRefs, CrossSheetReturnsRef) {
  Sheet sheet("Sheet1");
  const Value v = EvalInSheet(sheet, "=Sheet2!A1");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(TreeWalkerRefs, WholeColumnReturnsValue) {
  Sheet sheet("Sheet1");
  const Value v = EvalInSheet(sheet, "=A:A");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TreeWalkerRefs, RefInIfBranch) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(7.0));
  const Value v = EvalInSheet(sheet, "=IF(TRUE,A1,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(TreeWalkerRefs, RefInSumCall) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(3.0));
  const Value v = EvalInSheet(sheet, "=SUM(A1,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

TEST(TreeWalkerRefs, NoContextStillNameError) {
  // The legacy two-arg overload must continue to surface #NAME? for any ref,
  // since its default context has no bound sheet.
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=A1+1", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const Value v = evaluate(*root, eval_arena, default_registry());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// Recursive formula-cell evaluation through the full pipeline
// ---------------------------------------------------------------------------

// Parses `src` and evaluates it through the full pipeline with a bound
// `EvalState`, so any reference to a formula cell in `sheet` recurses into
// the target cell's `formula_text`.
Value EvalInSheetWithState(const Sheet& sheet, std::string_view src) {
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
  EvalState state;
  const EvalContext ctx(sheet, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

TEST(TreeWalkerFormulaRefs, FormulaCellInArithmetic) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=10");
  const Value v = EvalInSheetWithState(sheet, "=A1+32");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(TreeWalkerFormulaRefs, ChainedFormulasInSum) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(5.0));
  sheet.set_cell_formula(1, 0, "=A1+1");
  sheet.set_cell_formula(2, 0, "=A2*2");
  const Value v = EvalInSheetWithState(sheet, "=A3+1");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 13.0);
}

TEST(TreeWalkerFormulaRefs, CycleInFormulaResolvesToRef) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=A1+1");
  const Value v = EvalInSheetWithState(sheet, "=A1");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(TreeWalkerFormulaRefs, FormulaErrorPropagatesThroughArith) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=1/0");
  const Value v = EvalInSheetWithState(sheet, "=A1+1");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(TreeWalkerFormulaRefs, FormulaTextResultConcat) {
  Sheet sheet("Sheet1");
  sheet.set_cell_formula(0, 0, "=\"hi\"");
  const Value v = EvalInSheetWithState(sheet, "=A1&\"!\"");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hi!");
}

// ---------------------------------------------------------------------------
// RangeOp argument expansion in aggregator calls
// ---------------------------------------------------------------------------

TEST(TreeWalkerRanges, SumThreeCellColumn) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(1, 0, Value::number(2.0));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalInSheet(sheet, "=SUM(A1:A3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(TreeWalkerRanges, SumTwoByTwo) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(0, 1, Value::number(2.0));
  sheet.set_cell_value(1, 0, Value::number(3.0));
  sheet.set_cell_value(1, 1, Value::number(4.0));
  const Value v = EvalInSheet(sheet, "=SUM(A1:B2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(TreeWalkerRanges, SumWithBlankInMiddle) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(5.0));
  // A2 is intentionally blank; SUM treats blank as 0.
  sheet.set_cell_value(2, 0, Value::number(7.0));
  const Value v = EvalInSheet(sheet, "=SUM(A1:A3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(TreeWalkerRanges, AverageOverThree) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(3.0));
  // A2 is blank; the range provenance filter drops blank cells before the
  // AVERAGE impl runs, so the divisor is 2 (matches Excel).
  sheet.set_cell_value(2, 0, Value::number(9.0));
  const Value v = EvalInSheet(sheet, "=AVERAGE(A1:A3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(TreeWalkerRanges, MinOverThree) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(5.0));
  sheet.set_cell_value(1, 0, Value::number(-1.0));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalInSheet(sheet, "=MIN(A1:A3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -1.0);
}

TEST(TreeWalkerRanges, MaxOverThree) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(5.0));
  sheet.set_cell_value(1, 0, Value::number(-1.0));
  sheet.set_cell_value(2, 0, Value::number(9.0));
  const Value v = EvalInSheet(sheet, "=MAX(A1:A3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 9.0);
}

TEST(TreeWalkerRanges, ProductOverThree) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(2.0));
  sheet.set_cell_value(1, 0, Value::number(3.0));
  sheet.set_cell_value(2, 0, Value::number(4.0));
  const Value v = EvalInSheet(sheet, "=PRODUCT(A1:A3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 24.0);
}

TEST(TreeWalkerRanges, MixedRangeAndScalar) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(1, 0, Value::number(2.0));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  sheet.set_cell_value(0, 1, Value::number(10.0));
  const Value v = EvalInSheet(sheet, "=SUM(A1:A3, B1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 16.0);
}

TEST(TreeWalkerRanges, ErrorCellPropagates) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalInSheet(sheet, "=SUM(A1:A3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(TreeWalkerRanges, TextCellInRangeIsSkipped) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(1, 0, Value::text("hello"));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  // The range provenance filter drops text cells silently for SUM / AVERAGE
  // / MIN / MAX / PRODUCT, matching Excel. A text *literal* passed directly
  // (=SUM(1,"hello",3)) still surfaces #VALUE!.
  const Value v = EvalInSheet(sheet, "=SUM(A1:A3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(TreeWalkerRanges, CrossSheetRangeReturnsRef) {
  Sheet sheet("Sheet1");
  const Value v = EvalInSheet(sheet, "=SUM(Sheet2!A1:A3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(TreeWalkerRanges, WholeColumnRangeReturnsValue) {
  Sheet sheet("Sheet1");
  const Value v = EvalInSheet(sheet, "=SUM(A:A)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TreeWalkerRanges, ReversedRangeSameResult) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(1, 0, Value::number(2.0));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  const Value forward = EvalInSheet(sheet, "=SUM(A1:A3)");
  const Value reversed = EvalInSheet(sheet, "=SUM(A3:A1)");
  ASSERT_TRUE(forward.is_number());
  ASSERT_TRUE(reversed.is_number());
  EXPECT_EQ(reversed.as_number(), forward.as_number());
}

TEST(TreeWalkerRanges, RangeInNonAggregatorStillValue) {
  Sheet sheet("Sheet1");
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(1, 0, Value::number(2.0));
  // LEN is not range-aware; its RangeOp argument still surfaces #VALUE!.
  const Value v = EvalInSheet(sheet, "=LEN(A1:A3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TreeWalkerRanges, CycleViaAggregator_ReturnsRef) {
  Sheet sheet("Sheet1");
  // A1 aggregates a range that includes itself; EvalState catches the
  // re-entry on A1 and surfaces #REF!.
  sheet.set_cell_formula(0, 0, "=SUM(A1:A3)");
  sheet.set_cell_value(1, 0, Value::number(2.0));
  sheet.set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalInSheetWithState(sheet, "=A1");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// _xlfn. / _xlfn._xlws. prefix stripping
// ---------------------------------------------------------------------------
//
// The xlsx format tags post-2007 functions with `_xlfn.` and modern
// worksheet-only ones with `_xlfn._xlws.` as a storage-side compatibility
// marker. Dispatch must treat those the same as the bare name.

TEST(TreeWalkerXlfnPrefix, EagerRegistryName) {
  // `CONCAT` is a regular (eager) entry in the function registry.
  const Value bare = EvalSource("=CONCAT(\"a\",\"b\",\"c\")");
  const Value tagged = EvalSource("=_xlfn.CONCAT(\"a\",\"b\",\"c\")");
  ASSERT_TRUE(bare.is_text());
  ASSERT_TRUE(tagged.is_text());
  EXPECT_EQ(bare.as_text(), tagged.as_text());
}

TEST(TreeWalkerXlfnPrefix, LazyDispatchName) {
  // `IFS` is handled through the lazy dispatch table.
  const Value bare = EvalSource("=IFS(TRUE,1,TRUE,2)");
  const Value tagged = EvalSource("=_xlfn.IFS(TRUE,1,TRUE,2)");
  ASSERT_TRUE(bare.is_number());
  ASSERT_TRUE(tagged.is_number());
  EXPECT_EQ(bare.as_number(), tagged.as_number());
}

TEST(TreeWalkerXlfnPrefix, XlwsPrefixedName) {
  // `_xlfn._xlws.` is the double-prefix form used for worksheet-only modern
  // functions; still an eager registry name (`SORT` chosen as an available
  // case-compat sample — fall through to #NAME? if unimplemented is fine).
  const Value tagged_ifs = EvalSource("=_xlfn._xlws.IFS(TRUE,1,TRUE,2)");
  const Value bare_ifs = EvalSource("=IFS(TRUE,1,TRUE,2)");
  ASSERT_TRUE(bare_ifs.is_number());
  ASSERT_TRUE(tagged_ifs.is_number());
  EXPECT_EQ(bare_ifs.as_number(), tagged_ifs.as_number());
}

TEST(TreeWalkerXlfnPrefix, UnknownAfterStripStillNameError) {
  // Stripping the prefix must not mask a genuinely unknown function.
  const Value v = EvalSource("=_xlfn.NOSUCHFN(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
