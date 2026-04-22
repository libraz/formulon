// Copyright 2026 libraz. Licensed under the MIT License.
//
// Happy-path golden tests for the Pratt parser. Each test parses a
// formula string and compares the S-expression dump of the resulting AST
// against an expected fixture. The dump format is the parser corpus
// contract (see `src/parser/ast_dump.h`).

#include "parser/parser.h"

#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

// Parses `src`, asserts a non-null root and zero errors, and returns the
// S-expression dump for comparison against an expected literal.
std::string ParseToSexpr(std::string_view src) {
  Arena a;
  Parser p(src, a);
  AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  EXPECT_TRUE(p.errors().empty()) << "unexpected errors for: " << src;
  if (root == nullptr) {
    return "<null>";
  }
  return dump_sexpr(*root);
}

// ---------------------------------------------------------------------------
// Atoms
// ---------------------------------------------------------------------------

TEST(ParserAtoms, IntegerLiteral) { EXPECT_EQ(ParseToSexpr("=42"), "(num 42)"); }

TEST(ParserAtoms, NegativeInteger) { EXPECT_EQ(ParseToSexpr("=-7"), "(unary - (num 7))"); }

TEST(ParserAtoms, DecimalLiteral) { EXPECT_EQ(ParseToSexpr("=3.14"), "(num 3.14)"); }

TEST(ParserAtoms, BoolTrueUpperCase) { EXPECT_EQ(ParseToSexpr("=TRUE"), "(bool true)"); }

TEST(ParserAtoms, BoolTrueLowerCase) { EXPECT_EQ(ParseToSexpr("=true"), "(bool true)"); }

TEST(ParserAtoms, BoolFalse) { EXPECT_EQ(ParseToSexpr("=FALSE"), "(bool false)"); }

TEST(ParserAtoms, ErrorLiteralDiv0) { EXPECT_EQ(ParseToSexpr("=#DIV/0!"), "(err-lit #DIV/0!)"); }

TEST(ParserAtoms, ErrorLiteralName) { EXPECT_EQ(ParseToSexpr("=#NAME?"), "(err-lit #NAME?)"); }

TEST(ParserAtoms, ErrorLiteralNA) { EXPECT_EQ(ParseToSexpr("=#N/A"), "(err-lit #N/A)"); }

TEST(ParserAtoms, CellRefA1) { EXPECT_EQ(ParseToSexpr("=A1"), "(ref A1)"); }

TEST(ParserAtoms, CellRefAbsoluteBoth) { EXPECT_EQ(ParseToSexpr("=$A$1"), "(ref $A$1)"); }

TEST(ParserAtoms, CellRefAbsoluteCol) { EXPECT_EQ(ParseToSexpr("=$A1"), "(ref $A1)"); }

TEST(ParserAtoms, CellRefAbsoluteRow) { EXPECT_EQ(ParseToSexpr("=A$1"), "(ref A$1)"); }

TEST(ParserAtoms, CellRefMultiLetter) { EXPECT_EQ(ParseToSexpr("=AA10"), "(ref AA10)"); }

TEST(ParserAtoms, CellRefXfd1048576) { EXPECT_EQ(ParseToSexpr("=XFD1048576"), "(ref XFD1048576)"); }

TEST(ParserAtoms, SheetQualifiedRef) { EXPECT_EQ(ParseToSexpr("=Sheet1!A1"), "(ref Sheet1!A1)"); }

TEST(ParserAtoms, SheetQualifiedRefAbsolute) {
  EXPECT_EQ(ParseToSexpr("=Sheet1!$A$1"), "(ref Sheet1!$A$1)");
}

TEST(ParserAtoms, QuotedSheetQualifiedRef) {
  EXPECT_EQ(ParseToSexpr("='Sheet 1'!A1"), "(ref 'Sheet 1'!A1)");
}

TEST(ParserAtoms, FullColumnRef) { EXPECT_EQ(ParseToSexpr("=A:A"), "(ref A:A)"); }

TEST(ParserAtoms, FullRowRef) { EXPECT_EQ(ParseToSexpr("=1:1"), "(ref 1:1)"); }

TEST(ParserAtoms, SheetQualifiedFullColumn) {
  EXPECT_EQ(ParseToSexpr("=Sheet1!A:A"), "(ref Sheet1!A:A)");
}

TEST(ParserAtoms, SheetQualifiedFullRow) {
  EXPECT_EQ(ParseToSexpr("=Sheet1!1:1"), "(ref Sheet1!1:1)");
}

TEST(ParserAtoms, NameRef) { EXPECT_EQ(ParseToSexpr("=foo"), "(name foo)"); }

TEST(ParserAtoms, NameRefMixedCase) { EXPECT_EQ(ParseToSexpr("=Foo_bar"), "(name Foo_bar)"); }

// ---------------------------------------------------------------------------
// String literals
// ---------------------------------------------------------------------------

TEST(ParserStrings, PlainString) { EXPECT_EQ(ParseToSexpr("=\"hello\""), "(text \"hello\")"); }

TEST(ParserStrings, EmptyString) { EXPECT_EQ(ParseToSexpr("=\"\""), "(text \"\")"); }

TEST(ParserStrings, EscapedDoubleQuotes) {
  // Excel doubles `"` inside string literals; the tokenizer resolves the
  // escape so the payload contains a single `"`. The dumper then escapes
  // `"` as `\"` for unambiguous goldens.
  EXPECT_EQ(ParseToSexpr("=\"he said \"\"hi\"\"\""), "(text \"he said \\\"hi\\\"\")");
}

TEST(ParserStrings, ConcatTwoStrings) {
  EXPECT_EQ(ParseToSexpr("=\"a\"&\"b\""), "(binary & (text \"a\") (text \"b\"))");
}

TEST(ParserStrings, IfWithStringBranches) {
  EXPECT_EQ(ParseToSexpr("=IF(TRUE,\"yes\",\"no\")"),
            "(call IF (bool true) (text \"yes\") (text \"no\"))");
}

TEST(ParserStrings, LenOfString) {
  EXPECT_EQ(ParseToSexpr("=LEN(\"foo\")"), "(call LEN (text \"foo\"))");
}

// ---------------------------------------------------------------------------
// Optional leading `=` is consumed exactly once
// ---------------------------------------------------------------------------

TEST(ParserPrefix, EqualsOptional) {
  EXPECT_EQ(ParseToSexpr("1+2"), "(binary + (num 1) (num 2))");
  EXPECT_EQ(ParseToSexpr("=1+2"), "(binary + (num 1) (num 2))");
}

TEST(ParserPrefix, InternalEqualsIsComparison) {
  // Only the very first `=` is the formula prefix; `A1=B1` parses as Eq.
  EXPECT_EQ(ParseToSexpr("=A1=B1"), "(binary = (ref A1) (ref B1))");
}

// ---------------------------------------------------------------------------
// Parens
// ---------------------------------------------------------------------------

TEST(ParserParens, ParensOverridePrecedence) {
  EXPECT_EQ(ParseToSexpr("=(1+2)*3"), "(binary * (binary + (num 1) (num 2)) (num 3))");
}

TEST(ParserParens, RedundantParens) {
  EXPECT_EQ(ParseToSexpr("=((A1))"), "(ref A1)");
}

// ---------------------------------------------------------------------------
// Calls
// ---------------------------------------------------------------------------

TEST(ParserCalls, ThreeArgSum) {
  EXPECT_EQ(ParseToSexpr("=SUM(1,2,3)"), "(call SUM (num 1) (num 2) (num 3))");
}

TEST(ParserCalls, ZeroArgNow) { EXPECT_EQ(ParseToSexpr("=NOW()"), "(call NOW)"); }

TEST(ParserCalls, NestedIfWithRange) {
  EXPECT_EQ(ParseToSexpr("=IF(A1>0,SUM(B1:B10),0)"),
            "(call IF (binary > (ref A1) (num 0)) (call SUM (range (ref B1) (ref B10))) (num 0))");
}

TEST(ParserCalls, FunctionNameCasePreserved) {
  EXPECT_EQ(ParseToSexpr("=Sum(1)"), "(call Sum (num 1))");
}

TEST(ParserCalls, OneArg) { EXPECT_EQ(ParseToSexpr("=ABS(-1)"), "(call ABS (unary - (num 1)))"); }

// ---------------------------------------------------------------------------
// CellRef-to-call disambiguation
//
// Function names like `LOG10` lex as CellRef (column "LOG", row 10) because
// the tokenizer is grammar-agnostic. The parser's token-stream postprocess
// rewrites such tokens to Ident when immediately followed by `(`, so they
// reach the function-call dispatch path.
// ---------------------------------------------------------------------------

TEST(ParserCellRefCall, Log10WithArgIsCall) {
  EXPECT_EQ(ParseToSexpr("=LOG10(100)"), "(call LOG10 (num 100))");
}

TEST(ParserCellRefCall, Log10WithoutParenStaysCellRef) {
  EXPECT_EQ(ParseToSexpr("=LOG10"), "(ref LOG10)");
}

TEST(ParserCellRefCall, Log10WithSpaceBeforeParenIsCall) {
  // Whitespace is filtered before the postprocess loop, so a space between
  // the name and `(` does not block the rewrite.
  EXPECT_EQ(ParseToSexpr("=LOG10 (100)"), "(call LOG10 (num 100))");
}

TEST(ParserCellRefCall, SheetQualifiedLog10StaysCellRef) {
  // Bang-guard preserves sheet-qualified references.
  EXPECT_EQ(ParseToSexpr("=Sheet1!LOG10"), "(ref Sheet1!LOG10)");
}

TEST(ParserCellRefCall, RangeOfLog10IsRangeOfRefs) {
  // Neither side is followed by `(`, so the rewrite does not fire and both
  // operands remain CellRefs.
  EXPECT_EQ(ParseToSexpr("=LOG10:LOG10"), "(range (ref LOG10) (ref LOG10))");
}

TEST(ParserCellRefCall, Log10InArithmeticStaysCellRef) {
  EXPECT_EQ(ParseToSexpr("=LOG10+1"), "(binary + (ref LOG10) (num 1))");
}

TEST(ParserCellRefCall, Log10InsideSumArgsStaysCellRef) {
  // Inside SUM's arg list, LOG10 is not followed by `(`, so it remains a ref.
  EXPECT_EQ(ParseToSexpr("=SUM(LOG10)"), "(call SUM (ref LOG10))");
}

TEST(ParserCellRefCall, Log10EmptyCall) {
  EXPECT_EQ(ParseToSexpr("=LOG10()"), "(call LOG10)");
}

// ---------------------------------------------------------------------------
// Range operator
// ---------------------------------------------------------------------------

TEST(ParserRange, BasicCellPairRange) {
  EXPECT_EQ(ParseToSexpr("=A1:B2"), "(range (ref A1) (ref B2))");
}

TEST(ParserRange, ChainedRangeIsLeftAssoc) {
  EXPECT_EQ(ParseToSexpr("=A1:B2:C3"), "(range (range (ref A1) (ref B2)) (ref C3))");
}

TEST(ParserRange, RangeInsideSum) {
  EXPECT_EQ(ParseToSexpr("=SUM(A1:A10)"), "(call SUM (range (ref A1) (ref A10)))");
}

// ---------------------------------------------------------------------------
// Unary operators
// ---------------------------------------------------------------------------

TEST(ParserUnary, UnaryMinusInteger) { EXPECT_EQ(ParseToSexpr("=-1"), "(unary - (num 1))"); }

TEST(ParserUnary, UnaryPlusInteger) { EXPECT_EQ(ParseToSexpr("=+1"), "(unary + (num 1))"); }

TEST(ParserUnary, UnaryMinusOnRef) { EXPECT_EQ(ParseToSexpr("=-A1"), "(unary - (ref A1))"); }

TEST(ParserUnary, DoubleNegation) {
  EXPECT_EQ(ParseToSexpr("=--1"), "(unary - (unary - (num 1)))");
}

TEST(ParserUnary, PostfixPercentOnRef) { EXPECT_EQ(ParseToSexpr("=A1%"), "(unary % (ref A1))"); }

TEST(ParserUnary, PercentBindsTighterThanStar) {
  // `2*A1%` should parse as `2*(A1%)` (postfix `%` is tighter than `*`).
  EXPECT_EQ(ParseToSexpr("=2*A1%"), "(binary * (num 2) (unary % (ref A1)))");
}

TEST(ParserUnary, UnaryPrefixTighterThanPercent) {
  // `-2%` -> `(-2)%`: prefix unary binds tighter than postfix `%`, so the
  // postfix wraps the already-negated value.
  EXPECT_EQ(ParseToSexpr("=-2%"), "(unary % (unary - (num 2)))");
}

// ---------------------------------------------------------------------------
// Binary precedence
// ---------------------------------------------------------------------------

TEST(ParserBinary, MulBindsTighterThanAdd) {
  EXPECT_EQ(ParseToSexpr("=1+2*3"), "(binary + (num 1) (binary * (num 2) (num 3)))");
}

TEST(ParserBinary, AddIsLeftAssoc) {
  EXPECT_EQ(ParseToSexpr("=1+2+3"), "(binary + (binary + (num 1) (num 2)) (num 3))");
}

TEST(ParserBinary, PowIsRightAssoc) {
  EXPECT_EQ(ParseToSexpr("=2^3^4"), "(binary ^ (num 2) (binary ^ (num 3) (num 4)))");
}

TEST(ParserBinary, ComparisonChainsLeftAssoc) {
  EXPECT_EQ(ParseToSexpr("=1<2=TRUE"),
            "(binary = (binary < (num 1) (num 2)) (bool true))");
}

TEST(ParserBinary, ConcatRefs) {
  EXPECT_EQ(ParseToSexpr("=A1&B1"), "(binary & (ref A1) (ref B1))");
}

TEST(ParserBinary, AllSixComparisons) {
  EXPECT_EQ(ParseToSexpr("=1=2"), "(binary = (num 1) (num 2))");
  EXPECT_EQ(ParseToSexpr("=1<>2"), "(binary <> (num 1) (num 2))");
  EXPECT_EQ(ParseToSexpr("=1<2"), "(binary < (num 1) (num 2))");
  EXPECT_EQ(ParseToSexpr("=1<=2"), "(binary <= (num 1) (num 2))");
  EXPECT_EQ(ParseToSexpr("=1>2"), "(binary > (num 1) (num 2))");
  EXPECT_EQ(ParseToSexpr("=1>=2"), "(binary >= (num 1) (num 2))");
}

TEST(ParserBinary, CompareTighterThanConcat) {
  // `&` (bp 20) binds tighter than `=` (bp 10): so `A1&B1=C1` parses as
  // `(A1&B1)=C1`.
  EXPECT_EQ(ParseToSexpr("=A1&B1=C1"),
            "(binary = (binary & (ref A1) (ref B1)) (ref C1))");
}

TEST(ParserBinary, RangeBindsTighterThanArith) {
  EXPECT_EQ(ParseToSexpr("=A1:B2+C1"), "(binary + (range (ref A1) (ref B2)) (ref C1))");
}

// ---------------------------------------------------------------------------
// Implicit intersection
// ---------------------------------------------------------------------------

TEST(ParserAt, AtPrefixOnRef) { EXPECT_EQ(ParseToSexpr("=@A1"), "(at (ref A1))"); }

TEST(ParserAt, AtConsumesRestOfExpression) {
  // `@` has the lowest precedence, so it wraps the whole RHS expression.
  EXPECT_EQ(ParseToSexpr("=@A1+B1"), "(at (binary + (ref A1) (ref B1)))");
}

TEST(ParserAt, AtBeforeCall) {
  EXPECT_EQ(ParseToSexpr("=@SUM(A1:A10)"), "(at (call SUM (range (ref A1) (ref A10))))");
}

// ---------------------------------------------------------------------------
// Array literals
// ---------------------------------------------------------------------------

TEST(ParserArray, TwoByTwo) {
  EXPECT_EQ(ParseToSexpr("={1,2;3,4}"), "(array 2 2 (num 1) (num 2) (num 3) (num 4))");
}

TEST(ParserArray, OneByThree) {
  EXPECT_EQ(ParseToSexpr("={1,2,3}"), "(array 1 3 (num 1) (num 2) (num 3))");
}

TEST(ParserArray, ThreeByOne) {
  EXPECT_EQ(ParseToSexpr("={1;2;3}"), "(array 3 1 (num 1) (num 2) (num 3))");
}

TEST(ParserArray, BoolElements) {
  EXPECT_EQ(ParseToSexpr("={TRUE,FALSE}"), "(array 1 2 (bool true) (bool false))");
}

TEST(ParserArray, ErrorElement) {
  EXPECT_EQ(ParseToSexpr("={#DIV/0!}"), "(array 1 1 (err-lit #DIV/0!))");
}

TEST(ParserArray, UnaryMinusInElement) {
  EXPECT_EQ(ParseToSexpr("={-1,2}"), "(array 1 2 (num -1) (num 2))");
}

// ---------------------------------------------------------------------------
// Composite formulas
// ---------------------------------------------------------------------------

TEST(ParserComposite, SumOverCount) {
  EXPECT_EQ(ParseToSexpr("=SUM(A1:A10)/COUNT(A1:A10)"),
            "(binary / (call SUM (range (ref A1) (ref A10))) (call COUNT (range (ref A1) (ref A10))))");
}

TEST(ParserComposite, IfWithSqrt) {
  EXPECT_EQ(ParseToSexpr("=IF(A1>=0, SQRT(A1), -SQRT(-A1))"),
            "(call IF (binary >= (ref A1) (num 0)) (call SQRT (ref A1)) "
            "(unary - (call SQRT (unary - (ref A1)))))");
}

TEST(ParserComposite, RoundExpression) {
  EXPECT_EQ(ParseToSexpr("=ROUND((A1+B1)*C1, 2)"),
            "(call ROUND (binary * (binary + (ref A1) (ref B1)) (ref C1)) (num 2))");
}

TEST(ParserComposite, NestedFunctions) {
  EXPECT_EQ(ParseToSexpr("=MAX(MIN(A1,B1),C1)"),
            "(call MAX (call MIN (ref A1) (ref B1)) (ref C1))");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
