//
// Round-trip golden tests for `format_formula`. Each case formats a
// hand-built or parsed AST, re-parses the result, and asserts that the
// `dump_sexpr` output matches the input. The contract is structural
// equivalence — byte-stable formatting is explicitly NOT required.

#include "parser/ast_format.h"

#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace parser {
namespace {

// Parses `src` and returns the dump of the resulting AST. Asserts no parse
// errors so the test fails loudly when a round-trip produces bad text.
std::string ParseDump(std::string_view src, Arena& arena) {
  Parser p(src, arena);
  AstNode* root = p.parse();
  EXPECT_TRUE(p.errors().empty()) << "unexpected parse errors for: " << src;
  EXPECT_NE(root, nullptr);
  if (root == nullptr) {
    return "<null>";
  }
  return dump_sexpr(*root);
}

// Round-trips a parsed formula: parse, format, parse again, dump. Both
// dumps must match exactly.
void ExpectRoundTripsToSame(std::string_view src) {
  Arena a1;
  Parser p1(src, a1);
  AstNode* root1 = p1.parse();
  ASSERT_NE(root1, nullptr) << "first parse failed for: " << src;
  ASSERT_TRUE(p1.errors().empty()) << "first parse errored for: " << src;
  const std::string sexpr1 = dump_sexpr(*root1);
  const std::string formatted = format_formula(*root1);

  Arena a2;
  // Re-parse: the formatter never emits a leading `=`, so do not prepend one.
  const std::string sexpr2 = ParseDump(formatted, a2);
  EXPECT_EQ(sexpr2, sexpr1) << "round-trip diverged for: " << src << "\n  formatted = " << formatted;
}

const AstNode* MakeTooDeepAst(Arena& arena) {
  AstNode* node = make_literal(arena, Value::number(1));
  for (std::uint32_t depth = 1; depth <= kMaxFormulaAstDepth; ++depth) {
    node = make_unary_op(arena, UnaryOp::Plus, node);
  }
  return node;
}

TEST(AstFormat, RejectsAstDeeperThanSharedLimit) {
  Arena arena;
  const AstNode* root = MakeTooDeepAst(arena);
  ASSERT_NE(root, nullptr);
  EXPECT_FALSE(ast_depth_within_limit(*root, kMaxFormulaAstDepth));
  EXPECT_EQ(format_formula(*root), "#REF!");
}

// ---------------------------------------------------------------------------
// Literals
// ---------------------------------------------------------------------------

TEST(AstFormat, NumberInteger) {
  ExpectRoundTripsToSame("=42");
}
TEST(AstFormat, NumberDecimal) {
  ExpectRoundTripsToSame("=3.14");
}
TEST(AstFormat, NumberZero) {
  ExpectRoundTripsToSame("=0");
}

TEST(AstFormat, BoolTrue) {
  ExpectRoundTripsToSame("=TRUE");
}
TEST(AstFormat, BoolFalse) {
  ExpectRoundTripsToSame("=FALSE");
}

TEST(AstFormat, TextEmpty) {
  ExpectRoundTripsToSame("=\"\"");
}
TEST(AstFormat, TextSimple) {
  ExpectRoundTripsToSame("=\"hello\"");
}
TEST(AstFormat, TextEscapesEmbeddedQuote) {
  ExpectRoundTripsToSame("=\"he said \"\"hi\"\"\"");
}

TEST(AstFormat, OmittedCallArgumentsRoundTripAsOmittedSlots) {
  // Blank AST literals encode omitted arguments, not empty-text literals.
  // This form is rewritten during sheet-name and structural edits, so the
  // formatter must preserve the exact argument kind through a re-parse.
  ExpectRoundTripsToSame("=VLOOKUP(A1,B:C,2,)");
  ExpectRoundTripsToSame("=IF(,1,)");
}

TEST(AstFormat, ErrorLiteralDiv0) {
  ExpectRoundTripsToSame("=#DIV/0!");
}
TEST(AstFormat, ErrorLiteralRef) {
  ExpectRoundTripsToSame("=#REF!");
}
TEST(AstFormat, ErrorLiteralName) {
  ExpectRoundTripsToSame("=#NAME?");
}

// ---------------------------------------------------------------------------
// References
// ---------------------------------------------------------------------------

TEST(AstFormat, RefBareA1) {
  ExpectRoundTripsToSame("=A1");
}
TEST(AstFormat, RefAbsoluteBoth) {
  ExpectRoundTripsToSame("=$A$1");
}
TEST(AstFormat, RefAbsoluteCol) {
  ExpectRoundTripsToSame("=$A1");
}
TEST(AstFormat, RefAbsoluteRow) {
  ExpectRoundTripsToSame("=A$1");
}
TEST(AstFormat, RefSheetQualified) {
  ExpectRoundTripsToSame("=Sheet1!A1");
}
TEST(AstFormat, RefSheetQuoted) {
  ExpectRoundTripsToSame("='Sheet 1'!A1");
}
TEST(AstFormat, RefNumericSheetQuoted) {
  ExpectRoundTripsToSame("='2026'!A1");
}
TEST(AstFormat, RefFullColumn) {
  ExpectRoundTripsToSame("=A:A");
}
TEST(AstFormat, RefFullRow) {
  ExpectRoundTripsToSame("=1:1");
}
TEST(AstFormat, RefMultiColumnFullRange) {
  ExpectRoundTripsToSame("=A:C");
}
TEST(AstFormat, RefMultiRowFullRange) {
  ExpectRoundTripsToSame("=1:3");
}
TEST(AstFormat, RefAbsoluteFullColumn) {
  ExpectRoundTripsToSame("=$A:$A");
}
TEST(AstFormat, RefAbsoluteFullRow) {
  ExpectRoundTripsToSame("=$1:$1");
}
TEST(AstFormat, RefAbsoluteMultiColumnFullRange) {
  ExpectRoundTripsToSame("=$A:$C");
}
TEST(AstFormat, RefAbsoluteMultiRowFullRange) {
  ExpectRoundTripsToSame("=$1:$3");
}
TEST(AstFormat, ThreeDSingleCell) {
  ExpectRoundTripsToSame("=SUM(Sheet1:Sheet3!A1)");
}
TEST(AstFormat, ThreeDRangeTail) {
  ExpectRoundTripsToSame("=SUM(Sheet1:Sheet3!A1:B2)");
}
TEST(AstFormat, ThreeDRangeTailQuotedSpan) {
  ExpectRoundTripsToSame("=SUM('Data:S2'!A1:B2)");
}
TEST(AstFormat, ThreeDRangeTailAbsolute) {
  ExpectRoundTripsToSame("=SUM(Sheet1:Sheet3!$A$1:$B$2)");
}

TEST(AstFormat, NameRef) {
  ExpectRoundTripsToSame("=foo");
}

TEST(AstFormat, SpillRef) {
  ExpectRoundTripsToSame("=A1#");
}

TEST(AstFormat, ExternalRefRoundTripsViaAst) {
  // The parser does not consume `[1]Sheet1!A1` source, so build the AST
  // manually and verify format → re-parse → equivalent shape. Re-parsing
  // an external-ref formatter output requires a parser path that is out of
  // scope for this bundle; instead verify the *shape* of the output text
  // matches the canonical form so future external-ref parsing slots in.
  Arena a;
  Reference cell;
  cell.col = 0;
  cell.row = 0;
  AstNode* n = make_external_ref(a, 1, "Sheet1", cell);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_formula(*n), "[1]Sheet1!A1");
}

TEST(AstFormat, ExternalRefQuotesSheetWithSpace) {
  Arena a;
  Reference cell;
  cell.col = 0;
  cell.row = 0;
  AstNode* n = make_external_ref(a, 2, "My Book", cell);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_formula(*n), "[2]'My Book'!A1");
}

TEST(AstFormat, ExternalRefQuotesNumericSheet) {
  Arena a;
  Reference cell;
  AstNode* n = make_external_ref(a, 2, "2026", cell);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_formula(*n), "[2]'2026'!A1");
}

TEST(AstFormat, StructuredRefRoundTripsColumn) {
  ExpectRoundTripsToSame("=Tbl[Region]");
}
TEST(AstFormat, StructuredRefRoundTripsAt) {
  ExpectRoundTripsToSame("=Tbl[@Region]");
}
TEST(AstFormat, StructuredRefRoundTripsAll) {
  ExpectRoundTripsToSame("=Tbl[#All]");
}
TEST(AstFormat, StructuredRefRoundTripsHeaderColumn) {
  ExpectRoundTripsToSame("=Tbl[[#Headers],[Region]]");
}

// ---------------------------------------------------------------------------
// Operator precedence and associativity
// ---------------------------------------------------------------------------

TEST(AstFormat, AddMulPrecedence) {
  ExpectRoundTripsToSame("=1+2*3");
}
TEST(AstFormat, ParenthesisedAddMul) {
  ExpectRoundTripsToSame("=(1+2)*3");
}
TEST(AstFormat, RightAssociativePower) {
  ExpectRoundTripsToSame("=2^3^4");
}
TEST(AstFormat, PowerLhsParen) {
  ExpectRoundTripsToSame("=(2+1)^3");
}
TEST(AstFormat, ConcatBindsBelowAdd) {
  ExpectRoundTripsToSame("=A1&B1+C1");
}
TEST(AstFormat, ComparisonAtBottom) {
  ExpectRoundTripsToSame("=A1+1<B1*2");
}
TEST(AstFormat, UnaryMinusBindsTighterThanMul) {
  ExpectRoundTripsToSame("=-A1*B1");
}
TEST(AstFormat, PercentPostfix) {
  ExpectRoundTripsToSame("=A1%");
}

// ---------------------------------------------------------------------------
// Range / union / intersect / implicit intersection
// ---------------------------------------------------------------------------

TEST(AstFormat, RangeOp) {
  ExpectRoundTripsToSame("=A1:B2");
}
TEST(AstFormat, RangeOpInsideSum) {
  ExpectRoundTripsToSame("=SUM(A1:B2)");
}
TEST(AstFormat, UnionInsideSum) {
  ExpectRoundTripsToSame("=SUM((A1:B2,C1:D2))");
}
TEST(AstFormat, IntersectOp) {
  ExpectRoundTripsToSame("=A1:A10 A2:A20");
}
TEST(AstFormat, ImplicitIntersection) {
  ExpectRoundTripsToSame("=@A1");
}

// ---------------------------------------------------------------------------
// Calls
// ---------------------------------------------------------------------------

TEST(AstFormat, CallZeroArg) {
  ExpectRoundTripsToSame("=NOW()");
}
TEST(AstFormat, CallOneArg) {
  ExpectRoundTripsToSame("=ABS(-1)");
}
TEST(AstFormat, CallSeveralArgs) {
  ExpectRoundTripsToSame("=SUM(1,2,3)");
}
TEST(AstFormat, NestedCalls) {
  ExpectRoundTripsToSame("=SUM(ABS(-1),MAX(1,2))");
}

// ---------------------------------------------------------------------------
// Array literals
// ---------------------------------------------------------------------------

TEST(AstFormat, ArrayLiteralRow) {
  ExpectRoundTripsToSame("={1,2,3}");
}
TEST(AstFormat, ArrayLiteralColumn) {
  ExpectRoundTripsToSame("={1;2;3}");
}
TEST(AstFormat, ArrayLiteralMatrix) {
  ExpectRoundTripsToSame("={1,2;3,4}");
}
TEST(AstFormat, ArrayLiteralWithText) {
  ExpectRoundTripsToSame("={\"a\",\"b\";\"c\",\"d\"}");
}

// ---------------------------------------------------------------------------
// Lambda / let / lambda call
// ---------------------------------------------------------------------------

TEST(AstFormat, LambdaTwoParams) {
  ExpectRoundTripsToSame("=LAMBDA(x,y,x+y)");
}
TEST(AstFormat, LambdaWithOptional) {
  ExpectRoundTripsToSame("=LAMBDA(x,[y],x+y)");
}
TEST(AstFormat, LambdaZeroParams) {
  ExpectRoundTripsToSame("=LAMBDA(42)");
}

TEST(AstFormat, LetSingle) {
  ExpectRoundTripsToSame("=LET(x,1,x+2)");
}
TEST(AstFormat, LetMultiple) {
  ExpectRoundTripsToSame("=LET(x,1,y,2,x+y)");
}

TEST(AstFormat, LambdaCallImmediate) {
  ExpectRoundTripsToSame("=LAMBDA(x,x+1)(5)");
}

// ---------------------------------------------------------------------------
// Composite shapes built from factories (covers ImplicitIntersection /
// non-default StructuredRef modifiers / external refs that the parser
// cannot itself produce in source form yet).
// ---------------------------------------------------------------------------

// Modifier-aware StructuredRef factories are not directly produced by the
// parser (the parser always packs the entire bracket payload into the
// `column` slot with modifier=None). The formatter still has to emit a
// surface form for hand-built ASTs; verify the output text directly
// rather than via round-trip.
TEST(AstFormat, AstStructuredRefWithDataModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Region", StructuredRefModifier::Data);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_formula(*n), "Tbl[[#Data],[Region]]");
}

TEST(AstFormat, AstStructuredRefWithTotalsModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Region", StructuredRefModifier::Totals);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_formula(*n), "Tbl[[#Totals],[Region]]");
}

TEST(AstFormat, AstStructuredRefWithAtModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "Region", StructuredRefModifier::At);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_formula(*n), "Tbl[@Region]");
}

TEST(AstFormat, AstStructuredRefBareAllModifier) {
  Arena a;
  AstNode* n = make_structured_ref(a, "Tbl", "", StructuredRefModifier::All);
  ASSERT_NE(n, nullptr);
  EXPECT_EQ(format_formula(*n), "Tbl[#All]");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
