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
// A sheet name opening with a digit is only writable in quoted form -- the
// tokenizer would otherwise read the leading run as a numeric literal.
TEST(AstFormat, ThreeDSpanWithDigitLeadingSheetName) {
  ExpectRoundTripsToSame("=SUM('3S1:Daa'!A1)");
}
TEST(AstFormat, RefWithDigitLeadingSheetName) {
  ExpectRoundTripsToSame("=SUM('3S1'!A1)");
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
TEST(AstFormat, PowerChain) {
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
// `A:C` is two bare column identifiers joined by `:`, so a `:` endpoint that
// is a defined name or a LAMBDA parameter spelled like a column has to be
// held apart from its neighbour or the pair reads back as a whole-column
// range instead of two operands.
TEST(AstFormat, RangeBetweenColumnShapedNames) {
  ExpectRoundTripsToSame("=SUM((RO):(r))");
  Arena a;
  Parser p("=SUM((RO):(r))", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(format_formula(*root), "SUM((RO):r)");
}
TEST(AstFormat, RangeFromColumnShapedNameToCall) {
  ExpectRoundTripsToSame("=LE:(NA())");
  Arena a;
  Parser p("=LE:(NA())", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(format_formula(*root), "(LE):NA()");
}
// Ordinary ranges and a name that is too long to be a column keep the plain
// join.
TEST(AstFormat, RangeEndpointsThatCannotFoldStayBare) {
  ExpectRoundTripsToSame("=A1:B2");
  ExpectRoundTripsToSame("=A:C");
  ExpectRoundTripsToSame("=LAMBDA(x,x:B2)");
  ExpectRoundTripsToSame("=A1:OFFSET(A1,1,1)");
}

// The whole-axis splice turns `RangeOp(A:A, C:C)` into the compact `A:C`.
// It must not fire when both endpoints sit on the same axis: `A:A` reads
// back as one whole-column reference rather than as the pair, so the pair
// the source spelled out would be lost. Anchoring does not separate them
// either -- `$A:A` is also a single token.
TEST(AstFormat, WholeAxisPairOnOneColumnKeepsBothEndpoints) {
  ExpectRoundTripsToSame("=A:A:A:A");
  ExpectRoundTripsToSame("=$A:$A:A:A");
  ExpectRoundTripsToSame("=A:A:A:A:A:A");
  Arena a;
  Parser p("=A:A:A:A", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(format_formula(*root), "A:A:A:A");
}
TEST(AstFormat, WholeAxisPairOnOneRowKeepsBothEndpoints) {
  ExpectRoundTripsToSame("=1:1:1:1");
  ExpectRoundTripsToSame("=1:1:$1:1");
  ExpectRoundTripsToSame("=1:1:1:1:3:3");
}
// A chain whose inner pair is a single axis is where the loss became visible
// in the emitted text: the inner pair compacted to one reference, and the
// enclosing range then compacted again on the next pass.
TEST(AstFormat, ChainedWholeAxisRangeIsStable) {
  ExpectRoundTripsToSame("=h:h:h:h:e:e");
  ExpectRoundTripsToSame("=SUM(h:h:h:h:e:e)");
  Arena a;
  Parser p("=h:h:h:h:e:e", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(format_formula(*root), "H:H:H:H:E:E");
}
// Endpoints on different axes still compact, which is the whole point of the
// splice: without it these would emit `A:A:C:C` and `1:1:3:3`.
TEST(AstFormat, WholeAxisPairOnDifferentAxesStillCompacts) {
  Arena a;
  Parser cols("=A:C", a);
  AstNode* cols_root = cols.parse();
  ASSERT_NE(cols_root, nullptr);
  EXPECT_EQ(format_formula(*cols_root), "A:C");
  Parser rows("=1:3", a);
  AstNode* rows_root = rows.parse();
  ASSERT_NE(rows_root, nullptr);
  EXPECT_EQ(format_formula(*rows_root), "1:3");
  Parser sheet("=Sheet1!A:C", a);
  AstNode* sheet_root = sheet.parse();
  ASSERT_NE(sheet_root, nullptr);
  EXPECT_EQ(format_formula(*sheet_root), "Sheet1!A:C");
}
// A round trip must not re-anchor a reference: the pair `$A:$A` with `A:A`
// widens under a one-column fill, the single `$A:$A` it used to compact to
// does not.
TEST(AstFormat, WholeAxisPairKeepsPerEndpointAnchoring) {
  Arena a;
  Parser p("=$A:$A:A:A", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const std::string text = format_formula(*root);
  Arena b;
  Parser reparse(text, b);
  AstNode* reparsed = reparse.parse();
  ASSERT_NE(reparsed, nullptr);
  ASSERT_TRUE(reparse.errors().empty());
  EXPECT_EQ(dump_sexpr(*reparsed), "(range (ref $A:$A) (ref A:A))");
}

TEST(AstFormat, IntersectOp) {
  ExpectRoundTripsToSame("=A1:A10 A2:A20");
}
// A bare cell reference on the left of an intersection formats without any
// bracketing token between it and the right operand, so the re-parse has to
// read `A1 (...)` as an intersection rather than a call.
TEST(AstFormat, IntersectOpWithBareRefLhs) {
  ExpectRoundTripsToSame("=A1 (B1:C5)");
}
TEST(AstFormat, IntersectOpWithUnionRhs) {
  ExpectRoundTripsToSame("=A1 (B1,C1)");
}
TEST(AstFormat, IntersectOpWithUnionLhs) {
  ExpectRoundTripsToSame("=(A1,B1) C1:D1");
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
// The array parser folds a leading sign into the element value, and Excel's
// array-constant grammar has no room for the parenthesisation used in
// operator slots, so a negative element must come back out with a bare sign.
TEST(AstFormat, ArrayLiteralNegativeElements) {
  ExpectRoundTripsToSame("={-1,2}");
  ExpectRoundTripsToSame("={1,-2;3,-4}");
  ExpectRoundTripsToSame("={\"a\",-1}");
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
