//
// Parser tests for the spilled-range postfix operator `#`.
//
// The `#` is a postfix operator that attaches to a single-cell `Ref` atom
// and yields a `NodeKind::SpillRef`. The shape rules tested here:
//
//   * `=A1#`         -> single SpillRef atom.
//   * `=Sheet1!A1#`  -> qualified SpillRef.
//   * `=#`           -> still UnsupportedConstruct (Hash as a primary).
//   * `=A1:B2#`      -> error (`#` cannot follow a range).
//   * `=A1#:B2`      -> error (SpillRef cannot start a range).
//
// The anchor may also be computed rather than written out. Excel accepts
// `#` after anything that names a reference -- a call, a name, a
// parenthesised reference -- and refuses it after a literal, so those two
// groups are tested separately.

#include <algorithm>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/parse_error.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

bool HasErrorCode(const std::vector<ParseError>& errs, ParseErrorCode code) {
  return std::any_of(errs.begin(), errs.end(), [code](const ParseError& e) { return e.code == code; });
}

TEST(SpillRef, A1HashParsesAsSpillRef) {
  Arena a;
  Parser p("=A1#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty()) << "unexpected errors";
  EXPECT_EQ(root->kind(), NodeKind::SpillRef);
  // Sanity: the anchor reference round-trips to canonical A1.
  EXPECT_EQ(dump_sexpr(*root), "(spill-ref A1#)");
}

TEST(SpillRef, AbsoluteAnchorParses) {
  // The anchor's $ markers are preserved on the SpillRef payload so the
  // dumper round-trips them exactly like an ordinary Ref.
  Arena a;
  Parser p("=$B$5#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  EXPECT_EQ(root->kind(), NodeKind::SpillRef);
  EXPECT_EQ(dump_sexpr(*root), "(spill-ref $B$5#)");
}

TEST(SpillRef, QualifiedHashParses) {
  Arena a;
  Parser p("=Sheet2!A1#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  EXPECT_EQ(root->kind(), NodeKind::SpillRef);
  EXPECT_EQ(dump_sexpr(*root), "(spill-ref Sheet2!A1#)");
}

TEST(SpillRef, BareHashStillErrors) {
  // `=#` has no Excel meaning. The tokenizer only emits TokenKind::Hash
  // when the byte is adjacent to a preceding CellRef; standalone `#`
  // falls through the error-literal path and surfaces as an Invalid
  // token (LexerInvalidErrorLiteral) plus a parser UnexpectedToken
  // diagnostic. Either is acceptable for this guard — what matters is
  // that the parser does not silently accept the bare `#`.
  Arena a;
  Parser p("=#", a);
  (void)p.parse();
  ASSERT_FALSE(p.errors().empty());
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LexerInvalidErrorLiteral) ||
              HasErrorCode(p.errors(), ParseErrorCode::UnexpectedToken) ||
              HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct));
}

TEST(SpillRef, RangeHashErrors) {
  // `=A1:B2#`: the postfix `#` binds tighter than `:` so the parse first
  // reduces to `RangeOp(A1, SpillRef(B2))`. The `:` shape check then rejects
  // the SpillRef RHS as InvalidRange.
  Arena a;
  Parser p("=A1:B2#", a);
  (void)p.parse();
  ASSERT_FALSE(p.errors().empty());
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::InvalidRange));
}

TEST(SpillRef, HashAsLhsOfColonErrors) {
  // `=A1#:B2`: the postfix `#` binds first, yielding SpillRef(A1) on the
  // LHS of `:`. The `:` shape check rejects SpillRef on either side because
  // the spill operator is terminal — its result is an array, not a range
  // endpoint.
  Arena a;
  Parser p("=A1#:B2", a);
  (void)p.parse();
  ASSERT_FALSE(p.errors().empty());
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::InvalidRange));
}

TEST(SpillRef, HashAfterFullColumnIsRejected) {
  // `=A:A` is a full-column Ref formed by `Ident Colon Ident` — neither
  // side is a CellRef token, so the tokenizer's `last_cellref_end_byte_`
  // never updates and the trailing `#` falls through the error-literal
  // path (Invalid + LexerInvalidErrorLiteral). The net effect is the
  // same as in production Mac Excel: `=A:A#` is rejected. We accept any
  // of the relevant diagnostic codes here.
  Arena a;
  Parser p("=A:A#", a);
  (void)p.parse();
  ASSERT_FALSE(p.errors().empty());
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LexerInvalidErrorLiteral) ||
              HasErrorCode(p.errors(), ParseErrorCode::UnexpectedToken) ||
              HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct));
}

TEST(SpillRef, ArithmeticBetweenSpillRefs) {
  // Two SpillRefs combined arithmetically: `=A1#+B1#`. The `#` postfix is
  // tighter than `+`, so the AST is `BinaryOp(+, SpillRef(A1), SpillRef(B1))`.
  Arena a;
  Parser p("=A1#+B1#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  EXPECT_EQ(root->kind(), NodeKind::BinaryOp);
  EXPECT_EQ(dump_sexpr(*root), "(binary + (spill-ref A1#) (spill-ref B1#))");
}

TEST(SpillRef, CallAnchorParsesAsComputedSpillRef) {
  // `OFFSET(A1,1,0)#`: the anchor is whatever the call names, so the node
  // carries the sub-expression rather than a Reference.
  Arena a;
  Parser p("=OFFSET(A1,1,0)#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty()) << "unexpected errors";
  ASSERT_EQ(root->kind(), NodeKind::SpillRef);
  ASSERT_NE(root->as_spill_ref_anchor_expr(), nullptr);
  EXPECT_EQ(root->as_spill_ref_anchor_expr()->kind(), NodeKind::Call);
  EXPECT_EQ(dump_sexpr(*root), "(spill-ref (call OFFSET (ref A1) (num 1) (num 0))#)");
}

TEST(SpillRef, WrittenOutAnchorCarriesNoExpression) {
  // The literal form keeps its Reference inline; the two forms are told
  // apart by the anchor-expression accessor being null.
  Arena a;
  Parser p("=A1#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::SpillRef);
  EXPECT_EQ(root->as_spill_ref_anchor_expr(), nullptr);
}

TEST(SpillRef, ParenthesisedRefAnchorParses) {
  // `(A4)#` -- the parentheses collapse, so this reaches the operator as
  // an ordinary written-out anchor.
  Arena a;
  Parser p("=(A4)#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::SpillRef);
  EXPECT_EQ(dump_sexpr(*root), "(spill-ref A4#)");
}

TEST(SpillRef, NameAnchorParsesAsComputedSpillRef) {
  // A defined name or a LET binding can anchor a spill.
  Arena a;
  Parser p("=Anchor#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::SpillRef);
  ASSERT_NE(root->as_spill_ref_anchor_expr(), nullptr);
  EXPECT_EQ(root->as_spill_ref_anchor_expr()->kind(), NodeKind::NameRef);
}

TEST(SpillRef, NestedCallAnchorParses) {
  Arena a;
  Parser p("=OFFSET(OFFSET(A1,3,0),0,0)#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::SpillRef);
  EXPECT_NE(root->as_spill_ref_anchor_expr(), nullptr);
}

TEST(SpillRef, ParenthesisedArithmeticAnchorIsRejected) {
  // A closing parenthesis arms the operator, so `(1+2)#` reaches the
  // parser as a `#` over an arithmetic result. Excel refuses that at
  // entry; here it is an unsupported construct.
  Arena a;
  Parser p("=(1+2)#", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct));
}

TEST(SpillRef, LiteralAnchorIsRejected) {
  // Excel refuses `1#` and `"A1"#` at entry: neither names a reference.
  // Nothing arms the operator after a number or a string, so the `#` is
  // read as the opening byte of an error literal and fails there instead
  // of reaching the postfix rule. Either way the formula does not parse,
  // which is the contract these spellings need.
  for (const char* src : {"=1#", "=\"A1\"#"}) {
    Arena a;
    Parser p(src, a);
    AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << src;
    EXPECT_FALSE(p.errors().empty()) << src;
  }
}

TEST(SpillRef, ErrorLiteralAfterAnAnchorTailStillLexes) {
  // Widening the `#` disambiguation must not swallow an error literal that
  // merely follows a call or a name. An operator always separates them, so
  // `#` there still opens `#REF!`.
  for (const char* src : {"=SUM(A1)+#REF!", "=Anchor+#N/A", "=A1+#REF!", "=#REF!"}) {
    Arena a;
    Parser p(src, a);
    AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << src;
    EXPECT_TRUE(p.errors().empty()) << src;
  }
}

}  // namespace
}  // namespace parser
}  // namespace formulon
