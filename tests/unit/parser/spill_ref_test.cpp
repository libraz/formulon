// Copyright 2026 libraz. Licensed under the MIT License.
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

}  // namespace
}  // namespace parser
}  // namespace formulon
