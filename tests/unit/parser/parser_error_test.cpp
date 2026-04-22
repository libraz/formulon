// Copyright 2026 libraz. Licensed under the MIT License.
//
// Minimal error-path tests for the Pratt parser. The full negative corpus,
// panic-mode recovery, and the suggestion engine are deferred to follow-up
// work; these tests just pin the contract that the parser stops at the
// first hard failure and emits an appropriate single error code.

#include "parser/parser.h"

#include <algorithm>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parse_error.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

// Returns true iff `errors` contains at least one entry whose code is `code`.
// Used because some failures (e.g. lexer errors) come paired with a parser
// stop and we only want to assert the meaningful one is present.
bool HasErrorCode(const std::vector<ParseError>& errors, ParseErrorCode code) {
  return std::any_of(errors.begin(), errors.end(),
                     [code](const ParseError& e) { return e.code == code; });
}

TEST(ParserErrors, EmptyInputIsUnexpectedEof) {
  Arena a;
  Parser p("", a);
  EXPECT_EQ(p.parse(), nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnexpectedEof));
}

TEST(ParserErrors, BareEqualsIsUnexpectedEof) {
  Arena a;
  Parser p("=", a);
  EXPECT_EQ(p.parse(), nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnexpectedEof));
}

TEST(ParserErrors, TrailingPlusReportsEof) {
  Arena a;
  Parser p("=1+", a);
  EXPECT_EQ(p.parse(), nullptr);
  ASSERT_FALSE(p.errors().empty());
  const ParseErrorCode c = p.errors().back().code;
  EXPECT_TRUE(c == ParseErrorCode::UnexpectedEof || c == ParseErrorCode::ExpectedExpression);
}

TEST(ParserErrors, UnclosedParen) {
  Arena a;
  Parser p("=(1+2", a);
  EXPECT_EQ(p.parse(), nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnclosedParen));
}

TEST(ParserErrors, UnclosedArrayBrace) {
  Arena a;
  Parser p("={1,2", a);
  EXPECT_EQ(p.parse(), nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnclosedBrace));
}

TEST(ParserErrors, ArrayRowMismatch) {
  Arena a;
  Parser p("={1,2;3}", a);
  EXPECT_EQ(p.parse(), nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::ArrayRowMismatch));
}

TEST(ParserErrors, TrailingCommaInCallReportsExpression) {
  Arena a;
  Parser p("=SUM(1,)", a);
  EXPECT_EQ(p.parse(), nullptr);
  ASSERT_FALSE(p.errors().empty());
  const ParseErrorCode c = p.errors().back().code;
  EXPECT_TRUE(c == ParseErrorCode::ExpectedExpression || c == ParseErrorCode::UnexpectedToken);
}

TEST(ParserErrors, UnclosedCallParen) {
  Arena a;
  Parser p("=SUM(1", a);
  EXPECT_EQ(p.parse(), nullptr);
  ASSERT_FALSE(p.errors().empty());
  const ParseErrorCode c = p.errors().back().code;
  EXPECT_TRUE(c == ParseErrorCode::UnclosedParen || c == ParseErrorCode::ExpectedRParenOrComma);
}

TEST(ParserErrors, StringLiteralIsUnsupported) {
  Arena a;
  Parser p("=\"hi\"", a);
  EXPECT_EQ(p.parse(), nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct));
}

TEST(ParserErrors, StructuredRefBracketIsUnsupported) {
  Arena a;
  Parser p("=Table[col]", a);
  // Without dedicated structured-ref support, the parser sees `Table` as a
  // NameRef and `[` as the start of an unsupported construct. We assert the
  // unsupported diagnostic surfaces somewhere in the error list (the AST
  // root is nullptr because `[` cannot follow a valid expression).
  EXPECT_EQ(p.parse(), nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct) ||
              HasErrorCode(p.errors(), ParseErrorCode::UnexpectedToken));
}

TEST(ParserErrors, SpilledHashIsUnsupported) {
  Arena a;
  Parser p("=A1#", a);
  EXPECT_EQ(p.parse(), nullptr);
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct) ||
              HasErrorCode(p.errors(), ParseErrorCode::UnexpectedToken));
}

TEST(ParserErrors, LexerErrorIsPromoted) {
  // Unterminated string at the lexer level. The parser will stop at the
  // resulting String token (UnsupportedConstruct), but we also expect the
  // promoted lexer error to appear in the error list.
  Arena a;
  Parser p("=\"abc", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LexerUnterminatedString));
}

TEST(ParserErrors, ExpectedRParenOrCommaInCall) {
  Arena a;
  Parser p("=SUM(1 2)", a);
  EXPECT_EQ(p.parse(), nullptr);
  ASSERT_FALSE(p.errors().empty());
  const ParseErrorCode c = p.errors().back().code;
  EXPECT_TRUE(c == ParseErrorCode::ExpectedRParenOrComma || c == ParseErrorCode::UnexpectedToken);
}

}  // namespace
}  // namespace parser
}  // namespace formulon
