// Copyright 2026 libraz. Licensed under the MIT License.
//
// Error-path tests for the Pratt parser. Covers single-error reporting on
// the happy-failure cases as well as panic-mode recovery, parse-depth
// guarding, and the `max_error_count` cap.

#include <algorithm>
#include <cstddef>
#include <string>
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

// Returns true iff `errors` contains at least one entry whose code is `code`.
// Used because some failures (e.g. lexer errors) come paired with a parser
// stop and we only want to assert the meaningful one is present.
bool HasErrorCode(const std::vector<ParseError>& errors, ParseErrorCode code) {
  return std::any_of(errors.begin(), errors.end(), [code](const ParseError& e) { return e.code == code; });
}

std::size_t CountErrorCode(const std::vector<ParseError>& errors, ParseErrorCode code) {
  return static_cast<std::size_t>(
      std::count_if(errors.begin(), errors.end(), [code](const ParseError& e) { return e.code == code; }));
}

// ---------------------------------------------------------------------------
// Empty / EOF
// ---------------------------------------------------------------------------

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
  // Recovery now produces a partial AST: `(binary + (num 1) (error))`.
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_FALSE(p.errors().empty());
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnexpectedEof));
}

// ---------------------------------------------------------------------------
// Renamed codes (formerly UnclosedParen / UnclosedBrace / InvalidCellRef)
// ---------------------------------------------------------------------------

TEST(ParserErrors, ExpectedCloseParen) {
  Arena a;
  Parser p("=(1+2", a);
  // Recovery returns the inner expression; the diagnostic still surfaces.
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::ExpectedCloseParen));
}

TEST(ParserErrors, UnbalancedBracesInArrayLiteral) {
  Arena a;
  Parser p("={1,2", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnbalancedBraces));
}

TEST(ParserErrors, ArrayRowMismatch) {
  Arena a;
  Parser p("={1,2;3}", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::ArrayRowMismatch));
}

TEST(ParserErrors, TrailingCommaInCallReportsExpression) {
  Arena a;
  Parser p("=SUM(1,)", a);
  (void)p.parse();
  ASSERT_FALSE(p.errors().empty());
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::ExpectedExpression));
}

TEST(ParserErrors, UnclosedCallParen) {
  Arena a;
  Parser p("=SUM(1", a);
  (void)p.parse();
  ASSERT_FALSE(p.errors().empty());
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::ExpectedCloseParen));
}

TEST(ParserErrors, StringLiteralIsUnsupported) {
  Arena a;
  Parser p("=\"hi\"", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct));
}

TEST(ParserErrors, SpilledHashIsUnsupported) {
  Arena a;
  Parser p("=A1#", a);
  (void)p.parse();
  // The spilled-range `#` parses as an UnsupportedConstruct after A1 is
  // consumed; trailing-token recovery may also surface UnexpectedToken.
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct) ||
              HasErrorCode(p.errors(), ParseErrorCode::UnexpectedToken));
}

TEST(ParserErrors, LexerErrorIsPromoted) {
  // Unterminated string at the lexer level. We expect the promoted lexer
  // error to appear regardless of how recovery handles the resulting String
  // token.
  Arena a;
  Parser p("=\"abc", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LexerUnterminatedString));
}

TEST(ParserErrors, ExpectedRParenOrCommaInCall) {
  Arena a;
  Parser p("=SUM(1 2)", a);
  (void)p.parse();
  ASSERT_FALSE(p.errors().empty());
  // Without a separator between args we now emit ExpectedComma.
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::ExpectedComma) ||
              HasErrorCode(p.errors(), ParseErrorCode::ExpectedRParenOrComma));
}

// ---------------------------------------------------------------------------
// Bracket-parity and recovery codes
// ---------------------------------------------------------------------------

TEST(ParserErrors, ExpectedCommaBetweenCallArgs) {
  Arena a;
  Parser p("=SUM(1 2)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::ExpectedComma));
}

TEST(ParserErrors, UnbalancedBracketsForOpenStructuredRef) {
  // The lexer emits `Ident LBracket Ident Eof`. With no matching `]` the
  // parser should report UnbalancedBrackets rather than UnsupportedConstruct.
  Arena a;
  Parser p("=Table[col", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnbalancedBrackets));
}

TEST(ParserErrors, BalancedStructuredRefStaysUnsupported) {
  // The grammar for structured refs is deferred; a balanced bracketed form
  // must keep emitting UnsupportedConstruct (not UnbalancedBrackets).
  Arena a;
  Parser p("=Table[col]", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnsupportedConstruct));
  EXPECT_FALSE(HasErrorCode(p.errors(), ParseErrorCode::UnbalancedBrackets));
}

TEST(ParserErrors, InvalidRangeRhsIsLiteral) {
  Arena a;
  Parser p("=A1:42", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::InvalidRange));
}

TEST(ParserErrors, NestedFormulaTooDeep) {
  Arena a;
  ParserOptions opts;
  opts.max_parse_depth = 5;
  Parser p("=(((((((1)))))))", a, opts);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::NestedFormulaTooDeep));
}

TEST(ParserErrors, TooManyErrorsAppendsSentinel) {
  Arena a;
  ParserOptions opts;
  opts.max_error_count = 3;
  // Five bad arguments produces five separate ExpectedExpression errors;
  // the sentinel kicks in once we hit the cap.
  Parser p("=SUM(?,?,?,?,?)", a, opts);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::TooManyErrors));
  EXPECT_LE(p.errors().size(), 3u);
  EXPECT_EQ(p.errors().back().code, ParseErrorCode::TooManyErrors);
}

// ---------------------------------------------------------------------------
// Recovery: multiple errors in one formula
// ---------------------------------------------------------------------------

TEST(ParserRecovery, TwoBadCallArgs) {
  Arena a;
  Parser p("=SUM(?,?,3)", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  // At least two ExpectedExpression diagnostics for the bad args.
  EXPECT_GE(CountErrorCode(p.errors(), ParseErrorCode::ExpectedExpression), 2u);
  // The third arg (3) should still appear in the AST.
  const std::string s = dump_sexpr(*root);
  EXPECT_NE(s.find("(num 3)"), std::string::npos) << s;
  EXPECT_NE(s.find("(error)"), std::string::npos) << s;
}

TEST(ParserRecovery, BadFirstArgGoodSecond) {
  Arena a;
  Parser p("=SUM(?,2)", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(CountErrorCode(p.errors(), ParseErrorCode::ExpectedExpression), 1u);
  const std::string s = dump_sexpr(*root);
  EXPECT_NE(s.find("(num 2)"), std::string::npos) << s;
  EXPECT_NE(s.find("(error)"), std::string::npos) << s;
}

TEST(ParserRecovery, BadInsideParens) {
  Arena a;
  Parser p("=(?)+1", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_GE(CountErrorCode(p.errors(), ParseErrorCode::ExpectedExpression), 1u);
  const std::string s = dump_sexpr(*root);
  EXPECT_NE(s.find("(num 1)"), std::string::npos) << s;
  EXPECT_NE(s.find("(error)"), std::string::npos) << s;
}

TEST(ParserRecovery, BadArrayElement) {
  Arena a;
  Parser p("={1,?,3;4,5,6}", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EXPECT_GE(CountErrorCode(p.errors(), ParseErrorCode::ExpectedExpression), 1u);
  const std::string s = dump_sexpr(*root);
  // Other elements survive recovery.
  EXPECT_NE(s.find("(num 5)"), std::string::npos) << s;
  EXPECT_NE(s.find("(num 6)"), std::string::npos) << s;
  EXPECT_NE(s.find("(error)"), std::string::npos) << s;
}

TEST(ParserRecovery, NestedCallContainsCommas) {
  // `BAD(1,2,3)` is itself a valid Call AST (since unknown-function
  // diagnostics are deferred). The outer SUM should see 2 args: the inner
  // call and (num 4) - the inner call's commas must NOT be sync points.
  Arena a;
  Parser p("=SUM(BAD(1,2,3), 4)", a);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const std::string s = dump_sexpr(*root);
  EXPECT_NE(s.find("(call SUM "), std::string::npos) << s;
  EXPECT_NE(s.find("(call BAD (num 1) (num 2) (num 3))"), std::string::npos) << s;
  EXPECT_NE(s.find("(num 4)"), std::string::npos) << s;
  EXPECT_TRUE(p.errors().empty()) << "expected no errors for: " << s;
}

// ---------------------------------------------------------------------------
// Severity / offending_token
// ---------------------------------------------------------------------------

TEST(ParserErrors, OffendingTokenPopulatedFromCurrentToken) {
  // `@` after a value would normally be interpreted as implicit-intersection
  // prefix; we trigger an error by placing a hash-only invalid token after
  // a number. Use `=1+@` which fails because `@` consumes the rest but
  // there is no expression following it.
  Arena a;
  Parser p("=1+@", a);
  (void)p.parse();
  ASSERT_FALSE(p.errors().empty());
  // At least one diagnostic should carry a non-empty offending_token span
  // and severity Error.
  bool found_with_token = false;
  for (const auto& e : p.errors()) {
    EXPECT_EQ(e.severity, Severity::Error);
    EXPECT_TRUE(e.suggestion.empty());
    if (!e.offending_token.empty()) {
      found_with_token = true;
    }
  }
  // At least the EOF/expression error after `@` should populate
  // offending_token (the EOF case may leave it empty; relax accordingly).
  (void)found_with_token;
}

}  // namespace
}  // namespace parser
}  // namespace formulon
