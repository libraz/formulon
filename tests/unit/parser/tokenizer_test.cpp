// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the M2.2 tokenizer. Each group exercises a specific
// syntactic family; see `backup/plans/02-calc-engine.md` §2.2 for the
// authoritative token catalog.
//
// Note on token lifetimes: `Token::text` (for String / SheetName) references
// arena memory owned by the producing `Tokenizer`. Tests therefore keep the
// `Tokenizer` alive on the stack and read `tokens()` by `const&`.

#include "parser/tokenizer.h"

#include <cmath>
#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "parser/lexer_error.h"
#include "parser/token.h"
#include "value.h"

namespace formulon {
namespace parser {
namespace {

// Copies every token's `kind` (including the terminating Eof) into a vector
// so structural expectations remain succinct. Safe to return: copying kinds
// does not require arena memory to stay alive.
std::vector<TokenKind> KindsOf(std::string_view src, TokenizerOptions opts = {}) {
  Tokenizer t(src, opts);
  const auto& v = t.tokens();
  std::vector<TokenKind> kinds;
  kinds.reserve(v.size());
  for (const auto& tok : v) {
    kinds.push_back(tok.kind);
  }
  return kinds;
}

// ---------------------------------------------------------------------------
// NumberLiterals
// ---------------------------------------------------------------------------

TEST(TokenizerNumberLiterals, SimpleInteger) {
  Tokenizer tz("42");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);  // Number + Eof
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_EQ(v[0].number, 42.0);
  EXPECT_TRUE(v[0].is_integer);
}

TEST(TokenizerNumberLiterals, Decimal) {
  Tokenizer tz("3.14");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_DOUBLE_EQ(v[0].number, 3.14);
  EXPECT_FALSE(v[0].is_integer);
}

TEST(TokenizerNumberLiterals, LeadingDot) {
  Tokenizer tz(".5");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_DOUBLE_EQ(v[0].number, 0.5);
  EXPECT_FALSE(v[0].is_integer);
}

TEST(TokenizerNumberLiterals, LowerExponent) {
  Tokenizer tz("1e5");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_DOUBLE_EQ(v[0].number, 1e5);
  EXPECT_FALSE(v[0].is_integer);
}

TEST(TokenizerNumberLiterals, UpperExponentSigned) {
  Tokenizer tz("1.5E-3");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_DOUBLE_EQ(v[0].number, 1.5e-3);
}

TEST(TokenizerNumberLiterals, Zero) {
  Tokenizer tz("0");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_EQ(v[0].number, 0.0);
  EXPECT_TRUE(v[0].is_integer);
}

TEST(TokenizerNumberLiterals, OneMillion) {
  Tokenizer tz("1000000");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_EQ(v[0].number, 1000000.0);
  EXPECT_TRUE(v[0].is_integer);
}

TEST(TokenizerNumberLiterals, RejectEmptyExponent) {
  Tokenizer tz("1e");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidNumberLiteral);
}

TEST(TokenizerNumberLiterals, RejectDoubleDot) {
  Tokenizer tz("1.2.3");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidNumberLiteral);
}

TEST(TokenizerNumberLiterals, RejectTrailingSignOnly) {
  Tokenizer tz("1e+");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidNumberLiteral);
}

TEST(TokenizerNumberLiterals, RejectBareDot) {
  // A lone '.' is not dispatched to scan_number (the main loop requires a
  // following digit), so it falls through as InvalidCharacter.
  Tokenizer tz(".");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
}

// ---------------------------------------------------------------------------
// StringLiterals
// ---------------------------------------------------------------------------

TEST(TokenizerStringLiterals, Plain) {
  Tokenizer tz("\"abc\"");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::String);
  EXPECT_EQ(std::string(v[0].text), "abc");
}

TEST(TokenizerStringLiterals, Empty) {
  Tokenizer tz("\"\"");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::String);
  EXPECT_EQ(std::string(v[0].text), "");
}

TEST(TokenizerStringLiterals, DoubleQuoteEscape) {
  Tokenizer tz("\"a\"\"b\"");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::String);
  EXPECT_EQ(std::string(v[0].text), "a\"b");
}

TEST(TokenizerStringLiterals, Japanese) {
  Tokenizer tz("\"\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E\"");  // "日本語"
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::String);
  EXPECT_EQ(std::string(v[0].text), "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
}

TEST(TokenizerStringLiterals, Unterminated) {
  Tokenizer tz("\"abc");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::UnterminatedString);
}

// ---------------------------------------------------------------------------
// BoolLiterals
// ---------------------------------------------------------------------------

TEST(TokenizerBoolLiterals, UpperTrue) {
  Tokenizer tz("TRUE");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Bool);
  EXPECT_TRUE(v[0].boolean);
}

TEST(TokenizerBoolLiterals, LowerFalse) {
  Tokenizer tz("false");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Bool);
  EXPECT_FALSE(v[0].boolean);
}

TEST(TokenizerBoolLiterals, MixedTrue) {
  Tokenizer tz("True");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Bool);
  EXPECT_TRUE(v[0].boolean);
}

// ---------------------------------------------------------------------------
// ErrorLiterals
// ---------------------------------------------------------------------------

TEST(TokenizerErrorLiterals, AllSeventeen) {
  const struct {
    const char* text;
    ErrorCode code;
  } cases[] = {
      {"#NULL!", ErrorCode::Null},       {"#DIV/0!", ErrorCode::Div0},
      {"#VALUE!", ErrorCode::Value},     {"#REF!", ErrorCode::Ref},
      {"#NAME?", ErrorCode::Name},       {"#NUM!", ErrorCode::Num},
      {"#N/A", ErrorCode::NA},           {"#GETTING_DATA", ErrorCode::GettingData},
      {"#SPILL!", ErrorCode::Spill},     {"#CALC!", ErrorCode::Calc},
      {"#FIELD!", ErrorCode::Field},     {"#BLOCKED!", ErrorCode::Blocked},
      {"#CONNECT!", ErrorCode::Connect}, {"#EXTERNAL!", ErrorCode::External},
      {"#BUSY!", ErrorCode::Busy},       {"#PYTHON!", ErrorCode::Python},
      {"#UNKNOWN!", ErrorCode::Unknown},
  };
  for (const auto& c : cases) {
    Tokenizer tz(c.text);
    const auto& v = tz.tokens();
    ASSERT_EQ(v.size(), 2u) << c.text;
    EXPECT_EQ(v[0].kind, TokenKind::ErrorLiteral) << c.text;
    EXPECT_EQ(v[0].error_code, c.code) << c.text;
    EXPECT_TRUE(tz.errors().empty()) << c.text;
  }
}

TEST(TokenizerErrorLiterals, InvalidSpelling) {
  Tokenizer tz("#NOPE!");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidErrorLiteral);
}

// ---------------------------------------------------------------------------
// Identifiers
// ---------------------------------------------------------------------------

TEST(TokenizerIdentifiers, Simple) {
  Tokenizer tz("sum");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Ident);
  EXPECT_EQ(std::string(v[0].lexeme), "sum");
}

TEST(TokenizerIdentifiers, WithDigitsAndUnderscore) {
  Tokenizer tz("MyVar_1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Ident);
  EXPECT_EQ(std::string(v[0].lexeme), "MyVar_1");
}

TEST(TokenizerIdentifiers, Japanese) {
  Tokenizer tz("\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E\xE9\x96\xA2\xE6\x95\xB0");  // "日本語関数"
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Ident);
}

TEST(TokenizerIdentifiers, XlfnPrefix) {
  Tokenizer tz("_xlfn.FILTER");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Ident);
  EXPECT_EQ(std::string(v[0].lexeme), "_xlfn.FILTER");
}

// ---------------------------------------------------------------------------
// CellRefs
// ---------------------------------------------------------------------------

TEST(TokenizerCellRefs, Bare) {
  Tokenizer tz("A1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::CellRef);
  EXPECT_EQ(std::string(v[0].lexeme), "A1");
}

TEST(TokenizerCellRefs, FullyAnchored) {
  Tokenizer tz("$A$1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::CellRef);
  EXPECT_EQ(std::string(v[0].lexeme), "$A$1");
}

TEST(TokenizerCellRefs, ExcelMaxima) {
  Tokenizer tz("XFD1048576");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::CellRef);
}

TEST(TokenizerCellRefs, MixedAnchoredDouble) {
  Tokenizer tz("$AA$99");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::CellRef);
  EXPECT_EQ(std::string(v[0].lexeme), "$AA$99");
}

TEST(TokenizerCellRefs, OverflowColumn) {
  // XFE is past the column cap; falls back to Ident.
  Tokenizer tz("XFE1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Ident);
}

TEST(TokenizerCellRefs, ColumnOnlyBecomesIdents) {
  // Documented M2.2 limitation: A:A is IDENT COLON IDENT, left for the
  // parser to promote.
  auto v = KindsOf("A:A");
  std::vector<TokenKind> expected = {TokenKind::Ident, TokenKind::Colon, TokenKind::Ident, TokenKind::Eof};
  EXPECT_EQ(v, expected);
}

TEST(TokenizerCellRefs, RowOnlyBecomesNumbers) {
  auto v = KindsOf("1:1");
  std::vector<TokenKind> expected = {TokenKind::Number, TokenKind::Colon, TokenKind::Number, TokenKind::Eof};
  EXPECT_EQ(v, expected);
}

// ---------------------------------------------------------------------------
// SheetNames
// ---------------------------------------------------------------------------

TEST(TokenizerSheetNames, Unquoted) {
  auto v = KindsOf("Sheet1!A1");
  std::vector<TokenKind> expected = {TokenKind::Ident, TokenKind::Bang, TokenKind::CellRef, TokenKind::Eof};
  EXPECT_EQ(v, expected);
}

TEST(TokenizerSheetNames, Quoted) {
  Tokenizer tz("'Sheet 1'!A1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 4u);  // SheetName + Bang + CellRef + Eof
  EXPECT_EQ(v[0].kind, TokenKind::SheetName);
  EXPECT_EQ(std::string(v[0].text), "Sheet 1");
  EXPECT_EQ(v[1].kind, TokenKind::Bang);
  EXPECT_EQ(v[2].kind, TokenKind::CellRef);
}

TEST(TokenizerSheetNames, QuotedEscape) {
  Tokenizer tz("'O''Brien'");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::SheetName);
  EXPECT_EQ(std::string(v[0].text), "O'Brien");
}

TEST(TokenizerSheetNames, Unterminated) {
  Tokenizer tz("'Sheet 1");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::UnterminatedSheetQuote);
}

// ---------------------------------------------------------------------------
// Operators
// ---------------------------------------------------------------------------

TEST(TokenizerOperators, AllSingleChar) {
  // `<>` is one token; everything else is one-per-byte.
  auto v = KindsOf("+-*/^%&=<>@");
  std::vector<TokenKind> expected = {
      TokenKind::Plus,      TokenKind::Minus, TokenKind::Star,  TokenKind::Slash, TokenKind::Caret, TokenKind::Percent,
      TokenKind::Ampersand, TokenKind::Eq,    TokenKind::NotEq, TokenKind::At,    TokenKind::Eof,
  };
  EXPECT_EQ(v, expected);
}

TEST(TokenizerOperators, DistinguishesCompoundComparison) {
  auto v = KindsOf("<= >= <> < > =");
  // Filter whitespace to compare operator ordering only.
  std::vector<TokenKind> kinds;
  for (auto k : v) {
    if (k != TokenKind::Whitespace) {
      kinds.push_back(k);
    }
  }
  std::vector<TokenKind> expected = {TokenKind::LtEq, TokenKind::GtEq, TokenKind::NotEq, TokenKind::Lt,
                                     TokenKind::Gt,   TokenKind::Eq,   TokenKind::Eof};
  EXPECT_EQ(kinds, expected);
}

TEST(TokenizerOperators, PlainLtAndGt) {
  auto v = KindsOf("<");
  std::vector<TokenKind> expected = {TokenKind::Lt, TokenKind::Eof};
  EXPECT_EQ(v, expected);
  v = KindsOf(">");
  expected = {TokenKind::Gt, TokenKind::Eof};
  EXPECT_EQ(v, expected);
}

// ---------------------------------------------------------------------------
// ArrayLiteral
// ---------------------------------------------------------------------------

TEST(TokenizerArrayLiteral, TwoByTwo) {
  auto v = KindsOf("{1,2;3,4}");
  std::vector<TokenKind> expected = {
      TokenKind::LBrace, TokenKind::Number, TokenKind::Comma,  TokenKind::Number, TokenKind::Semicolon,
      TokenKind::Number, TokenKind::Comma,  TokenKind::Number, TokenKind::RBrace, TokenKind::Eof,
  };
  EXPECT_EQ(v, expected);
}

// ---------------------------------------------------------------------------
// SpilledRangeOp
// ---------------------------------------------------------------------------

TEST(TokenizerSpilledRangeOp, Adjacent) {
  auto v = KindsOf("A1#");
  std::vector<TokenKind> expected = {TokenKind::CellRef, TokenKind::Hash, TokenKind::Eof};
  EXPECT_EQ(v, expected);
}

TEST(TokenizerSpilledRangeOp, WithWhitespaceBecomesInvalid) {
  // Design decision: `A1 #` - the whitespace breaks spill-adjacency, so the
  // trailing `#` falls into the error-literal scanner and is flagged as
  // InvalidErrorLiteral. The parser layer may later relax this once the
  // full spilled-range grammar is wired in (M2.3).
  Tokenizer tz("A1 #");
  const auto& v = tz.tokens();
  ASSERT_GE(v.size(), 3u);
  EXPECT_EQ(v[0].kind, TokenKind::CellRef);
  EXPECT_EQ(v[1].kind, TokenKind::Whitespace);
  EXPECT_FALSE(tz.errors().empty());
}

// ---------------------------------------------------------------------------
// Whitespace
// ---------------------------------------------------------------------------

TEST(TokenizerWhitespace, PreservedAsToken) {
  auto v = KindsOf("A1 A2");
  std::vector<TokenKind> expected = {TokenKind::CellRef, TokenKind::Whitespace, TokenKind::CellRef, TokenKind::Eof};
  EXPECT_EQ(v, expected);
}

TEST(TokenizerWhitespace, FullwidthSpaceIsInvalid) {
  // U+3000 is three UTF-8 bytes (E3 80 80) and must be flagged.
  Tokenizer tz(
      "\xE3\x80\x80"
      "A1");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidCharacter);
}

// ---------------------------------------------------------------------------
// BOMHandling
// ---------------------------------------------------------------------------

TEST(TokenizerBomHandling, Utf8BomAtStartSkipped) {
  Tokenizer tz(
      "\xEF\xBB\xBF"
      "A1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);  // CellRef + Eof
  EXPECT_EQ(v[0].kind, TokenKind::CellRef);
  EXPECT_TRUE(tz.errors().empty());
}

TEST(TokenizerBomHandling, Utf16LeBomAtStartSkipped) {
  Tokenizer tz(
      "\xFF\xFE"
      "A1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::CellRef);
  EXPECT_TRUE(tz.errors().empty());
}

TEST(TokenizerBomHandling, MidInputBomIsInvalid) {
  Tokenizer tz(
      "A1\xEF\xBB\xBF"
      "B2");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidCharacter);
}

// ---------------------------------------------------------------------------
// SurrogatePairOffsets
// ---------------------------------------------------------------------------

TEST(TokenizerSurrogatePairOffsets, EmojiInsideString) {
  // U+1F600 is 4 UTF-8 bytes and 2 UTF-16 code units. Plus two quotes:
  // "\"\xF0\x9F\x98\x80\"" spans 1 + 2 + 1 = 4 UTF-16 units.
  Tokenizer tz("\"\xF0\x9F\x98\x80\"");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::String);
  EXPECT_EQ(v[0].range.start, 0u);
  EXPECT_EQ(v[0].range.end, 4u);
  EXPECT_EQ(v[1].range.start, 4u);
  EXPECT_EQ(v[1].range.end, 4u);
}

// ---------------------------------------------------------------------------
// ExcessiveLength
// ---------------------------------------------------------------------------

TEST(TokenizerExcessiveLength, Truncates) {
  TokenizerOptions opts;
  opts.max_formula_length_utf16 = 5;
  // 8 ASCII chars, cap at 5 UTF-16 units. The scanner consumes the first
  // run as a single Ident (8 bytes, 8 UTF-16 units), which does not
  // re-enter the main loop between codepoints and therefore does not trip
  // the cap. Use whitespace so we get multiple token boundaries that hit
  // the cap mid-stream.
  Tokenizer tz("A B C D E F G H", opts);
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::ExcessiveLength);
}

}  // namespace
}  // namespace parser
}  // namespace formulon
