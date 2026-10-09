//
// Unit tests for the Excel formula tokenizer. Each group exercises a
// specific syntactic family.
//
// Note on token lifetimes: `Token::text` (for String / SheetName) references
// arena memory owned by the producing `Tokenizer`. Tests therefore keep the
// `Tokenizer` alive on the stack and read `tokens()` by `const&`.

#include "parser/tokenizer.h"

#include <cmath>
#include <cstdint>
#include <string>
#include <vector>

#include "excel_locale.h"
#include "excel_profile.h"
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

TEST(TokenizerNumberLiterals, TrailingDot) {
  // `1.` -- an empty fractional part -- is a valid Excel literal equal to 1,
  // not the degenerate case `RejectBareDot` below covers.
  Tokenizer tz("1.");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_DOUBLE_EQ(v[0].number, 1.0);
  EXPECT_TRUE(tz.errors().empty());
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

TEST(TokenizerNumberLiterals, OverflowMagnitudeBecomesNumError) {
  // `1E309` overflows the double range; strtod returns +inf. Rather than
  // emitting a non-finite Number token, the tokenizer surfaces `#NUM!` so
  // the formula evaluates to Excel's overflow error.
  Tokenizer tz("1E309");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);  // ErrorLiteral + Eof
  EXPECT_EQ(v[0].kind, TokenKind::ErrorLiteral);
  EXPECT_EQ(v[0].error_code, ErrorCode::Num);
}

TEST(TokenizerNumberLiterals, UnderflowMagnitudeRoundsToZero) {
  // `1E-400` is well past even the smallest subnormal, so strtod itself
  // already floors it to 0.0.
  Tokenizer tz("1E-400");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);  // Number + Eof
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_EQ(v[0].number, 0.0);
}

TEST(TokenizerNumberLiterals, SubnormalMagnitudeFlushesToZero) {
  // Measured directly against Excel 365 (Mac, via xlwings): every literal
  // below DBL_MIN (~2.2250738585072014E-308) -- the whole subnormal range,
  // not just the underflow-to-zero tail strtod already floors -- evaluates
  // to exactly 0 in Excel, which does not keep denormalized precision.
  for (const char* lex : {"2.5E-310", "1E-320", "1E-310", "4.9E-324", "1E-308"}) {
    Tokenizer tz(lex);
    const auto& v = tz.tokens();
    ASSERT_EQ(v.size(), 2u) << lex;
    EXPECT_EQ(v[0].kind, TokenKind::Number) << lex;
    EXPECT_EQ(v[0].number, 0.0) << lex;
  }
}

TEST(TokenizerNumberLiterals, MagnitudeAboveDblMinIsKeptExactly) {
  // `2.3E-308` is above DBL_MIN and must survive as a genuine (non-flushed)
  // double, matching Excel.
  Tokenizer tz("2.3E-308");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_DOUBLE_EQ(v[0].number, 2.3e-308);
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

TEST(TokenizerNumberLiterals, RejectsOverlongNumberLiteralAsInvalid) {
  // 76 byte literal: `0.` + 75 zeros + `1`. Earlier bug silently truncated
  // anything past 63 bytes, letting the surface lexeme and the parsed
  // numeric value disagree. The fix surfaces overlong literals as Invalid
  // with an `InvalidNumberLiteral` diagnostic so callers cannot rely on a
  // mismatched semantic.
  std::string overlong = "0.";
  overlong.append(75, '0');
  overlong.push_back('1');
  ASSERT_GT(overlong.size(), 63u);

  Tokenizer tz(overlong);
  const auto& v = tz.tokens();
  // Tokenizer emits at least one token for the literal plus the trailing
  // Eof; the literal must be the Invalid kind, not Number.
  ASSERT_GE(v.size(), 1u);
  EXPECT_EQ(v[0].kind, TokenKind::Invalid);
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidNumberLiteral);
}

TEST(TokenizerNumberLiterals, FifteenSignificantDigitTruncation) {
  // Excel stores at most 15 significant digits; digits past the fifteenth
  // are zeroed (truncation + zero-fill, not rounding). These are the four
  // oracle-divergence literals plus two that need no truncation.
  struct Case {
    const char* literal;
    double expected;
  };
  const Case cases[] = {
      {"1234567890123456", 1234567890123450.0},       {"99999999999999990000", 99999999999999900000.0},
      {"12345678901234567", 12345678901234500.0},     {"1.23456789012345678", 1.23456789012345},
      {"0.000000007123456", 0.000000007123456},  // 7 sig digits, unchanged
      {"1000000000000000000", 1000000000000000000.0},
  };
  for (const auto& c : cases) {
    Tokenizer tz(c.literal);
    const auto& v = tz.tokens();
    ASSERT_EQ(v.size(), 2u) << c.literal;  // Number + Eof
    EXPECT_EQ(v[0].kind, TokenKind::Number) << c.literal;
    EXPECT_DOUBLE_EQ(v[0].number, c.expected) << c.literal;
    EXPECT_TRUE(tz.errors().empty()) << c.literal;
  }
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

TEST(TokenizerErrorLiterals, LongestPrefixLeavesTrailingOperator) {
  // Excel rewrites operands of a deleted column into `#REF!` literals, so a
  // surviving formula reads `#REF!/2`. The scanner must commit on the
  // literal and leave the `/2` for the main loop instead of treating the
  // whole run as an invalid spelling.
  Tokenizer tz("#REF!/2");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 4u);  // ErrorLiteral, Slash, Number, Eof
  EXPECT_EQ(v[0].kind, TokenKind::ErrorLiteral);
  EXPECT_EQ(v[0].error_code, ErrorCode::Ref);
  EXPECT_EQ(v[1].kind, TokenKind::Slash);
  EXPECT_EQ(v[2].kind, TokenKind::Number);
  EXPECT_TRUE(tz.errors().empty());
}

TEST(TokenizerErrorLiterals, NaFollowedBySlashReference) {
  // `#N/A/B1`: the four-byte `#N/A` matches, leaving `/B1`. The trailing
  // `/` after the literal must not be folded into the error run.
  Tokenizer tz("#N/A/B1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 4u);  // ErrorLiteral, Slash, CellRef, Eof
  EXPECT_EQ(v[0].kind, TokenKind::ErrorLiteral);
  EXPECT_EQ(v[0].error_code, ErrorCode::NA);
  EXPECT_EQ(v[1].kind, TokenKind::Slash);
  EXPECT_EQ(v[2].kind, TokenKind::CellRef);
  EXPECT_TRUE(tz.errors().empty());
}

TEST(TokenizerErrorLiterals, DivZeroWithTrailingSlashReference) {
  // `#DIV/0!/A1`: `#DIV/0!` itself contains a `/`, so the match must span
  // the whole literal (7 bytes) and only then leave `/A1`.
  Tokenizer tz("#DIV/0!/A1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 4u);  // ErrorLiteral, Slash, CellRef, Eof
  EXPECT_EQ(v[0].kind, TokenKind::ErrorLiteral);
  EXPECT_EQ(v[0].error_code, ErrorCode::Div0);
  EXPECT_EQ(std::string(v[0].lexeme), "#DIV/0!");
  EXPECT_EQ(v[1].kind, TokenKind::Slash);
  EXPECT_EQ(v[2].kind, TokenKind::CellRef);
  EXPECT_TRUE(tz.errors().empty());
}

TEST(TokenizerErrorLiterals, AllVariantsFollowedByOperator) {
  // Every catalog literal immediately followed by `+1` must tokenize as
  // ErrorLiteral, Plus, Number, Eof.
  const char* literals[] = {
      "#NULL!", "#DIV/0!", "#VALUE!",   "#REF!",     "#NAME?",     "#NUM!",  "#N/A",     "#GETTING_DATA", "#SPILL!",
      "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#EXTERNAL!", "#BUSY!", "#PYTHON!", "#UNKNOWN!",
  };
  for (const char* lit : literals) {
    const std::string src = std::string(lit) + "+1";
    Tokenizer tz(src);
    const auto& v = tz.tokens();
    ASSERT_EQ(v.size(), 4u) << lit;
    EXPECT_EQ(v[0].kind, TokenKind::ErrorLiteral) << lit;
    EXPECT_EQ(std::string(v[0].lexeme), lit) << lit;
    EXPECT_EQ(v[1].kind, TokenKind::Plus) << lit;
    EXPECT_EQ(v[2].kind, TokenKind::Number) << lit;
    EXPECT_TRUE(tz.errors().empty()) << lit;
  }
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

TEST(TokenizerIdentifiers, BackslashStartIsExcelsThirdLegalNameCharacter) {
  // Excel's name-manager rule accepts a letter, underscore, or backslash as
  // the first character of a defined name.
  Tokenizer tz("\\foo");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Ident);
  EXPECT_EQ(std::string(v[0].lexeme), "\\foo");
  EXPECT_TRUE(tz.errors().empty());
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

TEST(TokenizerCellRefs, RepeatedAbsoluteAnchorIsInvalidReference) {
  Tokenizer tz("A$$1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Invalid);
  ASSERT_EQ(tz.errors().size(), 1u);
  EXPECT_EQ(tz.errors()[0].code, LexerErrorCode::InvalidReference);
}

TEST(TokenizerCellRefs, OverflowColumn) {
  // XFE is past the column cap; falls back to Ident.
  Tokenizer tz("XFE1");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Ident);
}

TEST(TokenizerCellRefs, ColumnOnlyBecomesIdents) {
  // Documented tokenizer limitation: A:A is IDENT COLON IDENT, left for
  // the parser to promote.
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
  // full spilled-range grammar is wired in at the parser layer.
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

TEST(TokenizerIdentBytes, ContinuationByteIsNotIdentStart) {
  // Regression: U+0080..U+00BF are UTF-8 continuation bytes and must
  // not begin an identifier. is_ident_start_byte previously returned
  // true for any byte >= 0x80, which let a stray continuation slip into
  // the identifier slot and shift all subsequent tokens off by a byte.
  // A standalone continuation byte must surface as InvalidCharacter.
  Tokenizer tz(
      "\x80"
      "abc");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidCharacter);
}

TEST(TokenizerIdentBytes, MidRangeContinuationByteIsNotIdentStart) {
  // 0xBF is the highest UTF-8 continuation byte; same expectation.
  Tokenizer tz(
      "\xBF"
      "abc");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidCharacter);
}

TEST(TokenizerIdentBytes, TruncatedLeadByteTerminates) {
  // A lead byte with no continuation byte behind it decodes to nothing, so
  // the identifier scanner consumed zero bytes and returned to a dispatcher
  // that classified the same byte as an identifier start all over again.
  // The tokenizer appended an empty token per turn and grew without bound;
  // reaching Eof at all is the assertion.
  Tokenizer tz("\xD7");
  const std::vector<Token>& toks = tz.tokens();
  ASSERT_FALSE(toks.empty());
  EXPECT_EQ(toks.back().kind, TokenKind::Eof);
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidCharacter);
}

TEST(TokenizerIdentBytes, LeadByteFollowedByNonContinuationTerminates) {
  // Same non-progress path, reached with a 4-byte lead whose second byte is
  // not a continuation.
  Tokenizer tz(
      "\xF3"
      "\x0B");
  const std::vector<Token>& toks = tz.tokens();
  ASSERT_FALSE(toks.empty());
  EXPECT_EQ(toks.back().kind, TokenKind::Eof);
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidCharacter);
}

TEST(TokenizerIdentBytes, TruncatedLeadByteDoesNotSwallowWhatFollows) {
  // Consuming exactly one byte keeps the rest of the formula tokenizable,
  // so a malformed prefix degrades the input instead of destroying it.
  Tokenizer tz(
      "\xD7"
      "A1");
  const std::vector<Token>& toks = tz.tokens();
  ASSERT_GE(toks.size(), 2U);
  EXPECT_EQ(toks.front().kind, TokenKind::CellRef);
  EXPECT_EQ(toks.front().lexeme, "A1");
}

TEST(TokenizerWhitespace, FullwidthSpaceIsInvalid) {
  // U+3000 is three UTF-8 bytes (E3 80 80) and must be flagged.
  Tokenizer tz(
      "\xE3\x80\x80"
      "A1");
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  const LexerError& err = tz.errors().front();
  EXPECT_EQ(err.code, LexerErrorCode::InvalidCharacter);
  // The fullwidth space is the very first codepoint (one UTF-16 unit), so
  // the range must start at 0, not at whatever position a stale
  // `mark_start()` from a previous token would have left behind.
  EXPECT_EQ(err.range.start, 0u);
  EXPECT_EQ(err.range.end, 1u);
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
  const LexerError& err = tz.errors().front();
  EXPECT_EQ(err.code, LexerErrorCode::InvalidCharacter);
  // "A1" occupies UTF-16 units [0,2); the BOM starts right after it.
  EXPECT_EQ(err.range.start, 2u);
  EXPECT_EQ(err.range.end, 3u);
}

TEST(TokenizerBomHandling, MidInputBomDoesNotDesyncSubsequentOffsets) {
  // Regression: a manual `byte_pos_ += 3` in the mid-input BOM branch never
  // advanced `utf16_pos_`, so every token after the BOM reported a UTF-16
  // offset one code unit behind its real position (colliding with the
  // BOM's own range instead of following it).
  Tokenizer tz(
      "A1\xEF\xBB\xBF"
      "B2");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 3u);  // CellRef("A1"), CellRef("B2"), Eof -- the BOM emits no token.
  EXPECT_EQ(v[0].kind, TokenKind::CellRef);
  EXPECT_EQ(v[0].range.start, 0u);
  EXPECT_EQ(v[0].range.end, 2u);
  EXPECT_EQ(v[1].kind, TokenKind::CellRef);
  EXPECT_EQ(v[1].range.start, 3u);
  EXPECT_EQ(v[1].range.end, 5u);
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
// ErrorRanges
// ---------------------------------------------------------------------------
// `Tokenizer::record_error` is documented to produce a range covering
// [err_start, byte_pos_) in UTF-16 code units. These probe the
// unclassifiable-byte default branch directly (the other stale-mark_start
// branches -- excessive length, mid-input BOM, fullwidth space -- are
// covered next to their own token-family tests above).

TEST(TokenizerErrorRanges, InvalidByteRangePointsAtOffendingByte) {
  // 0x01 (SOH) is not a valid formula byte anywhere. "A1+" is 3 ASCII
  // bytes / UTF-16 units, so the offending byte's range must be [3,4),
  // not the position of the preceding '+' token (the stale-`mark_start()`
  // bug pointed the squiggle one token too early).
  Tokenizer tz(
      "A1+\x01"
      "+B2");
  const auto& v = tz.tokens();
  // CellRef("A1"), Plus, [invalid byte: no token], Plus, CellRef("B2"), Eof.
  ASSERT_EQ(v.size(), 5u);
  ASSERT_FALSE(tz.errors().empty());
  const LexerError& err = tz.errors().front();
  EXPECT_EQ(err.code, LexerErrorCode::InvalidCharacter);
  EXPECT_EQ(err.range.start, 3u);
  EXPECT_EQ(err.range.end, 4u);
  // The token stream past the invalid byte must stay correctly offset.
  EXPECT_EQ(v[2].kind, TokenKind::Plus);
  EXPECT_EQ(v[2].range.start, 4u);
  EXPECT_EQ(v[3].kind, TokenKind::CellRef);
  EXPECT_EQ(v[3].range.start, 5u);
  EXPECT_EQ(v[3].range.end, 7u);
}

// ---------------------------------------------------------------------------
// ExcessiveLength
// ---------------------------------------------------------------------------

TEST(TokenizerExcessiveLength, Truncates) {
  TokenizerOptions opts;
  opts.max_formula_length_utf16 = 5;
  // Token boundaries land on the cap, so the main loop stops on it.
  Tokenizer tz("A B C D E F G H", opts);
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  const LexerError& err = tz.errors().front();
  EXPECT_EQ(err.code, LexerErrorCode::ExcessiveLength);
  // Regression: the cap branch never called `mark_start()` before jumping
  // `byte_pos_` to the end of source, so the range's start carried over
  // from whichever earlier token happened to call `mark_start()` last
  // (here, "E" at [4,5)) instead of the cap boundary itself.
  EXPECT_EQ(err.range.start, 5u);
  EXPECT_EQ(err.range.end, 5u);
}

TEST(TokenizerExcessiveLength, SingleOversizedIdentifierIsCapped) {
  TokenizerOptions opts;
  opts.max_formula_length_utf16 = 5;
  // One unbroken identifier: the scanner runs to the end of the input
  // without ever re-entering the main loop, so nothing but an up-front
  // bound can stop it at the cap.
  Tokenizer tz("ABCDEFGHIJKLMNOP", opts);
  const std::vector<Token>& toks = tz.tokens();
  ASSERT_FALSE(toks.empty());
  EXPECT_EQ(toks.front().lexeme, "ABCDE");
  EXPECT_EQ(toks.back().kind, TokenKind::Eof);
  EXPECT_EQ(toks.back().range.end, 5u);
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::ExcessiveLength);
}

TEST(TokenizerExcessiveLength, SingleOversizedStringLiteralIsCapped) {
  TokenizerOptions opts;
  opts.max_formula_length_utf16 = 5;
  // Same shape through the string scanner. Cutting the literal leaves it
  // unterminated, which is the honest reading of a formula that exceeds
  // what Excel itself accepts.
  Tokenizer tz("\"ABCDEFGHIJKLMNOP\"", opts);
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  bool saw_excessive_length = false;
  for (const LexerError& err : tz.errors()) {
    saw_excessive_length = saw_excessive_length || err.code == LexerErrorCode::ExcessiveLength;
  }
  EXPECT_TRUE(saw_excessive_length);
}

TEST(TokenizerExcessiveLength, MultibyteRunStopsNearTheCap) {
  TokenizerOptions opts;
  opts.max_formula_length_utf16 = 4;
  // Each emoji is one codepoint but two UTF-16 units, so the cap falls on
  // a surrogate pair. Cutting on the codepoint boundary keeps the token
  // stream well-formed rather than splitting a character in half.
  Tokenizer tz("\"\xF0\x9F\x98\x80\xF0\x9F\x98\x80\xF0\x9F\x98\x80\"", opts);
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  bool saw_excessive_length = false;
  for (const LexerError& err : tz.errors()) {
    saw_excessive_length = saw_excessive_length || err.code == LexerErrorCode::ExcessiveLength;
  }
  EXPECT_TRUE(saw_excessive_length);
}

TEST(TokenizerExcessiveLength, DefaultCapMatchesExcelsFormulaLengthLimit) {
  // Excel 365's real per-formula ceiling is 8,192 characters (Microsoft 365
  // Excel specifications and limits; see tests/divergence.yaml's
  // left_associative_chain_default_depth_cap entry for the directly
  // measured Excel observation). A formula one character over that limit
  // is truncated with ExcessiveLength using the *default* options, and one
  // at exactly the limit is accepted.
  {
    const std::string over(8193, 'a');
    Tokenizer tz(over);
    (void)tz.tokens();
    ASSERT_FALSE(tz.errors().empty());
    EXPECT_EQ(tz.errors().front().code, LexerErrorCode::ExcessiveLength);
  }
  {
    const std::string at_cap(8192, 'a');
    Tokenizer tz(at_cap);
    (void)tz.tokens();
    EXPECT_TRUE(tz.errors().empty());
  }
}

TEST(TokenizerExcessiveLength, InputAtExactlyTheCapIsAccepted) {
  TokenizerOptions opts;
  opts.max_formula_length_utf16 = 5;
  // The bound must not fire one unit early: an identifier of exactly the
  // cap length is a legal formula.
  Tokenizer tz("ABCDE", opts);
  const std::vector<Token>& toks = tz.tokens();
  ASSERT_GE(toks.size(), 2U);
  EXPECT_EQ(toks.front().lexeme, "ABCDE");
  EXPECT_TRUE(tz.errors().empty());
}

// ---------------------------------------------------------------------------
// Locale-aware lexing
// ---------------------------------------------------------------------------

TokenizerOptions LocaleOptions(ExcelLocale locale) {
  TokenizerOptions opts;
  opts.locale = &locale_facts(ExcelProfile{ExcelHost::kMac365, locale});
  return opts;
}

TEST(TokenizerLocale, ListSeparatorAndDecimalPoint) {
  for (const ExcelLocale locale : {ExcelLocale::kDeDE, ExcelLocale::kFrFR}) {
    Tokenizer tz("SUMME(1,5;2)", LocaleOptions(locale));
    const auto& v = tz.tokens();
    ASSERT_EQ(v.size(), 7u);
    EXPECT_EQ(v[0].kind, TokenKind::Ident);
    EXPECT_EQ(v[1].kind, TokenKind::LParen);
    EXPECT_EQ(v[2].kind, TokenKind::Number);
    EXPECT_EQ(v[2].lexeme, "1,5");
    EXPECT_DOUBLE_EQ(v[2].number, 1.5);
    EXPECT_FALSE(v[2].is_integer);
    EXPECT_EQ(v[3].kind, TokenKind::Comma);
    EXPECT_EQ(v[4].kind, TokenKind::Number);
    EXPECT_EQ(v[5].kind, TokenKind::RParen);
    EXPECT_TRUE(tz.errors().empty());
  }
}

TEST(TokenizerLocale, LeadingDecimalSeparatorStartsANumber) {
  Tokenizer tz(",5+1", LocaleOptions(ExcelLocale::kDeDE));
  const auto& v = tz.tokens();
  ASSERT_GE(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::Number);
  EXPECT_DOUBLE_EQ(v[0].number, 0.5);
}

TEST(TokenizerLocale, SecondDecimalSeparatorIsAnInvalidLiteral) {
  Tokenizer tz("1,2,3", LocaleOptions(ExcelLocale::kDeDE));
  (void)tz.tokens();
  ASSERT_FALSE(tz.errors().empty());
  EXPECT_EQ(tz.errors().front().code, LexerErrorCode::InvalidNumberLiteral);
}

TEST(TokenizerLocale, ArrayConstantSeparators) {
  for (const ExcelLocale locale : {ExcelLocale::kDeDE, ExcelLocale::kFrFR}) {
    Tokenizer tz("{1,5.2,5;3,4}", LocaleOptions(locale));
    const auto& v = tz.tokens();
    const std::vector<TokenKind> expected = {TokenKind::LBrace, TokenKind::Number,    TokenKind::Comma,
                                             TokenKind::Number, TokenKind::Semicolon, TokenKind::Number,
                                             TokenKind::RBrace, TokenKind::Eof};
    ASSERT_EQ(v.size(), expected.size());
    for (std::size_t i = 0; i < expected.size(); ++i) {
      EXPECT_EQ(v[i].kind, expected[i]) << i;
    }
    EXPECT_DOUBLE_EQ(v[1].number, 1.5);
    EXPECT_DOUBLE_EQ(v[3].number, 2.5);
    EXPECT_DOUBLE_EQ(v[5].number, 3.4);
    EXPECT_TRUE(tz.errors().empty());
  }
}

TEST(TokenizerLocale, ColumnSeparatorDoesNotJoinNamesInsideBraces) {
  Tokenizer tz("{WAHR.FALSCH}", LocaleOptions(ExcelLocale::kDeDE));
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 6u);
  EXPECT_EQ(v[1].kind, TokenKind::Bool);
  EXPECT_TRUE(v[1].boolean);
  EXPECT_EQ(v[2].kind, TokenKind::Comma);
  EXPECT_EQ(v[3].kind, TokenKind::Bool);
  EXPECT_FALSE(v[3].boolean);
}

TEST(TokenizerLocale, ListSeparatorOutsideBracesIsNotARowSeparator) {
  EXPECT_EQ(KindsOf("A1;B1", LocaleOptions(ExcelLocale::kDeDE)),
            (std::vector<TokenKind>{TokenKind::CellRef, TokenKind::Comma, TokenKind::CellRef, TokenKind::Eof}));
}

TEST(TokenizerLocale, BooleanNames) {
  Tokenizer de("WAHR falsch", LocaleOptions(ExcelLocale::kDeDE));
  const auto& d = de.tokens();
  ASSERT_GE(d.size(), 3u);
  EXPECT_EQ(d[0].kind, TokenKind::Bool);
  EXPECT_TRUE(d[0].boolean);
  EXPECT_EQ(d[2].kind, TokenKind::Bool);
  EXPECT_FALSE(d[2].boolean);

  Tokenizer fr("VRAI FAUX TRUE", LocaleOptions(ExcelLocale::kFrFR));
  const auto& f = fr.tokens();
  ASSERT_GE(f.size(), 5u);
  EXPECT_TRUE(f[0].boolean);
  EXPECT_EQ(f[2].kind, TokenKind::Bool);
  EXPECT_FALSE(f[2].boolean);
  EXPECT_EQ(f[4].kind, TokenKind::Bool);

  // Without the locale the localized spelling is a name.
  EXPECT_EQ(KindsOf("WAHR"), (std::vector<TokenKind>{TokenKind::Ident, TokenKind::Eof}));
}

TEST(TokenizerLocale, LocalizedErrorName) {
  Tokenizer tz("#NV", LocaleOptions(ExcelLocale::kDeDE));
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 2u);
  EXPECT_EQ(v[0].kind, TokenKind::ErrorLiteral);
  EXPECT_EQ(v[0].error_code, ErrorCode::NA);
  // The canonical spelling is still accepted.
  Tokenizer canonical("#N/A", LocaleOptions(ExcelLocale::kDeDE));
  EXPECT_EQ(canonical.tokens()[0].error_code, ErrorCode::NA);
  // Without the locale `#NV` is not an error literal.
  Tokenizer plain("#NV");
  (void)plain.tokens();
  EXPECT_FALSE(plain.errors().empty());
}

TEST(TokenizerLocale, QuotedTextAndStructuredRefsKeepSeparators) {
  for (const ExcelLocale locale : {ExcelLocale::kDeDE, ExcelLocale::kFrFR}) {
    Tokenizer sheet("'a;b'!A1", LocaleOptions(locale));
    const auto& s = sheet.tokens();
    ASSERT_GE(s.size(), 3u);
    EXPECT_EQ(s[0].kind, TokenKind::SheetName);
    EXPECT_EQ(s[0].text, "a;b");

    Tokenizer text("\"x;y\"", LocaleOptions(locale));
    const auto& t = text.tokens();
    ASSERT_EQ(t.size(), 2u);
    EXPECT_EQ(t[0].kind, TokenKind::String);
    EXPECT_EQ(t[0].text, "x;y");

    Tokenizer table("Tbl[[#Headers],[a;b]]", LocaleOptions(locale));
    std::size_t semicolons = 0;
    std::size_t commas = 0;
    for (const Token& token : table.tokens()) {
      semicolons += token.kind == TokenKind::Semicolon ? 1 : 0;
      commas += token.kind == TokenKind::Comma ? 1 : 0;
    }
    EXPECT_EQ(semicolons, 1u);
    EXPECT_EQ(commas, 1u);
  }
}

TEST(TokenizerLocale, InvariantFactsLexLikeTheDefaultPath) {
  const char* const sources[] = {"SUM(1.5,2;3)", "{1,2.5;3,4}", "IF(TRUE,#N/A,FALSE)", "'a,b'!A1+.5",
                                 "T[[#Headers],[a]]"};
  for (const ExcelLocale locale :
       {ExcelLocale::kEnUS, ExcelLocale::kJaJP, ExcelLocale::kZhCN, ExcelLocale::kKoKR, ExcelLocale::kThTH}) {
    for (const char* source : sources) {
      EXPECT_EQ(KindsOf(source, LocaleOptions(locale)), KindsOf(source)) << source;
    }
  }
}

}  // namespace
}  // namespace parser
}  // namespace formulon
