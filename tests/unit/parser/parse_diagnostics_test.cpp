// Parse diagnostics: entry-path failures carry a ParseErrorCode and a UTF-16
// range, and map onto the 1000-range status codes.

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/builtin_names.h"
#include "gtest/gtest.h"
#include "parser/parse_error.h"
#include "parser/parse_status.h"
#include "parser/parser.h"
#include "parser/tokenizer.h"
#include "utils/arena.h"
#include "utils/error.h"

namespace formulon {
namespace parser {
namespace {

// Parses `src` through the diagnostics entry and returns the failure; fails
// the test when the parse succeeds.
ParseFailure FailureOf(std::string_view src) {
  Arena arena;
  ParseFailure failure;
  EXPECT_EQ(parse_entry_with_diagnostics(src, arena, &failure), nullptr) << src;
  return failure;
}

TEST(ParseDiagnostics, AppendedCodesKeepTheirValues) {
  EXPECT_EQ(static_cast<int>(ParseErrorCode::LambdaDuplicateParam), 28);
  EXPECT_EQ(static_cast<int>(ParseErrorCode::EntryWrongArity), 29);
  EXPECT_EQ(static_cast<int>(ParseErrorCode::EntryReferenceRequired), 30);
  EXPECT_EQ(static_cast<int>(ParseErrorCode::EntryQualifiedBuiltin), 31);
  EXPECT_EQ(static_cast<int>(ParseErrorCode::EntryRefusedWhitespace), 32);
  EXPECT_EQ(static_cast<int>(ParseErrorCode::EntryParamShapedName), 33);
  EXPECT_EQ(static_cast<int>(ParseErrorCode::EntryDoubleEquals), 34);
  EXPECT_EQ(static_cast<int>(ParseErrorCode::EntryBomNotSupported), 35);
  EXPECT_EQ(static_cast<int>(ParseErrorCode::StructuredRefInvalid), 36);
}

TEST(ParseDiagnostics, EveryCodeMapsToItsStatus) {
  const std::vector<std::pair<ParseErrorCode, int>> table = {
      {ParseErrorCode::LexerInvalidCharacter, 1000},
      {ParseErrorCode::LexerUnterminatedString, 1001},
      {ParseErrorCode::LexerUnterminatedSheetQuote, 1005},
      {ParseErrorCode::LexerInvalidNumberLiteral, 1003},
      {ParseErrorCode::LexerInvalidErrorLiteral, 1000},
      {ParseErrorCode::LexerInvalidEscape, 1000},
      {ParseErrorCode::LexerInvalidReference, 1004},
      {ParseErrorCode::LexerExcessiveLength, 1000},
      {ParseErrorCode::UnexpectedToken, 1000},
      {ParseErrorCode::UnexpectedEof, 1000},
      {ParseErrorCode::ExpectedExpression, 1000},
      {ParseErrorCode::ExpectedCloseParen, 1006},
      {ParseErrorCode::UnbalancedBraces, 1006},
      {ParseErrorCode::ArrayRowMismatch, 1009},
      {ParseErrorCode::ExpectedRParenOrComma, 1000},
      {ParseErrorCode::ExpectedCommaOrSemiInArray, 1009},
      {ParseErrorCode::InvalidReference, 1004},
      {ParseErrorCode::UnsupportedConstruct, 1000},
      {ParseErrorCode::ExpectedOpenParen, 1000},
      {ParseErrorCode::ExpectedComma, 1000},
      {ParseErrorCode::UnbalancedBrackets, 1006},
      {ParseErrorCode::InvalidRange, 1004},
      {ParseErrorCode::NestedFormulaTooDeep, 1007},
      {ParseErrorCode::TooManyErrors, 1014},
      {ParseErrorCode::LetInvalidName, 1000},
      {ParseErrorCode::LetWrongArity, 1000},
      {ParseErrorCode::LambdaInvalidParam, 1000},
      {ParseErrorCode::LambdaEmpty, 1000},
      {ParseErrorCode::LambdaDuplicateParam, 1000},
      {ParseErrorCode::EntryWrongArity, 1000},
      {ParseErrorCode::EntryReferenceRequired, 1000},
      {ParseErrorCode::EntryQualifiedBuiltin, 1000},
      {ParseErrorCode::EntryRefusedWhitespace, 1000},
      {ParseErrorCode::EntryParamShapedName, 1000},
      {ParseErrorCode::EntryDoubleEquals, 1000},
      {ParseErrorCode::EntryBomNotSupported, 1013},
      {ParseErrorCode::StructuredRefInvalid, 1010},
  };
  ASSERT_EQ(table.size(), 37U);
  for (std::size_t i = 0; i < table.size(); ++i) {
    EXPECT_EQ(static_cast<std::size_t>(table[i].first), i);
    EXPECT_EQ(static_cast<int>(parse_error_status(table[i].first)), table[i].second) << "code " << i;
  }
}

TEST(ParseDiagnostics, UnusedStatusesAreNeverProduced) {
  for (int code = 0; code <= 36; ++code) {
    const int status = static_cast<int>(parse_error_status(static_cast<ParseErrorCode>(code)));
    EXPECT_GE(status, 1000);
    EXPECT_LE(status, 1014);
    EXPECT_NE(status, 1002) << "code " << code;
    EXPECT_NE(status, 1008) << "code " << code;
    EXPECT_NE(status, 1011) << "code " << code;
    EXPECT_NE(status, 1012) << "code " << code;
  }
}

TEST(ParseDiagnostics, CleanParseReturnsRootAndLeavesFailureUntouched) {
  Arena arena;
  ParseFailure failure;
  failure.code = ParseErrorCode::TooManyErrors;
  EXPECT_NE(parse_entry_with_diagnostics("SUM(1,2)+A1", arena, &failure), nullptr);
  EXPECT_EQ(failure.code, ParseErrorCode::TooManyErrors);
  EXPECT_TRUE(failure.message.empty());
}

TEST(ParseDiagnostics, AgreesWithParseStrict) {
  const std::vector<std::string> sources = {
      "1+2", "SUM(A1:B2)", "(1+2", "{1,2;3}", "\"abc", "Table[col", "1 2", "LET(x,1)", "Sheet1!SUM(1)", "#",
  };
  for (const std::string& src : sources) {
    Arena strict_arena;
    Arena diag_arena;
    ParseFailure failure;
    const bool strict_ok = parse_strict(src, strict_arena) != nullptr;
    const bool diag_ok = parse_entry_with_diagnostics(src, diag_arena, &failure) != nullptr;
    EXPECT_EQ(strict_ok, diag_ok) << src;
  }
}

TEST(ParseDiagnostics, EmptyFormulaIsExpectedExpressionAtZero) {
  const ParseFailure failure = FailureOf("");
  EXPECT_EQ(failure.code, ParseErrorCode::ExpectedExpression);
  EXPECT_EQ(failure.range.start, 0U);
  EXPECT_EQ(failure.range.end, 0U);
  EXPECT_EQ(failure.message, default_message(ParseErrorCode::ExpectedExpression));
}

TEST(ParseDiagnostics, LoneOpenParenFailsAtTheEnd) {
  const ParseFailure failure = FailureOf("(");
  EXPECT_EQ(failure.range.start, 1U);
  EXPECT_EQ(failure.range.end, 1U);
  EXPECT_EQ(parse_error_status(failure.code), FormulonErrorCode::kParserUnexpectedToken);
}

TEST(ParseDiagnostics, LexerErrorIsReportedFirst) {
  const ParseFailure failure = FailureOf("SUM(1,\"abc");
  EXPECT_EQ(failure.code, ParseErrorCode::LexerUnterminatedString);
  EXPECT_EQ(failure.range.start, 6U);
  EXPECT_EQ(parse_error_status(failure.code), FormulonErrorCode::kParserUnterminatedString);
}

TEST(ParseDiagnostics, ShapeFailuresMapToTheirStatus) {
  EXPECT_EQ(parse_error_status(FailureOf("(1+2").code), FormulonErrorCode::kParserMismatchedBrackets);
  EXPECT_EQ(parse_error_status(FailureOf("Table[col").code), FormulonErrorCode::kParserMismatchedBrackets);
  EXPECT_EQ(parse_error_status(FailureOf("{1,2;3}").code), FormulonErrorCode::kParserInvalidArrayLiteral);
}

TEST(ParseDiagnostics, RangeCountsSupplementaryCodePointsAsTwoUnits) {
  // U+1F600 is four UTF-8 bytes and two UTF-16 units: the quoted literal
  // spans units 0..4 and the stray `1` sits at unit 4.
  const ParseFailure failure = FailureOf("\"\xF0\x9F\x98\x80\"1");
  EXPECT_EQ(failure.code, ParseErrorCode::UnexpectedToken);
  EXPECT_EQ(failure.range.start, 4U);
  EXPECT_EQ(failure.range.end, 5U);
}

TEST(ParseDiagnostics, OverlongFormulaIsExcessiveLength) {
  const std::uint32_t limit = TokenizerOptions{}.max_formula_length_utf16;
  std::string src = "1";
  while (src.size() <= limit) {
    src += "+1";
  }
  EXPECT_EQ(FailureOf(src).code, ParseErrorCode::LexerExcessiveLength);
}

TEST(ParseDiagnostics, QualifiedBuiltinIsRefusedOverItsCallee) {
  Arena arena;
  ParseFailure failure;
  EXPECT_NE(parse_entry_with_diagnostics("Sheet1!SUM(1)", arena, &failure), nullptr);
  EXPECT_EQ(eval::parse_formula_entry_with_diagnostics("Sheet1!SUM(1)", arena, &failure), nullptr);
  EXPECT_EQ(failure.code, ParseErrorCode::EntryQualifiedBuiltin);
  EXPECT_EQ(failure.range.start, 0U);
  EXPECT_EQ(failure.range.end, 10U);
  EXPECT_EQ(failure.message, default_message(ParseErrorCode::EntryQualifiedBuiltin));
}

TEST(ParseDiagnostics, QualifiedDefinedNameCallIsAccepted) {
  Arena arena;
  ParseFailure failure;
  EXPECT_NE(eval::parse_formula_entry_with_diagnostics("Sheet1!MyFn(1)", arena, &failure), nullptr);
}

TEST(ParseDiagnostics, FormulaEntryAgreesWithParseFormulaEntry) {
  const std::vector<std::string> sources = {"1+2", "Sheet1!SUM(1)", "Sheet1!MyFn(1)", "(1", ""};
  for (const std::string& src : sources) {
    Arena plain_arena;
    Arena diag_arena;
    ParseFailure failure;
    const bool plain_ok = eval::parse_formula_entry(src, plain_arena) != nullptr;
    const bool diag_ok = eval::parse_formula_entry_with_diagnostics(src, diag_arena, &failure) != nullptr;
    EXPECT_EQ(plain_ok, diag_ok) << src;
  }
}

TEST(ParseDiagnostics, EntryFailureIsSyntaxErrorBeforeQualifiedBuiltin) {
  Arena arena;
  ParseFailure failure;
  EXPECT_EQ(eval::parse_formula_entry_with_diagnostics("Sheet1!SUM(1", arena, &failure), nullptr);
  EXPECT_EQ(failure.code, ParseErrorCode::ExpectedCloseParen);
}

}  // namespace
}  // namespace parser
}  // namespace formulon
