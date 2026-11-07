//
// Why a strict formula-entry parse failed, and the status it fails with.
//
// `ParseFailure` is the one diagnostic a public entry point reports for a
// formula that does not parse: the first diagnostic the parser recorded
// (lexer errors precede parser errors), or the entry refusal that rejected a
// formula the parser accepted. `range` is in UTF-16 code units of the text
// handed to the parse.
//
// `parse_error_status` maps a code onto the 1000-range parser statuses; a
// code without a dedicated status maps to 1000. Statuses 1002 (unterminated
// comment), 1008 (unknown function), 1011 (invalid R1C1) and 1012 (invalid
// dot notation) are intentionally never produced: an unknown function is
// `#NAME?` at evaluation, not a syntax error.

#ifndef FORMULON_PARSER_PARSE_STATUS_H_
#define FORMULON_PARSER_PARSE_STATUS_H_

#include <string_view>

#include "parser/parse_error.h"
#include "parser/token.h"
#include "utils/error.h"

namespace formulon {
namespace parser {

/// The diagnostic a strict entry parse fails with. `message` is the static
/// `default_message(code)` literal.
struct ParseFailure {
  ParseErrorCode code = ParseErrorCode::UnexpectedToken;
  TextRange range{};
  std::string_view message;
};

/// Returns the status a formula failing with `code` reports.
constexpr FormulonErrorCode parse_error_status(ParseErrorCode code) noexcept {
  switch (code) {
    case ParseErrorCode::LexerUnterminatedString:
      return FormulonErrorCode::kParserUnterminatedString;
    case ParseErrorCode::LexerInvalidNumberLiteral:
      return FormulonErrorCode::kParserInvalidNumber;
    case ParseErrorCode::LexerInvalidReference:
    case ParseErrorCode::InvalidReference:
    case ParseErrorCode::InvalidRange:
      return FormulonErrorCode::kParserInvalidReference;
    case ParseErrorCode::LexerUnterminatedSheetQuote:
      return FormulonErrorCode::kParserInvalidSheetName;
    case ParseErrorCode::UnbalancedBrackets:
    case ParseErrorCode::UnbalancedBraces:
    case ParseErrorCode::ExpectedCloseParen:
      return FormulonErrorCode::kParserMismatchedBrackets;
    case ParseErrorCode::NestedFormulaTooDeep:
      return FormulonErrorCode::kParserExcessiveNesting;
    case ParseErrorCode::ArrayRowMismatch:
    case ParseErrorCode::ExpectedCommaOrSemiInArray:
      return FormulonErrorCode::kParserInvalidArrayLiteral;
    case ParseErrorCode::StructuredRefInvalid:
      return FormulonErrorCode::kParserInvalidStructRef;
    case ParseErrorCode::EntryBomNotSupported:
      return FormulonErrorCode::kParserBomNotSupported;
    case ParseErrorCode::TooManyErrors:
      return FormulonErrorCode::kParserTooManyErrors;
    case ParseErrorCode::LexerInvalidCharacter:
    case ParseErrorCode::LexerInvalidErrorLiteral:
    case ParseErrorCode::LexerInvalidEscape:
    case ParseErrorCode::LexerExcessiveLength:
    case ParseErrorCode::UnexpectedToken:
    case ParseErrorCode::UnexpectedEof:
    case ParseErrorCode::ExpectedExpression:
    case ParseErrorCode::ExpectedRParenOrComma:
    case ParseErrorCode::UnsupportedConstruct:
    case ParseErrorCode::ExpectedOpenParen:
    case ParseErrorCode::ExpectedComma:
    case ParseErrorCode::LetInvalidName:
    case ParseErrorCode::LetWrongArity:
    case ParseErrorCode::LambdaInvalidParam:
    case ParseErrorCode::LambdaEmpty:
    case ParseErrorCode::LambdaDuplicateParam:
    case ParseErrorCode::EntryWrongArity:
    case ParseErrorCode::EntryReferenceRequired:
    case ParseErrorCode::EntryQualifiedBuiltin:
    case ParseErrorCode::EntryRefusedWhitespace:
    case ParseErrorCode::EntryParamShapedName:
    case ParseErrorCode::EntryDoubleEquals:
      return FormulonErrorCode::kParserUnexpectedToken;
  }
  return FormulonErrorCode::kParserUnexpectedToken;
}

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_PARSE_STATUS_H_
