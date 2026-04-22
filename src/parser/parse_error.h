// Copyright 2026 libraz. Licensed under the MIT License.
//
// Parser-level error model.
//
// `ParseErrorCode` consolidates lexer-level codes (promoted from
// `LexerErrorCode`) with the syntactic parser-level codes that the Pratt
// parser produces. The semantic codes (unknown function, unknown name,
// arity mismatches, structured-ref / sheet validation, ...) are still
// deferred until the FunctionRegistry and Workbook surfaces land.
//
// Messages are static English literals (one per code, no parameterisation).
// The `string_view` in `ParseError::message` therefore points at program-
// lifetime storage and the AST consumer can hold onto it freely. The
// `suggestion` slot is reserved for the upcoming Suggestion engine and is
// always left empty for now.

#ifndef FORMULON_PARSER_PARSE_ERROR_H_
#define FORMULON_PARSER_PARSE_ERROR_H_

#include <cstdint>
#include <string_view>

#include "parser/token.h"

namespace formulon {
namespace parser {

/// Error catalogue for the Pratt parser.
///
/// Codes 0..7 mirror `LexerErrorCode` so callers see a single enum after
/// promotion. Codes 8.. are parser-level. Numeric values are stable; do
/// not reorder, append-only.
enum class ParseErrorCode : std::uint16_t {
  // Lexer-level (promoted from LexerErrorCode).
  LexerInvalidCharacter = 0,
  LexerUnterminatedString,
  LexerUnterminatedSheetQuote,
  LexerInvalidNumberLiteral,
  LexerInvalidErrorLiteral,
  LexerInvalidEscape,
  LexerInvalidReference,
  LexerExcessiveLength,
  // Parser-level: token-shape errors.
  UnexpectedToken,
  UnexpectedEof,
  ExpectedExpression,
  ExpectedCloseParen,          // `=(1+2`
  UnbalancedBraces,            // `={1,2`
  ArrayRowMismatch,            // `={1,2;3}`
  ExpectedRParenOrComma,       // arg list separator missing
  ExpectedCommaOrSemiInArray,  // `={1 2}`
  InvalidReference,            // structurally invalid ref / range endpoint
  UnsupportedConstruct,        // grammar form not yet implemented
  // Parser-level: panic-mode recovery and bracket-parity errors.
  ExpectedOpenParen,     // `=SUM 1,2)` (function name without `(`)
  ExpectedComma,         // `=SUM(1 2)`
  UnbalancedBrackets,    // `=Table[col`
  InvalidRange,          // `=A1:foo` (rhs of `:` not ref-like)
  NestedFormulaTooDeep,  // recursion depth exceeded ParserOptions limit
  TooManyErrors,         // accumulated error count hit the cap
};

/// Diagnostic severity. The parser currently emits only `Error`; `Warning`
/// is reserved for future use (e.g. style hints from the Suggestion engine).
enum class Severity : std::uint8_t { Error = 0, Warning = 1 };

/// A single parser error record.
///
/// `range` is in UTF-16 code units, matching the editor position model used
/// by every other Formulon diagnostic. `message` is a static English string
/// keyed off `code`; this will be swapped for an interned, parameterised
/// message once localisation lands.
///
/// `offending_token` is a span into the original source covering the token
/// that triggered the error (empty if not applicable). `suggestion` is
/// reserved for the upcoming Suggestion engine and is currently always
/// empty.
struct ParseError {
  ParseErrorCode code = ParseErrorCode::UnexpectedToken;
  TextRange range{};
  std::string_view message;
  std::string_view offending_token;
  Severity severity = Severity::Error;
  std::string_view suggestion;
};

/// Returns the canonical English message for `code`. The pointer references a
/// static string literal with program lifetime.
const char* default_message(ParseErrorCode code) noexcept;

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_PARSE_ERROR_H_
