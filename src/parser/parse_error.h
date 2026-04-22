// Copyright 2026 libraz. Licensed under the MIT License.
//
// Parser-level error model.
//
// `ParseErrorCode` consolidates lexer-level codes (promoted from
// `LexerErrorCode`) with the small set of parser-level codes that the
// happy-path Pratt parser needs. The full 35-code catalogue, panic-mode
// recovery, and the suggestion engine are deferred to follow-up work.
//
// Messages are static English literals (one per code, no parameterisation).
// The `string_view` in `ParseError::message` therefore points at program-
// lifetime storage and the AST consumer can hold onto it freely.

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
/// promotion. Codes 8.. are parser-level. The full catalogue (35+ codes,
/// per `backup/plans/19-parser-errors.md`) will follow.
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
  // Parser-level (minimal happy-path set).
  UnexpectedToken,
  UnexpectedEof,
  ExpectedExpression,
  UnclosedParen,
  UnclosedBrace,
  ArrayRowMismatch,
  ExpectedRParenOrComma,
  ExpectedCommaOrSemiInArray,
  InvalidCellRef,
  UnsupportedConstruct,
};

/// A single parser error record.
///
/// `range` is in UTF-16 code units, matching the editor position model used
/// by every other Formulon diagnostic. `message` is a static English string
/// keyed off `code`; this will be swapped for an interned, parameterised
/// message once localisation lands.
struct ParseError {
  ParseErrorCode code = ParseErrorCode::UnexpectedToken;
  TextRange range;
  std::string_view message;
};

/// Returns the canonical English message for `code`. The pointer references a
/// static string literal with program lifetime.
const char* default_message(ParseErrorCode code) noexcept;

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_PARSE_ERROR_H_
