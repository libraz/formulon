// Copyright 2026 libraz. Licensed under the MIT License.
//
// Minimal lexer-level error record. The Pratt parser (M2.3) will merge a
// vector of `LexerError` into its larger `ParseError` list (see
// `backup/plans/19-parser-errors.md` §19.3); at M2.2 we only need enough
// structure to round-trip the code, UTF-16 range, and offending source span.

#ifndef FORMULON_PARSER_LEXER_ERROR_H_
#define FORMULON_PARSER_LEXER_ERROR_H_

#include <cstdint>
#include <string_view>

#include "parser/token.h"

namespace formulon {
namespace parser {

/// Narrow enumeration of failure modes the tokenizer may flag.
///
/// These map to a subset of the parser-level `ParseErrorCode` catalog; the
/// M2.3 parser is responsible for promoting these into fully-formed
/// `ParseError` values with localized messages.
enum class LexerErrorCode : std::uint16_t {
  UnterminatedString,
  UnterminatedSheetQuote,
  InvalidNumberLiteral,
  InvalidErrorLiteral,
  InvalidCharacter,
  InvalidEscape,
  InvalidReference,
  ExcessiveLength,
};

/// Record of a single tokenizer-level failure.
///
/// `lexeme` points into the source buffer passed to the tokenizer; it covers
/// the offending run of input. `range` is in UTF-16 code units, matching the
/// editor position model.
struct LexerError {
  LexerErrorCode code = LexerErrorCode::InvalidCharacter;
  TextRange range;
  std::string_view lexeme;
};

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_LEXER_ERROR_H_
