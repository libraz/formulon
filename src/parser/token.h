// Copyright 2026 libraz. Licensed under the MIT License.
//
// Token representation.
//
// This header declares `TokenKind`, `TextRange`, and `Token`: the three
// types that the Pratt parser consumes from the tokenizer. The token
// catalog follows `backup/plans/02-calc-engine.md` §2.2; offsets use
// UTF-16 code units so diagnostics drop directly into Monaco / CodeMirror 6
// editors (see `backup/plans/19-parser-errors.md` §19.3).
//
// Token lifetime: `lexeme` and `text` are `string_view`s into storage owned
// by either the original source buffer (most tokens) or the `Arena` held by
// the `Tokenizer` that produced them (escape-expanded strings and sheet
// names). The tokenizer must outlive any code that reads from its tokens.

#ifndef FORMULON_PARSER_TOKEN_H_
#define FORMULON_PARSER_TOKEN_H_

#include <cstdint>
#include <string_view>

#include "value.h"

namespace formulon {
namespace parser {

/// Token kind enumeration.
///
/// Values are used as dispatch keys by the Pratt parser; add new kinds at
/// the end of the relevant group rather than reordering to keep ABI stable.
enum class TokenKind : std::uint8_t {
  // Literals.
  Number,
  String,
  Bool,
  ErrorLiteral,
  // Names.
  Ident,
  SheetName,  // The unquoted / escape-resolved content of a 'Sheet 1' ref.
  // References.
  CellRef,
  // Structural punctuation.
  LParen,
  RParen,
  LBrace,
  RBrace,
  LBracket,
  RBracket,
  Comma,
  Semicolon,
  Colon,
  Bang,
  // Operators.
  Plus,
  Minus,
  Star,
  Slash,
  Caret,
  Percent,
  Ampersand,
  Eq,
  NotEq,
  Lt,
  LtEq,
  Gt,
  GtEq,
  At,
  Hash,
  // Whitespace / end.
  Whitespace,
  Eof,
  // Lexer failure marker. `lexeme` covers the offending run so the parser
  // can keep advancing without re-scanning.
  Invalid,
};

/// Source-range descriptor in UTF-16 code units.
///
/// `start` and `end` are UTF-16 code unit offsets from the beginning of the
/// source buffer passed to the tokenizer. `line` is 1-based; `column` is
/// 1-based in UTF-16 code units within the line. Supplementary-plane code
/// points (emoji, rare CJK) contribute 2 UTF-16 code units but only one
/// grapheme of visual width - this matches the Monaco Editor position model.
struct TextRange {
  std::uint32_t start = 0;
  std::uint32_t end = 0;
  std::uint32_t line = 1;
  std::uint32_t column = 1;
};

/// A single token emitted by the `Tokenizer`.
///
/// Only the payload slot matching `kind` is meaningful:
///   - `Number`       -> `number`, `is_integer`.
///   - `Bool`         -> `boolean`.
///   - `String`       -> `text` (escape-resolved, arena-backed).
///   - `SheetName`    -> `text` (escape-resolved, arena-backed).
///   - `ErrorLiteral` -> `error_code`.
/// All other kinds leave the payload slots at their default values.
struct Token {
  TokenKind kind = TokenKind::Invalid;
  TextRange range;
  /// Verbatim source bytes spanned by this token. Always points into the
  /// source buffer passed to the tokenizer (never the arena).
  std::string_view lexeme;
  /// Parsed numeric payload (Number).
  double number = 0.0;
  /// True if the Number token lacked a decimal point / exponent.
  bool is_integer = false;
  /// Parsed boolean payload (Bool).
  bool boolean = false;
  /// Escape-resolved text. Backed by the tokenizer arena for String /
  /// SheetName; empty for other kinds.
  std::string_view text;
  /// Parsed error-literal code (ErrorLiteral).
  ErrorCode error_code = ErrorCode::Null;
};

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_TOKEN_H_
