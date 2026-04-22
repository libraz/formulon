// Copyright 2026 libraz. Licensed under the MIT License.
//
// Excel formula tokenizer.
//
// The tokenizer consumes a UTF-8 formula source and emits a vector of
// `Token` values terminated by a single `Eof` token. Errors do not halt
// tokenization: each failure is recorded in `errors()` and the offending
// run is either skipped or emitted as an `Invalid` token so the downstream
// Pratt parser can still produce partial ASTs for diagnostic purposes.
//
// Design notes (see `backup/plans/02-calc-engine.md` §2.2 and
// `backup/plans/19-parser-errors.md` §19.4.3):
//
//   * Whitespace is *not* collapsed: the parser needs exact runs to decide
//     whether a space is the intersection operator or ignorable indent.
//   * Offsets are in UTF-16 code units so diagnostics are drop-in for
//     Monaco Editor / CodeMirror 6.
//   * External book refs (`[Book1.xlsx]Sheet1!A1`, `[1]Sheet1!A1`) and
//     structured refs (`Table[@col]`) are *not* promoted to dedicated token
//     kinds here; the lexer emits the component punctuation tokens and the
//     parser reinterprets them. This keeps the lexer small and lets the
//     parser handle context-sensitive rules.
//   * Column-only (`A:A`) and row-only (`1:1`) references: the lexer emits
//     them as `Ident COLON Ident` and `Number COLON Number` respectively;
//     the parser promotes the adjacent tokens to full range references.

#ifndef FORMULON_PARSER_TOKENIZER_H_
#define FORMULON_PARSER_TOKENIZER_H_

#include <cstdint>
#include <string_view>
#include <vector>

#include "parser/lexer_error.h"
#include "parser/token.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {

/// Tokenizer configuration knobs. The defaults match the per-formula limits
/// enforced by Excel 365.
struct TokenizerOptions {
  /// Hard cap on input length measured in UTF-16 code units. Inputs longer
  /// than this are truncated and an `ExcessiveLength` error is recorded.
  std::uint32_t max_formula_length_utf16 = 32768;
};

/// Converts a UTF-8 formula source into a token stream.
///
/// Single-shot: call `tokens()` once to drive tokenization; subsequent calls
/// return the cached stream. `errors()` may be queried at any time after
/// `tokens()` returns.
class Tokenizer {
 public:
  /// Builds a tokenizer over `source`. `source` must outlive every read of
  /// `tokens()` / `errors()`: token `lexeme` views point into it.
  explicit Tokenizer(std::string_view source, TokenizerOptions opts = {}) noexcept;

  Tokenizer(const Tokenizer&) = delete;
  Tokenizer& operator=(const Tokenizer&) = delete;
  Tokenizer(Tokenizer&&) = delete;
  Tokenizer& operator=(Tokenizer&&) = delete;

  /// Tokenizes the source lazily on first call and returns the cached
  /// stream. The returned vector always ends with a single `Eof` token.
  const std::vector<Token>& tokens();

  /// Returns the accumulated lexer-level errors. Valid after `tokens()`.
  const std::vector<LexerError>& errors() const noexcept { return errors_; }

 private:
  // Per-codepoint decode result from the UTF-8 cursor.
  struct CodepointInfo {
    std::uint32_t codepoint = 0;
    std::uint32_t byte_len = 0;
    std::uint32_t utf16_units = 0;
    bool valid = false;
  };

  // Decodes the next UTF-8 codepoint at byte offset `i` into `source_`.
  // Returns `{valid=false, byte_len=1}` for malformed bytes so the caller
  // can record an error and advance by one byte.
  CodepointInfo peek_codepoint(std::size_t i) const noexcept;

  // Advances `byte_pos_` / UTF-16 offset / line / column bookkeeping by one
  // codepoint starting at `byte_pos_`.
  void advance_one();

  // Snapshots the current position as the `start` of a new token.
  void mark_start() noexcept;

  // Builds a TextRange covering [start_*, current]. `lexeme_from` is the
  // byte offset at which the token began.
  TextRange make_range() const noexcept;

  // Emits a token of `kind`, with `lexeme` covering [lex_start, byte_pos_).
  void emit(TokenKind kind, std::size_t lex_start);

  // Emits a LexerError covering [err_start, byte_pos_).
  void record_error(LexerErrorCode code, std::size_t err_start);

  // Per-kind scanners. Each assumes `byte_pos_` points at the first byte of
  // a token matching its expected prefix.
  void scan_whitespace();
  void scan_string();
  void scan_quoted_sheet_name();
  void scan_number();
  void scan_error_literal();
  void scan_ident_or_cellref_or_bool();
  void scan_lt();
  void scan_gt();

  // Checks whether `word` (ASCII) is `TRUE` or `FALSE` case-insensitively.
  static bool is_bool_word(std::string_view word, bool* out) noexcept;

  // Classifies `run` against the 17-strong error-literal catalog.
  static bool match_error_literal(std::string_view run, ErrorCode* out) noexcept;

  // Classifies an identifier run as an A1 cell reference. Returns true iff
  // `run` has the shape `\$?[A-Za-z]{1,3}\$?[0-9]{1,7}` and the column /
  // row numeric values are within Excel 365 limits. When the letter run is
  // syntactically valid but the numeric run is missing, sets `*letters_only`
  // so the caller can emit an `Ident` instead.
  static bool looks_like_cellref(std::string_view run, bool* letters_only) noexcept;

  // Internal helpers ------------------------------------------------------
  static bool is_ascii_letter(char c) noexcept;
  static bool is_ascii_digit(char c) noexcept;
  static bool is_ident_start_byte(unsigned char c) noexcept;
  static bool is_ident_cont_byte(unsigned char c) noexcept;
  static bool is_formula_whitespace(char c) noexcept;

  // Source + position state.
  std::string_view source_;
  TokenizerOptions opts_;
  std::size_t byte_pos_ = 0;
  std::uint32_t utf16_pos_ = 0;
  std::uint32_t line_ = 1;
  std::uint32_t column_ = 1;

  // Token-start snapshot (filled by `mark_start`).
  std::uint32_t start_utf16_ = 0;
  std::uint32_t start_line_ = 1;
  std::uint32_t start_column_ = 1;

  // Output buffers.
  std::vector<Token> tokens_;
  std::vector<LexerError> errors_;
  Arena arena_;
  bool done_ = false;
  bool truncated_ = false;

  // The byte offset at which the most recent CellRef ends. Used to
  // disambiguate a trailing `#` into the spilled-range operator (OP_HASH).
  std::size_t last_cellref_end_byte_ = static_cast<std::size_t>(-1);
};

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_TOKENIZER_H_
