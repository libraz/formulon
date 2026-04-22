// Copyright 2026 libraz. Licensed under the MIT License.
//
// Pratt parser core.
//
// Consumes the token stream emitted by `Tokenizer` and produces an
// `AstNode` tree built with the parser's factory functions. Scope:
//
//   * Atoms: numeric / boolean / error literals, cell refs, sheet-qualified
//     refs, full-column / full-row refs, defined-name refs, parenthesised
//     expressions, inline array literals (`{1,2;3,4}`).
//   * Function calls (`IDENT(args, ...)`).
//   * Operators (high -> low precedence): `:` range, prefix `+`/`-`,
//     postfix `%`, `^` (right-assoc), `*` `/`, binary `+` `-`, `&`,
//     comparisons (`=`, `<>`, `<`, `<=`, `>`, `>=`), prefix `@`
//     (implicit-intersection, lowest precedence).
//
// Out of scope for the current parser (deferred to follow-up work):
//   * String literals (Value::text not yet implemented).
//   * Union (`,`) outside call arglists; intersection-as-space.
//   * External / structured / lambda / let / immediately-invoked-lambda forms.
//   * Suggestion engine (the `ParseError::suggestion` slot is reserved but
//     never populated yet).
//
// Recovery: panic-mode. On any error inside an expression the parser
// records a diagnostic, skips tokens up to the next
// matching synchronisation point (closing paren, comma, semicolon, etc.),
// and substitutes an `ErrorPlaceholder` node for the unparseable subtree.
// Subsequent siblings (next call argument, next array element, next infix
// term) are still parsed, so a single formula can yield multiple errors.

#ifndef FORMULON_PARSER_PARSER_H_
#define FORMULON_PARSER_PARSER_H_

#include <cstdint>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/lexer_error.h"
#include "parser/parse_error.h"
#include "parser/token.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {

/// Tunable parser limits.
///
/// `max_error_count` caps the size of the error list before the parser
/// short-circuits with a final `TooManyErrors` sentinel and stops doing
/// useful work. `max_parse_depth` bounds how deeply nested an expression
/// may recurse before we emit `NestedFormulaTooDeep` and bail to the
/// nearest sync point; this is the parser's stack-overflow guard.
struct ParserOptions {
  /// Maximum number of `ParseError` records to accumulate. When the limit is
  /// hit the offending error is recorded normally and a sentinel
  /// `TooManyErrors` entry is appended; further parsing is skipped.
  std::uint32_t max_error_count = 50;
  /// Maximum recursion depth of `parse_expression`. Excel-side formulas in
  /// the wild rarely exceed double-digit depths; 128 leaves a comfortable
  /// margin while still preventing pathological inputs from blowing the
  /// native stack.
  std::uint32_t max_parse_depth = 128;
};

/// Single-shot Pratt parser.
///
/// `source` and `arena` must outlive the parser and any `AstNode` it returns.
/// The parser drives a fresh `Tokenizer` internally; callers do not need to
/// pre-tokenise.
class Parser {
 public:
  /// Builds a parser over `source`. The AST is allocated in `arena`.
  Parser(std::string_view source, Arena& arena, ParserOptions opts = {}) noexcept;

  Parser(const Parser&) = delete;
  Parser& operator=(const Parser&) = delete;
  Parser(Parser&&) = delete;
  Parser& operator=(Parser&&) = delete;

  /// Parses the source. Returns the root expression node.
  ///
  /// With panic-mode recovery enabled the root may be non-null even when
  /// `errors()` is non-empty: the parser substitutes `ErrorPlaceholder`
  /// nodes for unparseable subtrees and keeps going. The result is `nullptr`
  /// only when the very first atom could not be produced (e.g. empty input,
  /// arena exhaustion). Subsequent calls return the cached result without
  /// re-parsing.
  AstNode* parse();

  /// Returns the accumulated parse / lex errors. Valid after `parse()`.
  const std::vector<ParseError>& errors() const noexcept { return errors_; }

 private:
  // Synchronisation context for panic-mode recovery. See parser.cpp for
  // the per-context skip-set definitions.
  enum class SyncContext : std::uint8_t {
    TopLevel = 0,   // sync to Eof only
    Paren = 1,      // sync to RParen
    CallArg = 2,    // sync to Comma or RParen
    ArrayElem = 3,  // sync to Comma, Semicolon, or RBrace
  };

  // Returns the kind of the token at `pos_`, or Eof if past end.
  TokenKind peek_kind() const noexcept;
  // Returns the kind of the token at `pos_ + offset` (still ignoring
  // whitespace because `tokens_` is the pre-filtered stream).
  TokenKind peek_kind_at(std::size_t offset) const noexcept;
  // Returns a const reference to the current token (or the synthesised Eof).
  const Token& peek() const noexcept;
  // Returns a const reference to a future token.
  const Token& peek_at(std::size_t offset) const noexcept;
  // Consumes and returns the current token.
  const Token& advance() noexcept;

  // Records an error at `range` with the given code. Honours the
  // `max_error_count` cap by appending a `TooManyErrors` sentinel and
  // setting `bailed_` once the threshold is reached.
  void record_error(ParseErrorCode code, TextRange range);
  // Same, but also captures the `offending_token` source span.
  void record_error_with_token(ParseErrorCode code, TextRange range, std::string_view offending);

  // Skips tokens until a sync point matching `ctx` is reached (or Eof).
  // Respects nested `(` `[` `{` so commas/semicolons inside an inner group
  // do not satisfy an outer sync. Does not consume the sync token itself.
  void skip_to_sync(SyncContext ctx) noexcept;

  // Pratt expression parser. `min_bp` is the minimum binding power required
  // to consume a left-denotation. `ctx` is the recovery context used when an
  // error occurs inside this subexpression. Always returns a non-null node:
  // an unparseable subtree becomes an `ErrorPlaceholder`.
  AstNode* parse_expression(int min_bp, SyncContext ctx);

  // Atom / null-denotation dispatch.
  AstNode* parse_atom(SyncContext ctx);

  // Per-atom helpers.
  AstNode* parse_number_atom();
  AstNode* parse_bool_atom();
  AstNode* parse_error_literal_atom();
  AstNode* parse_paren_atom();
  AstNode* parse_array_literal_atom();
  AstNode* parse_at_prefix_atom(SyncContext ctx);
  AstNode* parse_unary_prefix_atom(UnaryOp op, SyncContext ctx);
  AstNode* parse_ident_or_call_or_full_col();
  AstNode* parse_sheet_qualified_ref(std::string_view sheet, bool quoted, TextRange sheet_range);
  AstNode* parse_cellref_atom();
  AstNode* parse_full_row_or_number(const Token& first);

  // Builds a Reference from a CellRef token's lexeme.
  bool decode_cellref_lexeme(std::string_view lex, Reference* out) noexcept;
  // Returns the column index encoded by an Ident token's letters; 0 means
  // not a valid column-letter run.
  static std::uint32_t decode_column_letters(std::string_view lex, bool* col_abs) noexcept;

  // Promotes any tokenizer-level errors into `errors_`.
  void promote_lexer_errors(const std::vector<LexerError>& lex_errors);

  // Source + token storage.
  std::string_view source_;
  Arena& arena_;
  ParserOptions opts_;

  // Filtered token stream (whitespace removed).
  std::vector<Token> tokens_;
  // Always-present synthetic Eof returned by peek/advance past end.
  Token sentinel_eof_{};

  std::size_t pos_ = 0;
  std::vector<ParseError> errors_;
  AstNode* root_ = nullptr;
  bool parsed_ = false;
  // Latched when the error-count cap is reached or recovery cannot even
  // synthesise a placeholder; once set, no further useful work happens and
  // every parse_expression call returns an `ErrorPlaceholder` immediately.
  bool bailed_ = false;
  // Current recursion depth of `parse_expression`, used to enforce
  // `ParserOptions::max_parse_depth`.
  std::uint32_t depth_ = 0;
};

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_PARSER_H_
