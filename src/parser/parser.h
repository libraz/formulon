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
// Out of scope for the current happy-path parser (deferred to follow-up work):
//   * String literals (Value::text not yet implemented).
//   * Union (`,`) outside call arglists; intersection-as-space.
//   * External / structured / lambda / let / immediately-invoked-lambda forms.
//   * Panic-mode recovery and the full ParseErrorCode catalogue.
//
// The parser is single-shot: `parse()` caches the root and subsequent calls
// return the same pointer. Errors are accumulated into a vector reachable
// via `errors()`. On a hard error the parser bails out and returns null.

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

/// Tunable parser limits. `max_error_count` is currently ignored; it will
/// be used to short-circuit panic-mode recovery once enough errors have
/// accumulated.
struct ParserOptions {
  /// Maximum number of `ParseError` records to accumulate before bailing.
  /// Reserved for the forthcoming panic-mode recovery loop.
  std::uint32_t max_error_count = 100;
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

  /// Parses the source. Returns the root expression node, or `nullptr` on a
  /// hard parse failure (in which case `errors()` carries the diagnostic).
  /// Subsequent calls return the cached result without re-parsing.
  AstNode* parse();

  /// Returns the accumulated parse / lex errors. Valid after `parse()`.
  const std::vector<ParseError>& errors() const noexcept { return errors_; }

 private:
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

  // Records an error at `range` with the given code and stops parsing.
  void record_error(ParseErrorCode code, TextRange range);

  // Pratt expression parser. `min_bp` is the minimum binding power required
  // to consume a left-denotation. Returns nullptr on error.
  AstNode* parse_expression(int min_bp);

  // Atom / null-denotation dispatch.
  AstNode* parse_atom();

  // Per-atom helpers.
  AstNode* parse_number_atom();
  AstNode* parse_bool_atom();
  AstNode* parse_error_literal_atom();
  AstNode* parse_paren_atom();
  AstNode* parse_array_literal_atom();
  AstNode* parse_at_prefix_atom();
  AstNode* parse_unary_prefix_atom(UnaryOp op);
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
  // Set when `parse()` has produced or attempted to produce a result with
  // an unrecoverable error; further parsing stops.
  bool bailed_ = false;
};

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_PARSER_H_
