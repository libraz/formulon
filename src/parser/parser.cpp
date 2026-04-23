// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the Pratt parser. Driven by a single `Tokenizer` pass;
// whitespace tokens are stripped during ingest. A leading `=` (Excel
// formula prefix) is consumed if it sits at the very start.
//
// Operator binding-power scheme (higher = tighter):
//
//   Range `:`             80   (left-assoc binary)
//   Prefix unary `+`/`-`  70
//   Postfix `%`           60
//   Power `^`             50   (right-assoc binary)
//   `*` `/`               40
//   Binary `+` `-`        30
//   Concat `&`            20
//   Comparisons           10
//   Prefix `@`             1   (lowest - consumes the entire RHS)
//
// Right-associativity is implemented by recursing the RHS at the same
// binding power (`min_bp = bp`), left-associativity at `min_bp = bp + 1`.
//
// Error recovery: every `parse_expression` call accepts a `SyncContext`
// describing the syntactic frame it lives in. On error we
// record a diagnostic, skip tokens up to the next sync token (paying
// attention to nested `(` `[` `{`), and substitute an `ErrorPlaceholder`
// for the failed subtree so siblings continue to parse. The parser stops
// entirely once the error-count cap is reached or the recursion-depth
// limit is exceeded.

#include "parser/parser.h"

#include <cstdint>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/lexer_error.h"
#include "parser/parse_error.h"
#include "parser/reference.h"
#include "parser/token.h"
#include "parser/tokenizer.h"
#include "utils/arena.h"
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace parser {

namespace {

// Binding-power constants. See file header for the precedence table.
constexpr int kBpRange = 80;
constexpr int kBpUnaryPrefix = 70;
constexpr int kBpPostfixPercent = 60;
constexpr int kBpPow = 50;
constexpr int kBpMulDiv = 40;
constexpr int kBpAddSub = 30;
constexpr int kBpConcat = 20;
constexpr int kBpComparison = 10;
constexpr int kBpAtPrefix = 1;

constexpr std::uint32_t kMaxColumn = 16384;  // XFD
constexpr std::uint32_t kMaxRow = 1048576;   // 2^20

// ASCII helpers. Re-implemented locally to avoid depending on the tokenizer's
// privates and to keep `parser.cpp` self-contained.
bool IsAsciiLetter(char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

bool IsAsciiDigit(char c) noexcept {
  return c >= '0' && c <= '9';
}

// Returns the binary binding-power for a token, or 0 if it is not a binary
// operator at the parser level. `right_bp` receives the precedence to use for
// the recursive RHS parse (left-assoc -> bp+1, right-assoc -> bp).
int InfixBindingPower(TokenKind kind, int* right_bp) noexcept {
  switch (kind) {
    case TokenKind::Colon:
      *right_bp = kBpRange + 1;
      return kBpRange;
    case TokenKind::Caret:
      *right_bp = kBpPow;  // right-assoc.
      return kBpPow;
    case TokenKind::Star:
    case TokenKind::Slash:
      *right_bp = kBpMulDiv + 1;
      return kBpMulDiv;
    case TokenKind::Plus:
    case TokenKind::Minus:
      *right_bp = kBpAddSub + 1;
      return kBpAddSub;
    case TokenKind::Ampersand:
      *right_bp = kBpConcat + 1;
      return kBpConcat;
    case TokenKind::Eq:
    case TokenKind::NotEq:
    case TokenKind::Lt:
    case TokenKind::LtEq:
    case TokenKind::Gt:
    case TokenKind::GtEq:
      *right_bp = kBpComparison + 1;
      return kBpComparison;
    default:
      return 0;
  }
}

// Maps a binary token kind to its `BinOp` enum value.
BinOp TokenToBinOp(TokenKind kind) noexcept {
  switch (kind) {
    case TokenKind::Plus:
      return BinOp::Add;
    case TokenKind::Minus:
      return BinOp::Sub;
    case TokenKind::Star:
      return BinOp::Mul;
    case TokenKind::Slash:
      return BinOp::Div;
    case TokenKind::Caret:
      return BinOp::Pow;
    case TokenKind::Ampersand:
      return BinOp::Concat;
    case TokenKind::Eq:
      return BinOp::Eq;
    case TokenKind::NotEq:
      return BinOp::NotEq;
    case TokenKind::Lt:
      return BinOp::Lt;
    case TokenKind::LtEq:
      return BinOp::LtEq;
    case TokenKind::Gt:
      return BinOp::Gt;
    case TokenKind::GtEq:
      return BinOp::GtEq;
    default:
      return BinOp::Add;  // unreachable: caller checks InfixBindingPower first.
  }
}

ParseErrorCode PromoteLexerCode(LexerErrorCode lc) noexcept {
  switch (lc) {
    case LexerErrorCode::InvalidCharacter:
      return ParseErrorCode::LexerInvalidCharacter;
    case LexerErrorCode::UnterminatedString:
      return ParseErrorCode::LexerUnterminatedString;
    case LexerErrorCode::UnterminatedSheetQuote:
      return ParseErrorCode::LexerUnterminatedSheetQuote;
    case LexerErrorCode::InvalidNumberLiteral:
      return ParseErrorCode::LexerInvalidNumberLiteral;
    case LexerErrorCode::InvalidErrorLiteral:
      return ParseErrorCode::LexerInvalidErrorLiteral;
    case LexerErrorCode::InvalidEscape:
      return ParseErrorCode::LexerInvalidEscape;
    case LexerErrorCode::InvalidReference:
      return ParseErrorCode::LexerInvalidReference;
    case LexerErrorCode::ExcessiveLength:
      return ParseErrorCode::LexerExcessiveLength;
  }
  return ParseErrorCode::LexerInvalidCharacter;
}

// Builds a TextRange that spans from `a.start` (using a's line/column) to
// `b.end`. Used to attach a source span to a node assembled from children.
TextRange SpanRange(TextRange a, TextRange b) noexcept {
  TextRange r;
  r.start = a.start;
  r.end = b.end;
  r.line = a.line;
  r.column = a.column;
  return r;
}

}  // namespace

const char* default_message(ParseErrorCode code) noexcept {
  switch (code) {
    case ParseErrorCode::LexerInvalidCharacter:
      return "invalid character";
    case ParseErrorCode::LexerUnterminatedString:
      return "unterminated string literal";
    case ParseErrorCode::LexerUnterminatedSheetQuote:
      return "unterminated quoted sheet name";
    case ParseErrorCode::LexerInvalidNumberLiteral:
      return "invalid number literal";
    case ParseErrorCode::LexerInvalidErrorLiteral:
      return "invalid error literal";
    case ParseErrorCode::LexerInvalidEscape:
      return "invalid escape sequence";
    case ParseErrorCode::LexerInvalidReference:
      return "invalid reference";
    case ParseErrorCode::LexerExcessiveLength:
      return "formula exceeds maximum length";
    case ParseErrorCode::UnexpectedToken:
      return "unexpected token";
    case ParseErrorCode::UnexpectedEof:
      return "unexpected end of input";
    case ParseErrorCode::ExpectedExpression:
      return "expected expression";
    case ParseErrorCode::ExpectedCloseParen:
      return "expected ')'";
    case ParseErrorCode::UnbalancedBraces:
      return "unbalanced braces in array literal";
    case ParseErrorCode::ArrayRowMismatch:
      return "array literal rows have inconsistent column counts";
    case ParseErrorCode::ExpectedRParenOrComma:
      return "expected ')' or ','";
    case ParseErrorCode::ExpectedCommaOrSemiInArray:
      return "expected ',' or ';' in array literal";
    case ParseErrorCode::InvalidReference:
      return "invalid reference";
    case ParseErrorCode::UnsupportedConstruct:
      return "construct not yet supported";
    case ParseErrorCode::ExpectedOpenParen:
      return "expected '('";
    case ParseErrorCode::ExpectedComma:
      return "expected ','";
    case ParseErrorCode::UnbalancedBrackets:
      return "unbalanced brackets in structured reference";
    case ParseErrorCode::InvalidRange:
      return "invalid range expression";
    case ParseErrorCode::NestedFormulaTooDeep:
      return "formula nesting depth exceeds the configured limit";
    case ParseErrorCode::TooManyErrors:
      return "too many parse errors; stopping";
    case ParseErrorCode::LetInvalidName:
      return "invalid LET binding name";
    case ParseErrorCode::LetWrongArity:
      return "LET requires an odd number of arguments (name, expr, [name, expr, ...], body)";
  }
  return "parse error";
}

// ---------------------------------------------------------------------------
// Construction
// ---------------------------------------------------------------------------

Parser::Parser(std::string_view source, Arena& arena, ParserOptions opts) noexcept
    : source_(source), arena_(arena), opts_(opts) {
  sentinel_eof_.kind = TokenKind::Eof;
}

// ---------------------------------------------------------------------------
// Token cursor helpers
// ---------------------------------------------------------------------------

TokenKind Parser::peek_kind() const noexcept {
  return peek().kind;
}

TokenKind Parser::peek_kind_at(std::size_t offset) const noexcept {
  return peek_at(offset).kind;
}

const Token& Parser::peek() const noexcept {
  return peek_at(0);
}

const Token& Parser::peek_at(std::size_t offset) const noexcept {
  if (pos_ + offset >= tokens_.size()) {
    return sentinel_eof_;
  }
  return tokens_[pos_ + offset];
}

const Token& Parser::advance() noexcept {
  const Token& t = peek();
  if (pos_ < tokens_.size()) {
    ++pos_;
  }
  return t;
}

void Parser::record_error(ParseErrorCode code, TextRange range) {
  record_error_with_token(code, range, std::string_view{});
}

void Parser::record_error_with_token(ParseErrorCode code, TextRange range, std::string_view offending) {
  if (bailed_) {
    return;
  }
  // Append the original error first so the caller still sees the actual code
  // even when the cap is reached on this entry.
  ParseError e;
  e.code = code;
  e.range = range;
  e.message = std::string_view(default_message(code));
  e.offending_token = offending;
  e.severity = Severity::Error;
  e.suggestion = std::string_view{};
  errors_.push_back(e);

  // Cap check: once we are at or above the budget, append the sentinel and
  // latch `bailed_`. The budget counts the sentinel against the limit so
  // `errors_.size() <= max_error_count` is preserved.
  if (opts_.max_error_count > 0 && errors_.size() + 1 >= opts_.max_error_count) {
    ParseError sentinel;
    sentinel.code = ParseErrorCode::TooManyErrors;
    sentinel.range = range;
    sentinel.message = std::string_view(default_message(ParseErrorCode::TooManyErrors));
    sentinel.severity = Severity::Error;
    errors_.push_back(sentinel);
    bailed_ = true;
  }
}

void Parser::promote_lexer_errors(const std::vector<LexerError>& lex_errors) {
  for (const auto& le : lex_errors) {
    if (bailed_) {
      return;
    }
    record_error(PromoteLexerCode(le.code), le.range);
  }
}

void Parser::skip_to_sync(SyncContext ctx) noexcept {
  // Track nesting so that, e.g., `SUM(BAD(1,2,3), 4)` does not sync on the
  // inner commas while we are recovering inside the outer call's first arg.
  std::uint32_t depth = 0;
  while (true) {
    const TokenKind k = peek_kind();
    if (k == TokenKind::Eof) {
      return;
    }
    if (depth == 0) {
      switch (ctx) {
        case SyncContext::TopLevel:
          break;  // only Eof terminates.
        case SyncContext::Paren:
          if (k == TokenKind::RParen) {
            return;
          }
          break;
        case SyncContext::CallArg:
          if (k == TokenKind::Comma || k == TokenKind::RParen) {
            return;
          }
          break;
        case SyncContext::ArrayElem:
          if (k == TokenKind::Comma || k == TokenKind::Semicolon || k == TokenKind::RBrace) {
            return;
          }
          break;
      }
    }
    // Update nesting before consuming.
    if (k == TokenKind::LParen || k == TokenKind::LBrace || k == TokenKind::LBracket) {
      ++depth;
    } else if (k == TokenKind::RParen || k == TokenKind::RBrace || k == TokenKind::RBracket) {
      if (depth > 0) {
        --depth;
      }
    }
    advance();
  }
}

// ---------------------------------------------------------------------------
// Cell-ref decoding
// ---------------------------------------------------------------------------

bool Parser::decode_cellref_lexeme(std::string_view lex, Reference* out) noexcept {
  // Accepted shapes (validated by the tokenizer): `\$?[A-Za-z]{1,3}\$?[0-9]{1,7}`.
  std::size_t i = 0;
  bool col_abs = false;
  bool row_abs = false;
  if (i < lex.size() && lex[i] == '$') {
    col_abs = true;
    ++i;
  }
  const std::size_t letters_begin = i;
  while (i < lex.size() && IsAsciiLetter(lex[i])) {
    ++i;
  }
  const std::size_t letters_len = i - letters_begin;
  if (letters_len == 0 || letters_len > 3) {
    return false;
  }
  if (i < lex.size() && lex[i] == '$') {
    row_abs = true;
    ++i;
  }
  const std::size_t digits_begin = i;
  while (i < lex.size() && IsAsciiDigit(lex[i])) {
    ++i;
  }
  const std::size_t digits_len = i - digits_begin;
  if (digits_len == 0 || digits_len > 7 || i != lex.size()) {
    return false;
  }
  // Decode column letters.
  std::uint32_t col_value = 0;
  for (std::size_t k = 0; k < letters_len; ++k) {
    char ch = lex[letters_begin + k];
    if (ch >= 'a' && ch <= 'z') {
      ch = static_cast<char>(ch - ('a' - 'A'));
    }
    col_value = col_value * 26u + static_cast<std::uint32_t>(ch - 'A' + 1);
    if (col_value > kMaxColumn) {
      return false;
    }
  }
  // Decode row digits.
  std::uint64_t row_value = 0;
  for (std::size_t k = 0; k < digits_len; ++k) {
    row_value = row_value * 10u + static_cast<std::uint32_t>(lex[digits_begin + k] - '0');
    if (row_value > kMaxRow) {
      return false;
    }
  }
  if (row_value == 0) {
    return false;
  }
  out->col = col_value - 1;
  out->row = static_cast<std::uint32_t>(row_value - 1);
  out->col_abs = col_abs;
  out->row_abs = row_abs;
  out->is_full_col = false;
  out->is_full_row = false;
  return true;
}

std::uint32_t Parser::decode_column_letters(std::string_view lex, bool* col_abs) noexcept {
  *col_abs = false;
  std::size_t i = 0;
  if (i < lex.size() && lex[i] == '$') {
    *col_abs = true;
    ++i;
  }
  const std::size_t letters_begin = i;
  while (i < lex.size() && IsAsciiLetter(lex[i])) {
    ++i;
  }
  if (i != lex.size() || (i - letters_begin) == 0 || (i - letters_begin) > 3) {
    return 0;
  }
  std::uint32_t v = 0;
  for (std::size_t k = letters_begin; k < i; ++k) {
    char ch = lex[k];
    if (ch >= 'a' && ch <= 'z') {
      ch = static_cast<char>(ch - ('a' - 'A'));
    }
    v = v * 26u + static_cast<std::uint32_t>(ch - 'A' + 1);
    if (v > kMaxColumn) {
      return 0;
    }
  }
  return v;
}

// ---------------------------------------------------------------------------
// Top-level parse
// ---------------------------------------------------------------------------

AstNode* Parser::parse() {
  if (parsed_) {
    return root_;
  }
  parsed_ = true;

  // Drive the tokenizer and pre-filter whitespace into `tokens_`.
  Tokenizer tz(source_);
  const auto& raw = tz.tokens();
  tokens_.reserve(raw.size());
  for (const auto& t : raw) {
    if (t.kind == TokenKind::Whitespace) {
      continue;
    }
    tokens_.push_back(t);
  }

  // Disambiguate CellRef-shaped tokens that are actually function-call names.
  // `LOG10`, `LOG2`, and similar function names match the cell-reference
  // pattern `[A-Z]+[0-9]+`, so the tokenizer (which is grammar-agnostic)
  // emits CellRef. When such a token is immediately followed by `(`, the
  // only correct interpretation is a function call: rewrite to Ident here.
  //
  // Skip the rewrite when the CellRef is preceded by `!` (sheet-qualified
  // reference). `Sheet1!LOG10` must remain a CellRef so that
  // `parse_sheet_qualified_ref` resolves it as the sheet-scoped cell.
  for (std::size_t i = 0; i + 1 < tokens_.size(); ++i) {
    if (tokens_[i].kind != TokenKind::CellRef)
      continue;
    if (tokens_[i + 1].kind != TokenKind::LParen)
      continue;
    if (i > 0 && tokens_[i - 1].kind == TokenKind::Bang)
      continue;
    tokens_[i].kind = TokenKind::Ident;
  }

  // Lift any tokenizer-level errors before parsing so callers see them even
  // if the parse otherwise succeeds. Promotion does not bail the parser.
  promote_lexer_errors(tz.errors());

  // Optional Excel formula prefix `=`. Only stripped at the very front and
  // only if the very first token is `Eq` (so `=A1=B1` keeps the inner `=`).
  if (!tokens_.empty() && tokens_.front().kind == TokenKind::Eq) {
    pos_ = 1;
  }

  if (peek_kind() == TokenKind::Eof) {
    record_error(ParseErrorCode::UnexpectedEof, peek().range);
    return nullptr;
  }

  root_ = parse_expression(0, SyncContext::TopLevel);
  if (root_ == nullptr) {
    return nullptr;
  }

  // Trailing tokens that we did not consume are an error; we surface the
  // first such token as UnexpectedToken (recovery here just records and
  // skips so the user sees a single trailing diagnostic, not a cascade).
  // `LBracket` and `Hash` get their own diagnostic because the user almost
  // certainly meant a structured / spilled-range form.
  if (peek_kind() != TokenKind::Eof && !bailed_) {
    const Token& tok = peek();
    if (tok.kind == TokenKind::LBracket) {
      // Reuse parse_atom's balance check by routing through it; the result
      // is an `ErrorPlaceholder` whose payload we discard - we already have
      // a root.
      (void)parse_atom(SyncContext::TopLevel);
    } else if (tok.kind == TokenKind::Hash) {
      record_error_with_token(ParseErrorCode::UnsupportedConstruct, tok.range, tok.lexeme);
      skip_to_sync(SyncContext::TopLevel);
    } else {
      record_error_with_token(ParseErrorCode::UnexpectedToken, tok.range, tok.lexeme);
      skip_to_sync(SyncContext::TopLevel);
    }
  }
  return root_;
}

// ---------------------------------------------------------------------------
// Pratt expression loop
// ---------------------------------------------------------------------------

AstNode* Parser::parse_expression(int min_bp, SyncContext ctx) {
  // Depth guard. We use a manual increment / decrement around every return
  // path because we have no exceptions and forgetting `--depth_` would be
  // easy to introduce on a future edit.
  ++depth_;
  if (depth_ > opts_.max_parse_depth) {
    record_error(ParseErrorCode::NestedFormulaTooDeep, peek().range);
    skip_to_sync(ctx);
    AstNode* placeholder = make_error_placeholder(arena_);
    if (placeholder != nullptr) {
      placeholder->set_range(peek().range);
    }
    --depth_;
    return placeholder;
  }
  if (bailed_) {
    AstNode* placeholder = make_error_placeholder(arena_);
    if (placeholder != nullptr) {
      placeholder->set_range(peek().range);
    }
    --depth_;
    return placeholder;
  }

  AstNode* lhs = parse_atom(ctx);
  if (lhs == nullptr) {
    // parse_atom only returns nullptr on arena exhaustion; treat as a hard
    // bail by latching and returning nullptr up the stack.
    bailed_ = true;
    --depth_;
    return nullptr;
  }

  while (true) {
    if (bailed_) {
      --depth_;
      return lhs;
    }
    const TokenKind kind = peek_kind();

    // Postfix operators first (currently only `%`).
    if (kind == TokenKind::Percent) {
      if (kBpPostfixPercent < min_bp) {
        --depth_;
        return lhs;
      }
      const Token& tok = advance();
      AstNode* node = make_unary_op(arena_, UnaryOp::Percent, lhs);
      if (node == nullptr) {
        bailed_ = true;
        --depth_;
        return lhs;
      }
      node->set_range(SpanRange(lhs->range(), tok.range));
      lhs = node;
      continue;
    }

    int right_bp = 0;
    const int bp = InfixBindingPower(kind, &right_bp);
    if (bp == 0 || bp < min_bp) {
      --depth_;
      return lhs;
    }

    const Token& op_tok = advance();  // consume the operator token.
    AstNode* rhs = parse_expression(right_bp, ctx);
    if (rhs == nullptr) {
      // Parse_expression always returns a node post-recovery; null means
      // hard failure (arena exhaustion). Surface the existing LHS to the
      // caller so they can salvage a partial tree.
      --depth_;
      return lhs;
    }

    AstNode* node = nullptr;
    if (kind == TokenKind::Colon) {
      // Validate that `rhs` is something that can plausibly close a range.
      // A literal / call / unary-op on the rhs of `:` is not an Excel range.
      const NodeKind rk = rhs->kind();
      if (rk != NodeKind::Ref && rk != NodeKind::NameRef && rk != NodeKind::ExternalRef &&
          rk != NodeKind::StructuredRef && rk != NodeKind::RangeOp && rk != NodeKind::ErrorPlaceholder) {
        record_error_with_token(ParseErrorCode::InvalidRange, op_tok.range, op_tok.lexeme);
      }
      node = make_range_op(arena_, lhs, rhs);
    } else {
      node = make_binary_op(arena_, TokenToBinOp(kind), lhs, rhs);
    }
    if (node == nullptr) {
      bailed_ = true;
      --depth_;
      return lhs;
    }
    node->set_range(SpanRange(lhs->range(), rhs->range()));
    lhs = node;
  }
}

// ---------------------------------------------------------------------------
// Atom dispatch
// ---------------------------------------------------------------------------

AstNode* Parser::parse_atom(SyncContext ctx) {
  const TokenKind kind = peek_kind();
  switch (kind) {
    case TokenKind::Number:
      return parse_number_atom();
    case TokenKind::Bool:
      return parse_bool_atom();
    case TokenKind::ErrorLiteral:
      return parse_error_literal_atom();
    case TokenKind::CellRef:
      return parse_cellref_atom();
    case TokenKind::LParen:
      return parse_paren_atom();
    case TokenKind::LBrace:
      return parse_array_literal_atom();
    case TokenKind::At:
      return parse_at_prefix_atom(ctx);
    case TokenKind::Plus:
      return parse_unary_prefix_atom(UnaryOp::Plus, ctx);
    case TokenKind::Minus:
      return parse_unary_prefix_atom(UnaryOp::Minus, ctx);
    case TokenKind::Ident:
      return parse_ident_or_call_or_full_col();
    case TokenKind::SheetName: {
      const Token& sheet = advance();
      AstNode* n = parse_sheet_qualified_ref(sheet.text, /*quoted=*/true, sheet.range);
      if (n != nullptr) {
        return n;
      }
      skip_to_sync(ctx);
      AstNode* placeholder = make_error_placeholder(arena_);
      if (placeholder != nullptr) {
        placeholder->set_range(sheet.range);
      }
      return placeholder;
    }
    case TokenKind::LBracket: {
      // Structured-ref grammar is not implemented yet; pick the right code
      // depending on whether the bracket is balanced (UnsupportedConstruct
      // for `=Table[col]`, UnbalancedBrackets for `=Table[col`).
      const Token& lbracket = peek();
      bool found_close = false;
      std::uint32_t depth = 1;
      for (std::size_t i = 1;; ++i) {
        const TokenKind k = peek_kind_at(i);
        if (k == TokenKind::Eof) {
          break;
        }
        if (k == TokenKind::LBracket) {
          ++depth;
        } else if (k == TokenKind::RBracket) {
          --depth;
          if (depth == 0) {
            found_close = true;
            break;
          }
        }
      }
      const ParseErrorCode code =
          found_close ? ParseErrorCode::UnsupportedConstruct : ParseErrorCode::UnbalancedBrackets;
      record_error_with_token(code, lbracket.range, lbracket.lexeme);
      skip_to_sync(ctx);
      AstNode* placeholder = make_error_placeholder(arena_);
      if (placeholder != nullptr) {
        placeholder->set_range(lbracket.range);
      }
      return placeholder;
    }
    case TokenKind::String:
      return parse_string_atom();
    case TokenKind::Hash: {
      // The spilled-range `#` is not implemented yet. Surface a single,
      // unmistakable diagnostic and recover.
      const Token& tok = peek();
      record_error_with_token(ParseErrorCode::UnsupportedConstruct, tok.range, tok.lexeme);
      skip_to_sync(ctx);
      AstNode* placeholder = make_error_placeholder(arena_);
      if (placeholder != nullptr) {
        placeholder->set_range(tok.range);
      }
      return placeholder;
    }
    case TokenKind::Invalid: {
      // The tokenizer already recorded a LexerError for this; promote to a
      // parser-level UnexpectedToken so the caller sees one definite stop.
      const Token& tok = peek();
      record_error_with_token(ParseErrorCode::UnexpectedToken, tok.range, tok.lexeme);
      advance();
      skip_to_sync(ctx);
      AstNode* placeholder = make_error_placeholder(arena_);
      if (placeholder != nullptr) {
        placeholder->set_range(tok.range);
      }
      return placeholder;
    }
    case TokenKind::Eof: {
      record_error(ParseErrorCode::UnexpectedEof, peek().range);
      AstNode* placeholder = make_error_placeholder(arena_);
      if (placeholder != nullptr) {
        placeholder->set_range(peek().range);
      }
      return placeholder;
    }
    default: {
      const Token& tok = peek();
      record_error_with_token(ParseErrorCode::ExpectedExpression, tok.range, tok.lexeme);
      skip_to_sync(ctx);
      AstNode* placeholder = make_error_placeholder(arena_);
      if (placeholder != nullptr) {
        placeholder->set_range(tok.range);
      }
      return placeholder;
    }
  }
}

// ---------------------------------------------------------------------------
// Per-atom helpers
// ---------------------------------------------------------------------------

AstNode* Parser::parse_number_atom() {
  const Token& tok = peek();
  // Full-row promotion: `Number Colon Number` -> ref like `1:1`.
  if (peek_kind_at(1) == TokenKind::Colon && peek_kind_at(2) == TokenKind::Number) {
    return parse_full_row_or_number(tok);
  }
  advance();
  AstNode* n = make_literal(arena_, Value::number(tok.number));
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(tok.range);
  return n;
}

AstNode* Parser::parse_full_row_or_number(const Token& first) {
  // first is the leading Number, currently at pos_. Validate that both ends
  // are pure digit lexemes (no decimal / exponent) and decode the row index.
  const Token& second = peek_at(2);
  auto is_pure_digit_run = [](std::string_view lex) {
    if (lex.empty()) {
      return false;
    }
    for (char c : lex) {
      if (!IsAsciiDigit(c)) {
        return false;
      }
    }
    return true;
  };
  if (!first.is_integer || !second.is_integer || !is_pure_digit_run(first.lexeme) ||
      !is_pure_digit_run(second.lexeme)) {
    // Not a clean full-row form; fall back to a plain number literal and let
    // the binary `:` rule glue it to whatever follows.
    advance();
    AstNode* n = make_literal(arena_, Value::number(first.number));
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(first.range);
    return n;
  }
  std::uint64_t lhs_row = 0;
  for (char c : first.lexeme) {
    lhs_row = lhs_row * 10u + static_cast<std::uint32_t>(c - '0');
  }
  std::uint64_t rhs_row = 0;
  for (char c : second.lexeme) {
    rhs_row = rhs_row * 10u + static_cast<std::uint32_t>(c - '0');
  }
  if (lhs_row == 0 || rhs_row == 0 || lhs_row > kMaxRow || rhs_row > kMaxRow || lhs_row != rhs_row) {
    // Not a uniform full-row form (`1:2` would be a range of single-cell
    // refs); fall back so the binary `:` rule handles it later.
    advance();
    AstNode* n = make_literal(arena_, Value::number(first.number));
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(first.range);
    return n;
  }
  // Consume Number, Colon, Number.
  advance();
  advance();
  advance();
  Reference r;
  r.row = static_cast<std::uint32_t>(lhs_row - 1);
  r.is_full_row = true;
  AstNode* n = make_ref(arena_, r);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(first.range, second.range));
  return n;
}

AstNode* Parser::parse_bool_atom() {
  const Token& tok = advance();
  AstNode* n = make_literal(arena_, Value::boolean(tok.boolean));
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(tok.range);
  return n;
}

AstNode* Parser::parse_string_atom() {
  // The token's `text` field is the escape-resolved payload, already interned
  // in the tokenizer arena, so we can hand it straight to `Value::text`
  // without copying.
  const Token& tok = advance();
  AstNode* n = make_literal(arena_, Value::text(tok.text));
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(tok.range);
  return n;
}

AstNode* Parser::parse_error_literal_atom() {
  const Token& tok = advance();
  AstNode* n = make_error_literal(arena_, tok.error_code);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(tok.range);
  return n;
}

AstNode* Parser::parse_paren_atom() {
  const Token& lparen = advance();
  AstNode* inner = parse_expression(0, SyncContext::Paren);
  if (inner == nullptr) {
    // Hard arena failure during recovery; propagate.
    return nullptr;
  }
  if (peek_kind() != TokenKind::RParen) {
    record_error_with_token(ParseErrorCode::ExpectedCloseParen, lparen.range, lparen.lexeme);
    // `inner` is still a usable subtree (possibly a placeholder). Return it
    // so the surrounding context can keep going; do not consume EOF.
    return inner;
  }
  const Token& rparen = advance();
  inner->set_range(SpanRange(lparen.range, rparen.range));
  return inner;
}

AstNode* Parser::parse_array_literal_atom() {
  const Token& lbrace = advance();
  std::vector<const AstNode*> elements;  // row-major.
  std::uint32_t cols = 0;
  std::uint32_t cur_row_cols = 0;
  std::uint32_t rows = 1;

  // Consume one element. Excel inline arrays only allow scalar literals
  // (number / bool / error). We deliberately do NOT recurse into the full
  // expression grammar.
  auto parse_element = [&]() -> AstNode* {
    const TokenKind k = peek_kind();
    AstNode* node = nullptr;
    // Allow a unary minus / plus in front of a number literal (Excel does).
    UnaryOp prefix_op = UnaryOp::Plus;
    bool have_prefix = false;
    if (k == TokenKind::Minus) {
      have_prefix = true;
      prefix_op = UnaryOp::Minus;
      advance();
    } else if (k == TokenKind::Plus) {
      have_prefix = true;
      prefix_op = UnaryOp::Plus;
      advance();
    }
    const TokenKind k2 = peek_kind();
    switch (k2) {
      case TokenKind::Number: {
        const Token& nt = advance();
        const double signed_val = (have_prefix && prefix_op == UnaryOp::Minus) ? -nt.number : nt.number;
        node = make_literal(arena_, Value::number(signed_val));
        if (node != nullptr) {
          node->set_range(nt.range);
        }
        return node;
      }
      case TokenKind::Bool: {
        const Token& bt = advance();
        node = make_literal(arena_, Value::boolean(bt.boolean));
        if (node != nullptr) {
          node->set_range(bt.range);
        }
        return node;
      }
      case TokenKind::ErrorLiteral: {
        const Token& et = advance();
        node = make_error_literal(arena_, et.error_code);
        if (node != nullptr) {
          node->set_range(et.range);
        }
        return node;
      }
      default: {
        const Token& tok = peek();
        record_error_with_token(ParseErrorCode::ExpectedExpression, tok.range, tok.lexeme);
        skip_to_sync(SyncContext::ArrayElem);
        AstNode* placeholder = make_error_placeholder(arena_);
        if (placeholder != nullptr) {
          placeholder->set_range(tok.range);
        }
        return placeholder;
      }
    }
  };

  AstNode* first = parse_element();
  if (first == nullptr) {
    return nullptr;
  }
  elements.push_back(first);
  cur_row_cols = 1;

  while (true) {
    if (bailed_) {
      break;
    }
    const TokenKind k = peek_kind();
    if (k == TokenKind::RBrace) {
      break;
    }
    if (k == TokenKind::Comma) {
      advance();
      AstNode* el = parse_element();
      if (el == nullptr) {
        return nullptr;
      }
      elements.push_back(el);
      ++cur_row_cols;
      continue;
    }
    if (k == TokenKind::Semicolon) {
      // End of a row. Establish or check the column count.
      if (cols == 0) {
        cols = cur_row_cols;
      } else if (cur_row_cols != cols) {
        // A jagged array cannot be reified as a 2D AstNode. Record the
        // diagnostic, skip to the closing `}` (consuming it), and return
        // a placeholder so siblings continue to parse.
        record_error(ParseErrorCode::ArrayRowMismatch, peek().range);
        skip_to_sync(SyncContext::ArrayElem);
        if (peek_kind() == TokenKind::RBrace) {
          advance();
        }
        AstNode* placeholder = make_error_placeholder(arena_);
        if (placeholder != nullptr) {
          placeholder->set_range(lbrace.range);
        }
        return placeholder;
      }
      cur_row_cols = 0;
      ++rows;
      advance();
      AstNode* el = parse_element();
      if (el == nullptr) {
        return nullptr;
      }
      elements.push_back(el);
      ++cur_row_cols;
      continue;
    }
    if (k == TokenKind::Eof) {
      record_error_with_token(ParseErrorCode::UnbalancedBraces, lbrace.range, lbrace.lexeme);
      // Synthesise a placeholder to carry the brace's source range.
      AstNode* placeholder = make_error_placeholder(arena_);
      if (placeholder != nullptr) {
        placeholder->set_range(lbrace.range);
      }
      return placeholder;
    }
    record_error_with_token(ParseErrorCode::ExpectedCommaOrSemiInArray, peek().range, peek().lexeme);
    skip_to_sync(SyncContext::ArrayElem);
  }

  // Final row's column count must match (or initialise) the array width.
  if (cols == 0) {
    cols = cur_row_cols;
  } else if (cur_row_cols != cols) {
    record_error(ParseErrorCode::ArrayRowMismatch, peek().range);
    if (peek_kind() == TokenKind::RBrace) {
      advance();
    }
    AstNode* placeholder = make_error_placeholder(arena_);
    if (placeholder != nullptr) {
      placeholder->set_range(lbrace.range);
    }
    return placeholder;
  }

  if (peek_kind() != TokenKind::RBrace) {
    record_error_with_token(ParseErrorCode::UnbalancedBraces, lbrace.range, lbrace.lexeme);
    AstNode* placeholder = make_error_placeholder(arena_);
    if (placeholder != nullptr) {
      placeholder->set_range(lbrace.range);
    }
    return placeholder;
  }
  const Token& rbrace = advance();

  FM_CHECK(elements.size() == static_cast<std::size_t>(rows) * cols, "array element count mismatch after parse");
  AstNode* n = make_array_literal(arena_, rows, cols, elements.data());
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(lbrace.range, rbrace.range));
  return n;
}

AstNode* Parser::parse_at_prefix_atom(SyncContext ctx) {
  const Token& at_tok = advance();
  // `@` consumes the entire remaining expression (lowest precedence).
  AstNode* operand = parse_expression(kBpAtPrefix, ctx);
  if (operand == nullptr) {
    return nullptr;
  }
  AstNode* n = make_implicit_intersection(arena_, operand);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(at_tok.range, operand->range()));
  return n;
}

AstNode* Parser::parse_unary_prefix_atom(UnaryOp op, SyncContext ctx) {
  const Token& sign_tok = advance();
  AstNode* operand = parse_expression(kBpUnaryPrefix, ctx);
  if (operand == nullptr) {
    return nullptr;
  }
  AstNode* n = make_unary_op(arena_, op, operand);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(sign_tok.range, operand->range()));
  return n;
}

AstNode* Parser::parse_cellref_atom() {
  const Token& tok = advance();
  Reference r;
  if (!decode_cellref_lexeme(tok.lexeme, &r)) {
    record_error_with_token(ParseErrorCode::InvalidReference, tok.range, tok.lexeme);
    AstNode* placeholder = make_error_placeholder(arena_);
    if (placeholder != nullptr) {
      placeholder->set_range(tok.range);
    }
    return placeholder;
  }
  AstNode* n = make_ref(arena_, r);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(tok.range);
  return n;
}

AstNode* Parser::parse_ident_or_call_or_full_col() {
  // Ident followed by:
  //   `(`   -> function call (preserving Ident lexeme as the function name).
  //   `!`   -> unquoted sheet qualifier (the Ident is the sheet name).
  //   `:`   -> potential full-column reference, iff both sides are valid
  //            column-letter runs of equal absolute index.
  //   else  -> NameRef.
  const Token& ident = peek();
  const TokenKind next = peek_kind_at(1);

  if (next == TokenKind::LParen) {
    // LET is a special form: binding names are bare identifiers that must
    // NOT be resolved as cell references. Detect it before the generic
    // argument loop kicks in so name slots go through a dedicated parse path.
    if (strings::case_insensitive_eq(ident.lexeme, std::string_view("LET"))) {
      return parse_let_call(ident);
    }
    const std::string_view fn_name = ident.lexeme;
    const TextRange call_start = ident.range;
    advance();  // Ident
    advance();  // LParen
    std::vector<const AstNode*> args;
    if (peek_kind() != TokenKind::RParen) {
      while (true) {
        if (bailed_) {
          break;
        }
        AstNode* arg = parse_expression(0, SyncContext::CallArg);
        if (arg == nullptr) {
          return nullptr;
        }
        args.push_back(arg);
        if (peek_kind() == TokenKind::Comma) {
          advance();
          continue;
        }
        if (peek_kind() == TokenKind::RParen) {
          break;
        }
        if (peek_kind() == TokenKind::Eof) {
          record_error_with_token(ParseErrorCode::ExpectedCloseParen, call_start, fn_name);
          break;
        }
        // Unexpected token between args - record and try to recover.
        record_error_with_token(ParseErrorCode::ExpectedComma, peek().range, peek().lexeme);
        skip_to_sync(SyncContext::CallArg);
        if (peek_kind() == TokenKind::Comma) {
          advance();
          continue;
        }
        if (peek_kind() == TokenKind::RParen || peek_kind() == TokenKind::Eof) {
          break;
        }
      }
    }
    if (peek_kind() != TokenKind::RParen) {
      // Recovery already handled the diagnostic for the EOF case; only emit
      // here if we somehow stopped on a non-RParen non-EOF token.
      if (peek_kind() != TokenKind::Eof) {
        record_error_with_token(ParseErrorCode::ExpectedCloseParen, call_start, fn_name);
      }
      AstNode* n =
          make_call(arena_, fn_name, args.empty() ? nullptr : args.data(), static_cast<std::uint32_t>(args.size()));
      if (n != nullptr) {
        n->set_range(call_start);
      }
      return n;
    }
    const Token& rparen = advance();
    AstNode* n =
        make_call(arena_, fn_name, args.empty() ? nullptr : args.data(), static_cast<std::uint32_t>(args.size()));
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(SpanRange(call_start, rparen.range));
    return n;
  }

  if (next == TokenKind::Bang) {
    const std::string_view sheet = ident.lexeme;
    const TextRange sheet_range = ident.range;
    advance();  // Ident
    AstNode* n = parse_sheet_qualified_ref(sheet, /*quoted=*/false, sheet_range);
    if (n != nullptr) {
      return n;
    }
    AstNode* placeholder = make_error_placeholder(arena_);
    if (placeholder != nullptr) {
      placeholder->set_range(sheet_range);
    }
    return placeholder;
  }

  if (next == TokenKind::Colon && peek_kind_at(2) == TokenKind::Ident) {
    bool lhs_abs = false;
    bool rhs_abs = false;
    const std::uint32_t lhs_col = decode_column_letters(ident.lexeme, &lhs_abs);
    const std::uint32_t rhs_col = decode_column_letters(peek_at(2).lexeme, &rhs_abs);
    if (lhs_col != 0 && rhs_col != 0 && lhs_col == rhs_col) {
      const TextRange start_range = ident.range;
      const TextRange end_range = peek_at(2).range;
      advance();  // Ident
      advance();  // Colon
      advance();  // Ident
      Reference r;
      r.col = lhs_col - 1;
      r.col_abs = lhs_abs;
      r.is_full_col = true;
      AstNode* n = make_ref(arena_, r);
      if (n == nullptr) {
        return nullptr;
      }
      n->set_range(SpanRange(start_range, end_range));
      return n;
    }
    // Fall through: the Ident is a defined name; the binary `:` rule will
    // pair it with whatever the RHS atom parses to.
  }

  // Default: NameRef.
  advance();
  AstNode* n = make_name_ref(arena_, ident.lexeme);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(ident.range);
  return n;
}

namespace {

// Returns true iff `name` starts with `[A-Za-z_]` and continues with
// `[A-Za-z0-9_.?]*`. Length must be non-zero and <= 255 bytes. This is the
// identifier shape Excel accepts for LET bindings and defined names; stricter
// than the tokenizer's Ident rule (which accepts numeric suffixes alone).
bool IsLetNameShape(std::string_view name) noexcept {
  if (name.empty() || name.size() > 255) {
    return false;
  }
  const char first = name[0];
  if (!IsAsciiLetter(first) && first != '_') {
    return false;
  }
  for (std::size_t i = 1; i < name.size(); ++i) {
    const char c = name[i];
    const bool ok = IsAsciiLetter(c) || IsAsciiDigit(c) || c == '_' || c == '.' || c == '?';
    if (!ok) {
      return false;
    }
  }
  return true;
}

// Returns true iff `name` has the shape of an A1-style cell reference
// (`[A-Za-z]{1,3}[1-9][0-9]*`). Excel rejects such identifiers as LET
// binding names because they would be indistinguishable from cell refs when
// used in the body.
bool LooksLikeCellRef(std::string_view name) noexcept {
  std::size_t i = 0;
  while (i < name.size() && IsAsciiLetter(name[i])) {
    ++i;
  }
  const std::size_t letters = i;
  if (letters == 0 || letters > 3) {
    return false;
  }
  // First digit must not be zero so names like `A0` (not a valid ref) still
  // go through the LET path.
  if (i >= name.size() || name[i] < '1' || name[i] > '9') {
    return false;
  }
  ++i;
  while (i < name.size() && IsAsciiDigit(name[i])) {
    ++i;
  }
  return i == name.size();
}

}  // namespace

bool Parser::parse_let_binding_name(std::string_view* out_name, TextRange* out_range) {
  // The binding-name slot accepts a single Ident token. CellRef tokens
  // (produced for patterns that match the A1 shape) are rejected because the
  // tokenizer already routed them away from Ident; this is the expected
  // behaviour for names like `A1` that collide with cell refs.
  const Token& tok = peek();
  if (tok.kind == TokenKind::CellRef) {
    record_error_with_token(ParseErrorCode::LetInvalidName, tok.range, tok.lexeme);
    advance();
    return false;
  }
  if (tok.kind != TokenKind::Ident) {
    record_error_with_token(ParseErrorCode::LetInvalidName, tok.range, tok.lexeme);
    return false;
  }
  if (!IsLetNameShape(tok.lexeme) || LooksLikeCellRef(tok.lexeme)) {
    record_error_with_token(ParseErrorCode::LetInvalidName, tok.range, tok.lexeme);
    advance();
    return false;
  }
  *out_name = tok.lexeme;
  *out_range = tok.range;
  advance();
  return true;
}

AstNode* Parser::parse_let_call(const Token& name_tok) {
  const TextRange call_start = name_tok.range;
  advance();  // LET Ident
  advance();  // LParen

  // Accumulate name / expr pairs until a bare expression (the body) remains.
  std::vector<std::string_view> names;
  std::vector<const AstNode*> exprs;
  AstNode* body = nullptr;

  // Empty arg list is an arity error but still needs recovery; we synthesise
  // a placeholder body and return below.
  if (peek_kind() == TokenKind::RParen) {
    record_error_with_token(ParseErrorCode::LetWrongArity, name_tok.range, name_tok.lexeme);
    const Token& rparen_empty = advance();
    AstNode* placeholder = make_error_placeholder(arena_);
    if (placeholder != nullptr) {
      placeholder->set_range(SpanRange(call_start, rparen_empty.range));
    }
    return placeholder;
  }

  // Slot-walk loop. At each iteration pos_ sits on the first token of the
  // next slot. A slot is classified as a binding name when it is a lone
  // well-shaped Ident followed by a comma (so it cannot stand as a complete
  // expression); any other shape is the body. This works for all odd-arity
  // well-formed inputs because the body is always the *final* slot, never
  // followed by a comma.
  while (true) {
    if (bailed_) {
      break;
    }
    // CellRef-shaped tokens (e.g. `A1`, `AA10`) that sit in a binding-name
    // slot are specifically forbidden by Excel: the name would collide with
    // the A1 cell it spells. Detect the shape here so we emit the dedicated
    // LetInvalidName diagnostic instead of letting the token fall into the
    // body path (which would either parse it as a Ref or surface a
    // non-specific arity error).
    if (peek_kind() == TokenKind::CellRef && peek_kind_at(1) == TokenKind::Comma) {
      const Token& bad = peek();
      record_error_with_token(ParseErrorCode::LetInvalidName, bad.range, bad.lexeme);
      advance();  // consume the cell-ref
      if (peek_kind() == TokenKind::Comma) {
        advance();  // consume the comma
      }
      // Parse (and discard) the would-be initialiser so siblings continue.
      AstNode* expr = parse_expression(0, SyncContext::CallArg);
      if (expr == nullptr) {
        return nullptr;
      }
      // The LET grammar requires an odd total arity; since we dropped this
      // pair, the final arity is still consistent: continue to the next slot.
      if (peek_kind() == TokenKind::Comma) {
        advance();
        continue;
      }
      // No comma: the expression we just parsed was the tail; promote it
      // to body if no bindings were valid yet.
      if (body == nullptr && names.empty()) {
        body = expr;
      }
      break;
    }
    const bool is_name_slot = (peek_kind() == TokenKind::Ident) && IsLetNameShape(peek().lexeme) &&
                              !LooksLikeCellRef(peek().lexeme) && peek_kind_at(1) == TokenKind::Comma;
    if (is_name_slot) {
      std::string_view name;
      // The name's source span is written via parse_let_binding_name; we do
      // not currently attach it to the LetBinding node, but keeping the slot
      // lets that diagnostic link land without another signature change.
      TextRange name_range{};
      if (!parse_let_binding_name(&name, &name_range)) {
        (void)name_range;
        skip_to_sync(SyncContext::CallArg);
        if (peek_kind() == TokenKind::Comma) {
          advance();
        }
        continue;
      }
      // parse_let_binding_name already advanced over the name; the next
      // token is the comma guaranteed by `is_name_slot`.
      advance();  // Comma
      AstNode* expr = parse_expression(0, SyncContext::CallArg);
      if (expr == nullptr) {
        return nullptr;  // hard arena failure
      }
      names.push_back(name);
      exprs.push_back(expr);
      if (peek_kind() == TokenKind::Comma) {
        advance();
        continue;  // more slots follow
      }
      // No comma after an (name, expr) pair means this was the penultimate
      // slot and the expr we just parsed was the tail of a well-formed LET
      // that is missing its body. Treat as an arity error.
      record_error_with_token(ParseErrorCode::LetWrongArity, name_tok.range, name_tok.lexeme);
      break;
    }

    // Body slot: parse a full expression; the next token should be `)`.
    body = parse_expression(0, SyncContext::CallArg);
    if (body == nullptr) {
      return nullptr;  // hard arena failure
    }
    if (peek_kind() == TokenKind::Comma) {
      // Extra argument after the body (even arity). Record and recover by
      // skipping the rest of the arglist.
      record_error_with_token(ParseErrorCode::LetWrongArity, name_tok.range, name_tok.lexeme);
      skip_to_sync(SyncContext::Paren);
    }
    break;
  }

  // At this point we expect `)`; emit diagnostics and recover if not.
  TextRange end_range = call_start;
  if (peek_kind() == TokenKind::RParen) {
    const Token& rparen = advance();
    end_range = rparen.range;
  } else if (peek_kind() != TokenKind::Eof) {
    record_error_with_token(ParseErrorCode::ExpectedCloseParen, call_start, name_tok.lexeme);
  } else {
    record_error_with_token(ParseErrorCode::ExpectedCloseParen, call_start, name_tok.lexeme);
  }

  // Validate arity: need >= 1 binding and a body.
  if (names.empty() || body == nullptr) {
    if (body == nullptr) {
      record_error_with_token(ParseErrorCode::LetWrongArity, call_start, name_tok.lexeme);
    } else if (names.empty()) {
      record_error_with_token(ParseErrorCode::LetWrongArity, call_start, name_tok.lexeme);
    }
    AstNode* placeholder = make_error_placeholder(arena_);
    if (placeholder != nullptr) {
      placeholder->set_range(SpanRange(call_start, end_range));
    }
    return placeholder;
  }

  AstNode* n = make_let_binding(arena_, names.data(), exprs.data(), static_cast<std::uint32_t>(names.size()), body);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(call_start, end_range));
  return n;
}

AstNode* Parser::parse_sheet_qualified_ref(std::string_view sheet, bool quoted, TextRange sheet_range) {
  // Expect Bang next. (For SheetName tokens we have not yet consumed Bang;
  // for unquoted Ident the caller has consumed only the Ident.)
  if (peek_kind() != TokenKind::Bang) {
    record_error_with_token(ParseErrorCode::UnexpectedToken, peek().range, peek().lexeme);
    return nullptr;
  }
  advance();  // Bang

  // Three possibilities:
  //   1. CellRef: `Sheet1!A1`.
  //   2. Ident Colon Ident with matching column letters: `Sheet1!A:A`.
  //   3. Number Colon Number with matching row digits: `Sheet1!1:1`.
  // Anything else is an error.
  const TokenKind k = peek_kind();
  if (k == TokenKind::CellRef) {
    const Token& cell = advance();
    Reference r;
    if (!decode_cellref_lexeme(cell.lexeme, &r)) {
      record_error_with_token(ParseErrorCode::InvalidReference, cell.range, cell.lexeme);
      return nullptr;
    }
    r.sheet = sheet;
    r.sheet_quoted = quoted;
    AstNode* n = make_ref(arena_, r);
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(SpanRange(sheet_range, cell.range));
    return n;
  }
  if (k == TokenKind::Ident && peek_kind_at(1) == TokenKind::Colon && peek_kind_at(2) == TokenKind::Ident) {
    const Token& lhs_tok = peek();
    const Token& rhs_tok = peek_at(2);
    bool lhs_abs = false;
    bool rhs_abs = false;
    const std::uint32_t lhs_col = decode_column_letters(lhs_tok.lexeme, &lhs_abs);
    const std::uint32_t rhs_col = decode_column_letters(rhs_tok.lexeme, &rhs_abs);
    if (lhs_col != 0 && lhs_col == rhs_col) {
      const TextRange end_range = rhs_tok.range;
      advance();
      advance();
      advance();
      Reference r;
      r.col = lhs_col - 1;
      r.col_abs = lhs_abs;
      r.is_full_col = true;
      r.sheet = sheet;
      r.sheet_quoted = quoted;
      AstNode* n = make_ref(arena_, r);
      if (n == nullptr) {
        return nullptr;
      }
      n->set_range(SpanRange(sheet_range, end_range));
      return n;
    }
  }
  if (k == TokenKind::Number && peek_kind_at(1) == TokenKind::Colon && peek_kind_at(2) == TokenKind::Number) {
    const Token& lhs_tok = peek();
    const Token& rhs_tok = peek_at(2);
    auto is_pure_digit_run = [](std::string_view lex) {
      if (lex.empty()) {
        return false;
      }
      for (char c : lex) {
        if (!IsAsciiDigit(c)) {
          return false;
        }
      }
      return true;
    };
    if (lhs_tok.is_integer && rhs_tok.is_integer && is_pure_digit_run(lhs_tok.lexeme) &&
        is_pure_digit_run(rhs_tok.lexeme)) {
      std::uint64_t lhs_row = 0;
      for (char c : lhs_tok.lexeme) {
        lhs_row = lhs_row * 10u + static_cast<std::uint32_t>(c - '0');
      }
      std::uint64_t rhs_row = 0;
      for (char c : rhs_tok.lexeme) {
        rhs_row = rhs_row * 10u + static_cast<std::uint32_t>(c - '0');
      }
      if (lhs_row != 0 && lhs_row == rhs_row && lhs_row <= kMaxRow) {
        const TextRange end_range = rhs_tok.range;
        advance();
        advance();
        advance();
        Reference r;
        r.row = static_cast<std::uint32_t>(lhs_row - 1);
        r.is_full_row = true;
        r.sheet = sheet;
        r.sheet_quoted = quoted;
        AstNode* n = make_ref(arena_, r);
        if (n == nullptr) {
          return nullptr;
        }
        n->set_range(SpanRange(sheet_range, end_range));
        return n;
      }
    }
  }
  record_error_with_token(ParseErrorCode::InvalidReference, peek().range, peek().lexeme);
  return nullptr;
}

}  // namespace parser
}  // namespace formulon
