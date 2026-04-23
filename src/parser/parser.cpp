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
//
// Source layout: this TU owns the parser entry point, the Pratt expression
// loop, atom dispatch, and the panic-mode plumbing. Per-atom helpers live in
// `parser_atoms.cpp`; cell-ref decoding, LET special-form handling, and
// sheet-qualified refs live in `parser_reference.cpp`.

#include "parser/parser.h"

#include <cstdint>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/lexer_error.h"
#include "parser/parse_error.h"
#include "parser/parser_detail.h"
#include "parser/token.h"
#include "parser/tokenizer.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {

using detail::kBpAddSub;
using detail::kBpComparison;
using detail::kBpConcat;
using detail::kBpMulDiv;
using detail::kBpPostfixPercent;
using detail::kBpPow;
using detail::kBpRange;
using detail::SpanRange;

namespace {

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

}  // namespace parser
}  // namespace formulon
