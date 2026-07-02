// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the Pratt parser. Driven by a single `Tokenizer` pass;
// whitespace tokens are stripped during ingest. A leading `=` (Excel
// formula prefix) is consumed if it sits at the very start.
//
// Operator binding-power scheme (higher = tighter):
//
//   Postfix `#`           90   (spilled-range; only valid on a single Ref)
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
using detail::kBpIntersect;
using detail::kBpMulDiv;
using detail::kBpPostfixCall;
using detail::kBpPostfixHash;
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
    case TokenKind::Whitespace:
      // Excel's space-as-intersection operator. Only the whitespace tokens
      // that sit between two reference-shaped tokens reach this point; the
      // retention pass in `Parser::parse()` strips every other run.
      *right_bp = kBpIntersect + 1;
      return kBpIntersect;
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
    case ParseErrorCode::LambdaInvalidParam:
      return "invalid LAMBDA parameter name";
    case ParseErrorCode::LambdaEmpty:
      return "LAMBDA requires at least a body expression";
    case ParseErrorCode::LambdaDuplicateParam:
      return "LAMBDA parameter names must be unique";
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
  // Cap check happens *before* appending the real error: the budget
  // reserves one slot for the `TooManyErrors` sentinel, so at most
  // `max_error_count - 1` real diagnostics are ever recorded. Checking
  // first (rather than appending unconditionally and capping after) is
  // what keeps `errors_.size() <= max_error_count` true even at the
  // `max_error_count == 1` edge, where zero real diagnostics are
  // recorded and the sentinel alone occupies the single slot.
  if (opts_.max_error_count > 0 && errors_.size() + 1 >= opts_.max_error_count) {
    ParseError sentinel;
    sentinel.code = ParseErrorCode::TooManyErrors;
    sentinel.range = range;
    sentinel.message = std::string_view(default_message(ParseErrorCode::TooManyErrors));
    sentinel.severity = Severity::Error;
    errors_.push_back(sentinel);
    bailed_ = true;
    return;
  }

  ParseError e;
  e.code = code;
  e.range = range;
  e.message = std::string_view(default_message(code));
  e.offending_token = offending;
  e.severity = Severity::Error;
  e.suggestion = std::string_view{};
  errors_.push_back(e);
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

AstNode* Parser::make_recovery_placeholder(TextRange range) {
  AstNode* placeholder = make_error_placeholder(arena_);
  if (placeholder != nullptr) {
    placeholder->set_range(range);
  }
  return placeholder;
}

AstNode* Parser::recover_with_placeholder(ParseErrorCode code, const Token& tok, SyncContext ctx) {
  record_error_with_token(code, tok.range, tok.lexeme);
  skip_to_sync(ctx);
  return make_recovery_placeholder(tok.range);
}

// ---------------------------------------------------------------------------
// Top-level parse
// ---------------------------------------------------------------------------

AstNode* Parser::parse() {
  if (parsed_) {
    return root_;
  }
  parsed_ = true;

  // Drive the tokenizer and pre-filter whitespace into `tokens_`. Most
  // whitespace runs are pure layout and should be dropped, but Excel uses a
  // space character as the binary intersection operator when it sits between
  // two reference-shaped operands (e.g. `A1:C3 B2:D4`). We retain those
  // whitespace tokens so the Pratt loop can promote them via
  // `InfixBindingPower(TokenKind::Whitespace, ...)`. The candidate set here
  // is intentionally conservative -- it is shape-based and therefore admits
  // false positives like `1 2`; the LHS-shape guard in `parse_expression`
  // re-validates that the left operand is reference-shaped before consuming
  // the whitespace as an operator, so non-reference noise still degrades to
  // a normal whitespace strip.
  Tokenizer tz(source_);
  const auto& raw = tz.tokens();
  tokens_.reserve(raw.size());
  // `RefCandidate` includes structurally reference-shaped token kinds: bare
  // cell refs, identifiers (which may resolve to a defined name or, when
  // followed by `(`, a reference-returning call), a closing paren (which
  // terminates a parenthesised reference or function call), a closing
  // bracket (which terminates a structured reference), and a number (the
  // whole-row form `2:2` begins and ends with a Number, so an intersect
  // operand like `B:B 2:2` brackets the whitespace with Number on the
  // right). The Pratt parser ultimately decides whether the AST shape
  // allows the intersection: a bare-literal operand is re-rejected there,
  // so admitting Number here only widens the conservative candidate set.
  auto is_ref_candidate = [](TokenKind k) noexcept {
    return k == TokenKind::CellRef || k == TokenKind::Ident || k == TokenKind::RParen || k == TokenKind::RBracket ||
           k == TokenKind::Number;
  };
  // Walk `raw` with a sliding window over the most recent non-whitespace
  // token and the next non-whitespace token after each whitespace run. If
  // both bracket the whitespace as reference-shaped, keep it; otherwise drop.
  for (std::size_t i = 0; i < raw.size(); ++i) {
    const Token& t = raw[i];
    if (t.kind != TokenKind::Whitespace) {
      tokens_.push_back(t);
      continue;
    }
    // Find the next non-whitespace token after this run.
    TokenKind next_kind = TokenKind::Eof;
    for (std::size_t j = i + 1; j < raw.size(); ++j) {
      if (raw[j].kind != TokenKind::Whitespace) {
        next_kind = raw[j].kind;
        break;
      }
    }
    // Find the previous non-whitespace token already pushed onto `tokens_`.
    TokenKind prev_kind = tokens_.empty() ? TokenKind::Eof : tokens_.back().kind;
    if (is_ref_candidate(prev_kind) && is_ref_candidate(next_kind)) {
      tokens_.push_back(t);
    }
    // Otherwise: drop. Subsequent iterations resume on the next token.
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
  // We track which `errors_` entries were promoted from the lexer so a
  // post-parse cleanup pass can drop the ones whose source position falls
  // inside a structured-reference bracket payload (the lexer flags
  // `#All` / `#Headers` / ... as `InvalidErrorLiteral` since they do not
  // match the canonical Excel error catalogue, but those are valid
  // payload bytes once the parser has identified the surrounding
  // `Ident LBracket ... RBracket` shape).
  const std::size_t lexer_errors_first = errors_.size();
  promote_lexer_errors(tz.errors());
  const std::size_t lexer_errors_last = errors_.size();
  // Snapshot the lexeme byte-positions of every lexer error so we can
  // correlate them with `struct_ref_byte_spans_` after parsing. The
  // `errors_` vector itself only stores UTF-16 ranges; the lexeme view
  // gives us byte offsets straight into `source_`.
  std::vector<std::uint32_t> lexer_error_byte_starts;
  lexer_error_byte_starts.reserve(tz.errors().size());
  for (const auto& le : tz.errors()) {
    if (!le.lexeme.empty()) {
      lexer_error_byte_starts.push_back(static_cast<std::uint32_t>(le.lexeme.data() - source_.data()));
    } else {
      // Fallback: errors without a lexeme cannot be range-correlated; flag
      // them as "always preserve" by writing a sentinel that no span will
      // match.
      lexer_error_byte_starts.push_back(static_cast<std::uint32_t>(-1));
    }
  }

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

  // Drop any lexer-promoted errors whose lexeme byte-offset fell inside a
  // consumed structured-reference bracket payload. The lexer cannot know
  // it is scanning a structured-ref token (it has no parser-level
  // context), so it flags `#All`, `#Headers`, ... as `InvalidErrorLiteral`;
  // by the time we reach this point the parser has identified those bytes
  // as the bracket payload of a `StructuredRef` AST node, so they should
  // not surface as user-visible diagnostics.
  if (!struct_ref_byte_spans_.empty() && lexer_errors_last > lexer_errors_first) {
    auto in_struct_span = [this](std::uint32_t byte_pos) noexcept {
      if (byte_pos == static_cast<std::uint32_t>(-1)) {
        return false;
      }
      for (const auto& span : struct_ref_byte_spans_) {
        if (byte_pos >= span.first && byte_pos < span.second) {
          return true;
        }
      }
      return false;
    };
    std::size_t write = lexer_errors_first;
    for (std::size_t read = lexer_errors_first; read < lexer_errors_last; ++read) {
      const std::uint32_t byte_pos = lexer_error_byte_starts[read - lexer_errors_first];
      if (in_struct_span(byte_pos)) {
        continue;
      }
      if (write != read) {
        errors_[write] = errors_[read];
      }
      ++write;
    }
    if (write < lexer_errors_last) {
      const std::size_t removed = lexer_errors_last - write;
      // Shift any post-lexer parse-error entries left to fill the gap.
      for (std::size_t i = lexer_errors_last; i < errors_.size(); ++i) {
        errors_[i - removed] = errors_[i];
      }
      errors_.resize(errors_.size() - removed);
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
    AstNode* placeholder = make_recovery_placeholder(peek().range);
    --depth_;
    return placeholder;
  }
  if (bailed_) {
    AstNode* placeholder = make_recovery_placeholder(peek().range);
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

    // Postfix `(`: immediately-invoked function expression. Wraps the most
    // recent LHS in a `LambdaCall` node with the parsed argument list.
    // Covers `LAMBDA(x, x+1)(5)`, currying like
    // `LAMBDA(x, LAMBDA(y, x+y))(3)(4)`, and `(some_expr)(args)` where the
    // parenthesised expression evaluates to a Lambda. The normal
    // `Ident(args)` function-call path is unaffected because
    // `parse_ident_or_call_or_full_col` consumes the Ident and the matching
    // `(` together before this loop ever sees them; only an `(` that
    // appears AFTER a fully-parsed subtree falls into this rule.
    if (kind == TokenKind::LParen) {
      // Gate by LHS shape so we do not turn `=TRUE(1)` (Bool literal then
      // `(`) and similar non-callable forms into LambdaCall nodes the user
      // did not write. Only `Lambda` (an immediate IIFE) and `LambdaCall`
      // (chained curry) participate. Parenthesised lambda expressions like
      // `(LAMBDA(x, x))(5)` still work because `parse_paren_atom` unwraps a
      // single inner expression to its own kind; the outer Lambda kind is
      // preserved across the paren wrapper. The normal `Ident(args)`
      // function-call path is unaffected: `parse_ident_or_call_or_full_col`
      // consumes the Ident and the matching `(` together before this loop
      // ever sees them.
      const NodeKind lk = lhs->kind();
      if (lk != NodeKind::Lambda && lk != NodeKind::LambdaCall) {
        // Special-case: a literal LHS (Bool / Number / Text) followed by an
        // empty `()` is treated as a no-op so the surrounding Pratt loop can
        // continue and pick up trailing operators. The motivating case is
        // `=TRUE()+0`: the tokenizer always maps `TRUE` / `FALSE` to a Bool
        // literal token, so the Pratt loop sees `Bool '(' ')' '+' '0'`. Without
        // this branch the bail below leaves `()` unconsumed, which then
        // surfaces as `UnexpectedToken` and `skip_to_sync` discards `+0`.
        // Excel's documented behaviour is that `=TRUE()+0` evaluates to `1`.
        // Non-empty arg lists (`=TRUE(1)`) keep the existing bail path so the
        // `KeywordTokenIsAlwaysBoolLiteral` convention still holds.
        if (lk == NodeKind::Literal && peek_kind_at(1) == TokenKind::RParen) {
          if (kBpPostfixCall < min_bp) {
            --depth_;
            return lhs;
          }
          const Token& lparen_tok = advance();
          const Token& rparen_tok = advance();
          lhs->set_range(SpanRange(lhs->range(), rparen_tok.range));
          (void)lparen_tok;
          continue;
        }
        --depth_;
        return lhs;
      }
      if (kBpPostfixCall < min_bp) {
        --depth_;
        return lhs;
      }
      const Token& lparen = advance();
      std::vector<const AstNode*> args;
      if (peek_kind() != TokenKind::RParen) {
        while (true) {
          if (bailed_) {
            break;
          }
          AstNode* arg = nullptr;
          const TokenKind here = peek_kind();
          if (here == TokenKind::Comma || here == TokenKind::RParen) {
            // Empty arg slot: treat as a Blank, mirroring the regular
            // function-call argument loop in
            // `parse_ident_or_call_or_full_col`.
            arg = make_literal(arena_, Value::blank());
            if (arg != nullptr) {
              arg->set_range(peek().range);
            }
          } else {
            arg = parse_expression(0, SyncContext::CallArg);
          }
          if (arg == nullptr) {
            bailed_ = true;
            --depth_;
            return lhs;
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
            record_error_with_token(ParseErrorCode::ExpectedCloseParen, lparen.range, lparen.lexeme);
            break;
          }
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
      TextRange end_range = lparen.range;
      if (peek_kind() == TokenKind::RParen) {
        const Token& rparen = advance();
        end_range = rparen.range;
      } else if (peek_kind() != TokenKind::Eof) {
        record_error_with_token(ParseErrorCode::ExpectedCloseParen, lparen.range, lparen.lexeme);
      }
      AstNode* node =
          make_lambda_call(arena_, lhs, args.empty() ? nullptr : args.data(), static_cast<std::uint32_t>(args.size()));
      if (node == nullptr) {
        bailed_ = true;
        --depth_;
        return lhs;
      }
      node->set_range(SpanRange(lhs->range(), end_range));
      lhs = node;
      continue;
    }

    // Postfix `#`: spilled-range operator (e.g. `=A1#`). Only valid on a
    // single-cell `Ref` atom; reject anything else (range, full-column /
    // full-row, function call, arithmetic) with a diagnostic and consume
    // the `#` so the surrounding context keeps parsing. Bound tighter than
    // `:` so that `=A1:B2#` first consumes `#` against `B2` (yielding a
    // SpillRef), and the `:` shape check then rejects SpillRef as a range
    // endpoint. `=A1#+B1#` similarly parses as
    // `SpillRef(A1)+SpillRef(B1)`.
    if (kind == TokenKind::Hash) {
      if (kBpPostfixHash < min_bp) {
        --depth_;
        return lhs;
      }
      // Always consume `#` so we never re-enter the loop on the same token,
      // even on rejection. The diagnostic carries the Hash token's range.
      const Token& hash_tok = advance();
      if (lhs->kind() != NodeKind::Ref) {
        record_error_with_token(ParseErrorCode::UnsupportedConstruct, hash_tok.range, hash_tok.lexeme);
        // Surface an `ErrorPlaceholder` so siblings keep parsing.
        AstNode* placeholder = make_recovery_placeholder(SpanRange(lhs->range(), hash_tok.range));
        lhs = placeholder != nullptr ? placeholder : lhs;
        continue;
      }
      const Reference& r = lhs->as_ref();
      if (r.is_full_col || r.is_full_row) {
        // `A:A#` / `1:1#` are not legal Excel spill anchors.
        record_error_with_token(ParseErrorCode::UnsupportedConstruct, hash_tok.range, hash_tok.lexeme);
        AstNode* placeholder = make_recovery_placeholder(SpanRange(lhs->range(), hash_tok.range));
        lhs = placeholder != nullptr ? placeholder : lhs;
        continue;
      }
      AstNode* node = make_spill_ref(arena_, r);
      if (node == nullptr) {
        bailed_ = true;
        --depth_;
        return lhs;
      }
      node->set_range(SpanRange(lhs->range(), hash_tok.range));
      lhs = node;
      continue;
    }

    int right_bp = 0;
    int bp = InfixBindingPower(kind, &right_bp);
    // Whitespace only carries binding power when the LHS is reference-shaped.
    // The retention pass admits whitespace tokens whose neighbouring token
    // kinds are reference-candidates, but the parser-side AST shape is the
    // authoritative check: e.g. `1 2` slips through retention (Number is
    // Ident-like at the byte level only when treated as literal noise) but
    // must still parse as two separate atoms with the whitespace dropped.
    if (kind == TokenKind::Whitespace) {
      const NodeKind lk = lhs->kind();
      const bool lhs_ref_shaped = lk == NodeKind::Ref || lk == NodeKind::RangeOp || lk == NodeKind::ExternalRef ||
                                  lk == NodeKind::NameRef || lk == NodeKind::StructuredRef || lk == NodeKind::Call ||
                                  lk == NodeKind::IntersectOp;
      if (!lhs_ref_shaped) {
        // Treat the retained whitespace as layout: drop it and continue.
        advance();
        bp = 0;
      }
    }
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
      // A literal / unary-op on the rhs of `:` is not an Excel range.
      // `Call` is allowed here because Excel accepts reference-returning
      // calls (`OFFSET(...)`, `INDIRECT(...)`) as `:` endpoints; the
      // evaluator's `resolve_range_endpoint` discriminates between
      // reference-producing calls and value-producing calls (the latter
      // surface as `#VALUE!` at eval time via `resolve_reference_call`).
      // SpillRef is rejected on both sides: `A1#:B2` and `A1:B2#` are not
      // legal Excel range expressions (the spill operator is terminal).
      const NodeKind lk = lhs->kind();
      const NodeKind rk = rhs->kind();
      if (lk == NodeKind::SpillRef || rk == NodeKind::SpillRef) {
        record_error_with_token(ParseErrorCode::InvalidRange, op_tok.range, op_tok.lexeme);
      } else if (rk != NodeKind::Ref && rk != NodeKind::NameRef && rk != NodeKind::ExternalRef &&
                 rk != NodeKind::StructuredRef && rk != NodeKind::RangeOp && rk != NodeKind::Call &&
                 rk != NodeKind::ErrorPlaceholder) {
        record_error_with_token(ParseErrorCode::InvalidRange, op_tok.range, op_tok.lexeme);
      }
      node = make_range_op(arena_, lhs, rhs);
    } else if (kind == TokenKind::Whitespace) {
      // Space-as-intersection. The LHS shape was already validated above.
      // Validate the RHS shape with the same rules used for `:`; a non-
      // reference RHS records `InvalidRange` but we still wrap the children
      // so siblings keep parsing.
      const NodeKind rk = rhs->kind();
      if (rk != NodeKind::Ref && rk != NodeKind::NameRef && rk != NodeKind::ExternalRef &&
          rk != NodeKind::StructuredRef && rk != NodeKind::RangeOp && rk != NodeKind::Call &&
          rk != NodeKind::IntersectOp && rk != NodeKind::ErrorPlaceholder) {
        record_error_with_token(ParseErrorCode::InvalidRange, op_tok.range, op_tok.lexeme);
      }
      node = make_intersect_op(arena_, lhs, rhs);
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
      // 3-D reference whose first endpoint is a quoted sheet name
      // (`'My Sheet':Sheet3!A1`): SheetName followed by `:` then a sheet
      // name and then `!`.
      if (peek_kind_at(1) == TokenKind::Colon &&
          (peek_kind_at(2) == TokenKind::Ident || peek_kind_at(2) == TokenKind::SheetName) &&
          peek_kind_at(3) == TokenKind::Bang) {
        const Token& sheet1 = advance();  // SheetName
        AstNode* n3d = parse_3d_ref(sheet1.text, sheet1.range);
        if (n3d != nullptr) {
          return n3d;
        }
        skip_to_sync(ctx);
        return make_recovery_placeholder(sheet1.range);
      }
      const Token& sheet = advance();
      AstNode* n = parse_sheet_qualified_ref(sheet.text, /*quoted=*/true, sheet.range);
      if (n != nullptr) {
        return n;
      }
      skip_to_sync(ctx);
      return make_recovery_placeholder(sheet.range);
    }
    case TokenKind::LBracket: {
      // This arm only fires when `[` opens an atom with no preceding
      // identifier — i.e. a *bare* structured reference (`=[@col]`,
      // `=[col]`) or an external-book reference (`=[Book1.xlsx]Sheet1!A1`,
      // `=[1]Sheet1!A1`). Neither shape is supported: bare structured
      // refs have no table to qualify against, and no parser path builds
      // an `ExternalRef` AST node from source text at all (see
      // `tokenizer.h`'s design-notes comment). Both surface
      // `UnsupportedConstruct` (or `UnbalancedBrackets` if the bracket
      // never closes).
      //
      // Table-qualified structured refs (`=Table[col]`, `=Table[@col]`)
      // never reach this arm: they dispatch through `TokenKind::Ident`
      // to `parse_ident_or_call_or_full_col`, which builds a real
      // `StructuredRef` node.
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
      return recover_with_placeholder(code, lbracket, ctx);
    }
    case TokenKind::String:
      return parse_string_atom();
    case TokenKind::Hash: {
      // The spilled-range `#` is not implemented yet. Surface a single,
      // unmistakable diagnostic and recover.
      const Token& tok = peek();
      return recover_with_placeholder(ParseErrorCode::UnsupportedConstruct, tok, ctx);
    }
    case TokenKind::Invalid: {
      // The tokenizer already recorded a LexerError for this; promote to a
      // parser-level UnexpectedToken so the caller sees one definite stop.
      const Token& tok = peek();
      record_error_with_token(ParseErrorCode::UnexpectedToken, tok.range, tok.lexeme);
      advance();
      skip_to_sync(ctx);
      return make_recovery_placeholder(tok.range);
    }
    case TokenKind::Eof: {
      record_error(ParseErrorCode::UnexpectedEof, peek().range);
      return make_recovery_placeholder(peek().range);
    }
    default: {
      const Token& tok = peek();
      return recover_with_placeholder(ParseErrorCode::ExpectedExpression, tok, ctx);
    }
  }
}

}  // namespace parser
}  // namespace formulon
