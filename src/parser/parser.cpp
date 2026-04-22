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

bool IsAsciiDigit(char c) noexcept { return c >= '0' && c <= '9'; }

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
    case ParseErrorCode::UnclosedParen:
      return "unclosed parenthesis";
    case ParseErrorCode::UnclosedBrace:
      return "unclosed array literal";
    case ParseErrorCode::ArrayRowMismatch:
      return "array literal rows have inconsistent column counts";
    case ParseErrorCode::ExpectedRParenOrComma:
      return "expected ')' or ','";
    case ParseErrorCode::ExpectedCommaOrSemiInArray:
      return "expected ',' or ';' in array literal";
    case ParseErrorCode::InvalidCellRef:
      return "invalid cell reference";
    case ParseErrorCode::UnsupportedConstruct:
      return "construct not supported in this milestone";
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

TokenKind Parser::peek_kind() const noexcept { return peek().kind; }

TokenKind Parser::peek_kind_at(std::size_t offset) const noexcept { return peek_at(offset).kind; }

const Token& Parser::peek() const noexcept { return peek_at(0); }

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
  if (bailed_) {
    return;
  }
  ParseError e;
  e.code = code;
  e.range = range;
  e.message = std::string_view(default_message(code));
  errors_.push_back(e);
  bailed_ = true;
}

void Parser::promote_lexer_errors(const std::vector<LexerError>& lex_errors) {
  for (const auto& le : lex_errors) {
    ParseError e;
    e.code = PromoteLexerCode(le.code);
    e.range = le.range;
    e.message = std::string_view(default_message(e.code));
    errors_.push_back(e);
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

  root_ = parse_expression(0);
  if (root_ == nullptr) {
    return nullptr;
  }

  // Trailing tokens that we did not consume are an error; we surface the
  // first such token as UnexpectedToken.
  if (peek_kind() != TokenKind::Eof && !bailed_) {
    record_error(ParseErrorCode::UnexpectedToken, peek().range);
    return nullptr;
  }
  return root_;
}

// ---------------------------------------------------------------------------
// Pratt expression loop
// ---------------------------------------------------------------------------

AstNode* Parser::parse_expression(int min_bp) {
  AstNode* lhs = parse_atom();
  if (lhs == nullptr) {
    return nullptr;
  }

  while (true) {
    if (bailed_) {
      return nullptr;
    }
    const TokenKind kind = peek_kind();

    // Postfix operators first (currently only `%`).
    if (kind == TokenKind::Percent) {
      if (kBpPostfixPercent < min_bp) {
        return lhs;
      }
      const Token& tok = advance();
      AstNode* node = make_unary_op(arena_, UnaryOp::Percent, lhs);
      if (node == nullptr) {
        record_error(ParseErrorCode::UnexpectedToken, tok.range);
        return nullptr;
      }
      node->set_range(SpanRange(lhs->range(), tok.range));
      lhs = node;
      continue;
    }

    int right_bp = 0;
    const int bp = InfixBindingPower(kind, &right_bp);
    if (bp == 0 || bp < min_bp) {
      return lhs;
    }

    advance();  // consume the operator token.
    AstNode* rhs = parse_expression(right_bp);
    if (rhs == nullptr) {
      return nullptr;
    }

    AstNode* node = nullptr;
    if (kind == TokenKind::Colon) {
      node = make_range_op(arena_, lhs, rhs);
    } else {
      node = make_binary_op(arena_, TokenToBinOp(kind), lhs, rhs);
    }
    if (node == nullptr) {
      record_error(ParseErrorCode::UnexpectedToken, peek().range);
      return nullptr;
    }
    node->set_range(SpanRange(lhs->range(), rhs->range()));
    lhs = node;
  }
}

// ---------------------------------------------------------------------------
// Atom dispatch
// ---------------------------------------------------------------------------

AstNode* Parser::parse_atom() {
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
      return parse_at_prefix_atom();
    case TokenKind::Plus:
      return parse_unary_prefix_atom(UnaryOp::Plus);
    case TokenKind::Minus:
      return parse_unary_prefix_atom(UnaryOp::Minus);
    case TokenKind::Ident:
      return parse_ident_or_call_or_full_col();
    case TokenKind::SheetName: {
      const Token& sheet = advance();
      return parse_sheet_qualified_ref(sheet.text, /*quoted=*/true, sheet.range);
    }
    case TokenKind::Eof:
      record_error(ParseErrorCode::UnexpectedEof, peek().range);
      return nullptr;
    case TokenKind::String:
    case TokenKind::LBracket:
    case TokenKind::Hash:
      // Deferred to follow-up work (string literals, structured/external refs,
      // spilled-range `#`). Surface a single, unmistakable diagnostic and stop.
      record_error(ParseErrorCode::UnsupportedConstruct, peek().range);
      return nullptr;
    case TokenKind::Invalid:
      // The tokenizer already recorded a LexerError for this; promote to a
      // parser-level UnexpectedToken so the caller sees one definite stop.
      record_error(ParseErrorCode::UnexpectedToken, peek().range);
      return nullptr;
    default:
      record_error(ParseErrorCode::ExpectedExpression, peek().range);
      return nullptr;
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
    record_error(ParseErrorCode::ExpectedExpression, tok.range);
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
      record_error(ParseErrorCode::ExpectedExpression, first.range);
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
      record_error(ParseErrorCode::ExpectedExpression, first.range);
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
    record_error(ParseErrorCode::InvalidCellRef, first.range);
    return nullptr;
  }
  n->set_range(SpanRange(first.range, second.range));
  return n;
}

AstNode* Parser::parse_bool_atom() {
  const Token& tok = advance();
  AstNode* n = make_literal(arena_, Value::boolean(tok.boolean));
  if (n == nullptr) {
    record_error(ParseErrorCode::ExpectedExpression, tok.range);
    return nullptr;
  }
  n->set_range(tok.range);
  return n;
}

AstNode* Parser::parse_error_literal_atom() {
  const Token& tok = advance();
  AstNode* n = make_error_literal(arena_, tok.error_code);
  if (n == nullptr) {
    record_error(ParseErrorCode::ExpectedExpression, tok.range);
    return nullptr;
  }
  n->set_range(tok.range);
  return n;
}

AstNode* Parser::parse_paren_atom() {
  const Token& lparen = advance();
  AstNode* inner = parse_expression(0);
  if (inner == nullptr) {
    return nullptr;
  }
  if (peek_kind() != TokenKind::RParen) {
    record_error(ParseErrorCode::UnclosedParen, lparen.range);
    return nullptr;
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
      default:
        record_error(ParseErrorCode::ExpectedExpression, peek().range);
        return nullptr;
    }
  };

  AstNode* first = parse_element();
  if (first == nullptr) {
    return nullptr;
  }
  elements.push_back(first);
  cur_row_cols = 1;

  while (true) {
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
        record_error(ParseErrorCode::ArrayRowMismatch, peek().range);
        return nullptr;
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
      record_error(ParseErrorCode::UnclosedBrace, lbrace.range);
      return nullptr;
    }
    record_error(ParseErrorCode::ExpectedCommaOrSemiInArray, peek().range);
    return nullptr;
  }

  // Final row's column count must match (or initialise) the array width.
  if (cols == 0) {
    cols = cur_row_cols;
  } else if (cur_row_cols != cols) {
    record_error(ParseErrorCode::ArrayRowMismatch, peek().range);
    return nullptr;
  }

  if (peek_kind() != TokenKind::RBrace) {
    record_error(ParseErrorCode::UnclosedBrace, lbrace.range);
    return nullptr;
  }
  const Token& rbrace = advance();

  FM_CHECK(elements.size() == static_cast<std::size_t>(rows) * cols, "array element count mismatch after parse");
  AstNode* n = make_array_literal(arena_, rows, cols, elements.data());
  if (n == nullptr) {
    record_error(ParseErrorCode::UnclosedBrace, lbrace.range);
    return nullptr;
  }
  n->set_range(SpanRange(lbrace.range, rbrace.range));
  return n;
}

AstNode* Parser::parse_at_prefix_atom() {
  const Token& at_tok = advance();
  // `@` consumes the entire remaining expression (lowest precedence).
  AstNode* operand = parse_expression(kBpAtPrefix);
  if (operand == nullptr) {
    return nullptr;
  }
  AstNode* n = make_implicit_intersection(arena_, operand);
  if (n == nullptr) {
    record_error(ParseErrorCode::UnexpectedToken, at_tok.range);
    return nullptr;
  }
  n->set_range(SpanRange(at_tok.range, operand->range()));
  return n;
}

AstNode* Parser::parse_unary_prefix_atom(UnaryOp op) {
  const Token& sign_tok = advance();
  AstNode* operand = parse_expression(kBpUnaryPrefix);
  if (operand == nullptr) {
    return nullptr;
  }
  AstNode* n = make_unary_op(arena_, op, operand);
  if (n == nullptr) {
    record_error(ParseErrorCode::UnexpectedToken, sign_tok.range);
    return nullptr;
  }
  n->set_range(SpanRange(sign_tok.range, operand->range()));
  return n;
}

AstNode* Parser::parse_cellref_atom() {
  const Token& tok = advance();
  Reference r;
  if (!decode_cellref_lexeme(tok.lexeme, &r)) {
    record_error(ParseErrorCode::InvalidCellRef, tok.range);
    return nullptr;
  }
  AstNode* n = make_ref(arena_, r);
  if (n == nullptr) {
    record_error(ParseErrorCode::InvalidCellRef, tok.range);
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
    const std::string_view fn_name = ident.lexeme;
    const TextRange call_start = ident.range;
    advance();  // Ident
    advance();  // LParen
    std::vector<const AstNode*> args;
    if (peek_kind() != TokenKind::RParen) {
      while (true) {
        AstNode* arg = parse_expression(0);
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
        record_error(ParseErrorCode::ExpectedRParenOrComma, peek().range);
        return nullptr;
      }
    }
    if (peek_kind() != TokenKind::RParen) {
      record_error(ParseErrorCode::UnclosedParen, call_start);
      return nullptr;
    }
    const Token& rparen = advance();
    AstNode* n = make_call(arena_, fn_name, args.data(), static_cast<std::uint32_t>(args.size()));
    if (n == nullptr) {
      record_error(ParseErrorCode::UnexpectedToken, call_start);
      return nullptr;
    }
    n->set_range(SpanRange(call_start, rparen.range));
    return n;
  }

  if (next == TokenKind::Bang) {
    const std::string_view sheet = ident.lexeme;
    const TextRange sheet_range = ident.range;
    advance();  // Ident
    return parse_sheet_qualified_ref(sheet, /*quoted=*/false, sheet_range);
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
        record_error(ParseErrorCode::InvalidCellRef, start_range);
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
    record_error(ParseErrorCode::ExpectedExpression, ident.range);
    return nullptr;
  }
  n->set_range(ident.range);
  return n;
}

AstNode* Parser::parse_sheet_qualified_ref(std::string_view sheet, bool quoted, TextRange sheet_range) {
  // Expect Bang next. (For SheetName tokens we have not yet consumed Bang;
  // for unquoted Ident the caller has consumed only the Ident.)
  if (peek_kind() != TokenKind::Bang) {
    record_error(ParseErrorCode::UnexpectedToken, peek().range);
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
      record_error(ParseErrorCode::InvalidCellRef, cell.range);
      return nullptr;
    }
    r.sheet = sheet;
    r.sheet_quoted = quoted;
    AstNode* n = make_ref(arena_, r);
    if (n == nullptr) {
      record_error(ParseErrorCode::InvalidCellRef, sheet_range);
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
        record_error(ParseErrorCode::InvalidCellRef, sheet_range);
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
          record_error(ParseErrorCode::InvalidCellRef, sheet_range);
          return nullptr;
        }
        n->set_range(SpanRange(sheet_range, end_range));
        return n;
      }
    }
  }
  record_error(ParseErrorCode::InvalidCellRef, peek().range);
  return nullptr;
}

}  // namespace parser
}  // namespace formulon
