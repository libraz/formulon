// Copyright 2026 libraz. Licensed under the MIT License.

#include <cstdint>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/parse_error.h"
#include "parser/parser.h"
#include "parser/parser_detail.h"
#include "parser/reference.h"
#include "parser/token.h"

namespace formulon {
namespace parser {

using detail::IsAsciiDigit;
using detail::IsAsciiLetter;
using detail::kMaxColumn;
using detail::kMaxRow;
using detail::SpanRange;

namespace {

// Returns true iff `name` starts with `[A-Za-z_]` or any non-ASCII (UTF-8
// continuation) byte, and continues with `[A-Za-z0-9_.?]` plus non-ASCII
// bytes. Length must be non-zero and <= 255 bytes. This mirrors the
// tokenizer's identifier rule (`Tokenizer::is_ident_start_byte` /
// `is_ident_cont_byte`) and matches the identifier shape Excel accepts for
// LET bindings and defined names, including hiragana / katakana / kanji and
// other locale-specific scripts. The shape is still stricter than the
// tokenizer's Ident rule because it forbids a leading ASCII digit.
bool IsLetNameShape(std::string_view name) noexcept {
  if (name.empty() || name.size() > 255) {
    return false;
  }
  const auto first = static_cast<unsigned char>(name[0]);
  if (!IsAsciiLetter(static_cast<char>(first)) && first != '_' && first < 0x80) {
    return false;
  }
  for (std::size_t i = 1; i < name.size(); ++i) {
    const auto byte = static_cast<unsigned char>(name[i]);
    const char c = static_cast<char>(byte);
    const bool ok =
        IsAsciiLetter(c) || IsAsciiDigit(c) || c == '_' || c == '.' || c == '?' || byte >= 0x80;
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
// LET bindings
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Sheet-qualified refs
// ---------------------------------------------------------------------------

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
