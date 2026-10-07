
#include <cstdint>
#include <string_view>

#include "parser/ast.h"
#include "parser/parse_error.h"
#include "parser/parser.h"
#include "parser/parser_detail.h"
#include "parser/reference.h"
#include "parser/token.h"

namespace formulon {
namespace parser {

using detail::DecodeDigitRunClamped;
using detail::HasWorkbookExtension;
using detail::IsAsciiDigit;
using detail::IsAsciiLetter;
using detail::kMaxColumn;
using detail::kMaxRow;
using detail::SpanRange;

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

bool Parser::decode_full_col_endpoint(const Token& tok, Reference* out) noexcept {
  bool col_abs = false;
  const std::uint32_t col = decode_column_letters(tok.lexeme, &col_abs);
  if (col == 0) {
    return false;
  }
  out->col = col - 1U;
  out->col_abs = col_abs;
  out->is_full_col = true;
  return true;
}

bool Parser::decode_full_row_endpoint(const Token& tok, Reference* out) noexcept {
  // The tokenizer keeps an absolute anchor's `$` in the Number lexeme.
  std::string_view lex = tok.lexeme;
  const bool row_abs = !lex.empty() && lex.front() == '$';
  if (row_abs) {
    lex.remove_prefix(1);
  }
  if (!tok.is_integer || lex.empty()) {
    return false;
  }
  for (char c : lex) {
    if (!IsAsciiDigit(c)) {
      return false;
    }
  }
  // Clamped so a pathological 20-digit row literal cannot wrap a
  // `std::uint64_t` back into the valid range.
  const std::uint64_t row = DecodeDigitRunClamped(lex, kMaxRow);
  if (row == 0 || row > kMaxRow) {
    return false;
  }
  out->row = static_cast<std::uint32_t>(row - 1U);
  out->row_abs = row_abs;
  out->is_full_row = true;
  return true;
}

AstNode* Parser::make_whole_axis_ref(const Reference& lhs, const Reference& rhs, TextRange lhs_range,
                                     TextRange rhs_range) {
  if (lhs.col == rhs.col && lhs.row == rhs.row) {
    AstNode* n = make_ref(arena_, lhs);
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(SpanRange(lhs_range, rhs_range));
    return n;
  }
  // The evaluator's `expand_range` clamps the unbounded axis to the sheet's
  // used range, and inherits a sheet qualifier on the left endpoint.
  AstNode* lhs_node = make_ref(arena_, lhs);
  AstNode* rhs_node = make_ref(arena_, rhs);
  if (lhs_node == nullptr || rhs_node == nullptr) {
    return nullptr;
  }
  lhs_node->set_range(lhs_range);
  rhs_node->set_range(rhs_range);
  AstNode* n = make_range_op(arena_, lhs_node, rhs_node);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(lhs_range, rhs_range));
  return n;
}

// ---------------------------------------------------------------------------
// Sheet-qualified refs
// ---------------------------------------------------------------------------

bool Parser::defined_name_at(std::size_t offset) const noexcept {
  if (peek_kind_at(offset) != TokenKind::Ident) {
    return false;
  }
  if (peek_kind_at(offset + 1) == TokenKind::Colon && peek_kind_at(offset + 2) == TokenKind::Ident) {
    Reference lhs;
    Reference rhs;
    return !(decode_full_col_endpoint(peek_at(offset), &lhs) && decode_full_col_endpoint(peek_at(offset + 2), &rhs));
  }
  return true;
}

bool Parser::sheet_span_end_at(std::size_t offset) const noexcept {
  const Token& tok = peek_at(offset);
  if (tok.kind != TokenKind::Ident && tok.kind != TokenKind::SheetName && tok.kind != TokenKind::Bool) {
    return false;
  }
  const std::string_view name = tok.kind == TokenKind::SheetName ? tok.text : tok.lexeme;
  return !name.empty() && !IsAsciiDigit(name.front());
}

AstNode* Parser::parse_quoted_qualifier_ref(SyncContext ctx) {
  const Token& tok = peek();
  const std::string_view text = tok.text;
  const std::size_t open = text.rfind('[');
  const bool has_separator = text.find_first_of("/\\") != std::string_view::npos;
  auto reject = [&]() {
    record_error_with_token(ParseErrorCode::InvalidReference, tok.range, tok.lexeme);
    advance();
    skip_to_sync(ctx);
    return make_recovery_placeholder(tok.range);
  };
  auto finish = [&](AstNode* n) {
    if (n != nullptr) {
      return n;
    }
    skip_to_sync(ctx);
    return make_recovery_placeholder(tok.range);
  };

  // `'<path>[Book.xlsx]Sheet'` or the 3-D `'<path>[Book.xlsx]S1:S2'`. The
  // last `[` opens the book, so a directory may itself hold brackets or a
  // drive colon (`'C:\x\[Book.xlsx]Sheet'`); only the sheet part splits on
  // `:`.
  if (open != std::string_view::npos) {
    const std::size_t close = text.find(']', open);
    if (close == std::string_view::npos || close == open + 1U) {
      return reject();
    }
    std::string_view sheet = text.substr(close + 1U);
    std::string_view sheet_end;
    if (const std::size_t colon = sheet.find(':'); colon != std::string_view::npos) {
      sheet_end = sheet.substr(colon + 1U);
      sheet = sheet.substr(0, colon);
      if (sheet_end.empty()) {
        return reject();
      }
    }
    if (sheet.empty() || sheet.find(']') != std::string_view::npos) {
      return reject();
    }
    advance();
    return finish(parse_external_ref_tail(text.substr(0, open), text.substr(open + 1U, close - open - 1U), sheet,
                                          sheet_end, tok.range));
  }

  // `'<path>Book.xlsx'!Name`: a book-scope name of another workbook, told
  // apart from a sheet-local name by a directory or a workbook extension.
  if (peek_kind_at(1) == TokenKind::Bang && defined_name_at(2) && (has_separator || HasWorkbookExtension(text))) {
    const std::size_t split = text.find_last_of("/\\");
    const std::size_t book_begin = split == std::string_view::npos ? 0U : split + 1U;
    if (book_begin == text.size()) {
      return reject();
    }
    advance();
    return finish(parse_external_ref_tail(text.substr(0, book_begin), text.substr(book_begin), {}, {}, tok.range));
  }

  // A sheet name cannot hold a path separator, and a directory with no
  // bracketed book names no sheet to take cells from.
  if (has_separator) {
    return reject();
  }

  // 3-D reference whose first endpoint is a quoted sheet name
  // (`'My Sheet':Sheet3!A1`).
  if (peek_kind_at(1) == TokenKind::Colon && sheet_span_end_at(2) && peek_kind_at(3) == TokenKind::Bang) {
    advance();
    return finish(parse_3d_ref(text, tok.range));
  }
  advance();
  return finish(parse_sheet_qualified_ref(text, /*quoted=*/true, tok.range));
}

AstNode* Parser::parse_external_ref_tail(std::string_view path, std::string_view book, std::string_view sheet,
                                         std::string_view sheet_end, TextRange start_range) {
  if (peek_kind() != TokenKind::Bang) {
    record_error_with_token(ParseErrorCode::UnexpectedToken, peek().range, peek().lexeme);
    return nullptr;
  }
  advance();  // Bang

  // `[0]` is the formula's own workbook, which Excel only writes in the
  // name form (`[0]!Name`, the storage spelling of `Book!Name`).
  if (book == "0" && path.empty() && !sheet.empty()) {
    record_error_with_token(ParseErrorCode::InvalidReference, start_range, "[0]");
    return nullptr;
  }

  // `Book.xlsx!Name` / `[0]!Name`: a book-scope defined name, or with a
  // single sheet (`[Book.xlsx]Data!Rate`) one local to that sheet.
  if (sheet.empty() || (sheet_end.empty() && defined_name_at(0))) {
    if (peek_kind() != TokenKind::Ident) {
      record_error_with_token(ParseErrorCode::InvalidReference, peek().range, peek().lexeme);
      return nullptr;
    }
    const Token& name = advance();
    AstNode* node = make_external_name_ref(arena_, path, book, sheet, name.lexeme);
    if (node == nullptr) {
      return nullptr;
    }
    node->set_range(SpanRange(start_range, name.range));
    return node;
  }

  // A cell, a rectangle, or a whole column / row (`[Book.xlsx]Data!A:A`),
  // on one sheet or across the span.
  Reference first;
  Reference last;
  bool is_range = false;
  TextRange tail_range;
  if (!parse_3d_ref_tail(&first, &last, &is_range, &tail_range)) {
    return nullptr;
  }
  AstNode* node = make_external_ref(arena_, path, book, sheet, sheet_end, first, last, is_range);
  if (node == nullptr) {
    return nullptr;
  }
  node->set_range(SpanRange(start_range, tail_range));
  return node;
}

AstNode* Parser::parse_sheet_qualified_ref(std::string_view sheet, bool quoted, TextRange sheet_range) {
  // Expect Bang next. (For SheetName tokens we have not yet consumed Bang;
  // for unquoted Ident the caller has consumed only the Ident.)
  if (peek_kind() != TokenKind::Bang) {
    record_error_with_token(ParseErrorCode::UnexpectedToken, peek().range, peek().lexeme);
    return nullptr;
  }
  advance();  // Bang

  // Quoted 3-D sheet range (`'Data:S2'!B1`): a `:` inside the sheet
  // qualifier is the sheet-range separator, because a worksheet name can
  // never itself contain a colon. The tokenizer delivers the whole
  // `Data:S2` span as one (quoted) SheetName, so split it here into the
  // begin / end sheet names and build a `Ref3D`. A single-cell tail
  // (`'Data:S2'!B1`) builds a single-cell Ref3D; a range tail
  // (`'Data:S2'!A1:B2`) builds a range Ref3D. (The unquoted
  // `Sheet1:Sheet2!A1` shape arrives as separate tokens and is handled by
  // `parse_3d_ref`.)
  if (const std::size_t colon = sheet.find(':'); colon != std::string_view::npos) {
    const std::string_view sheet_begin = sheet.substr(0, colon);
    const std::string_view sheet_end = sheet.substr(colon + 1);
    if (sheet_begin.empty() || sheet_end.empty()) {
      record_error_with_token(ParseErrorCode::InvalidReference, peek().range, peek().lexeme);
      return nullptr;
    }
    Reference first;
    Reference last;
    bool is_range = false;
    TextRange tail_range;
    if (!parse_3d_ref_tail(&first, &last, &is_range, &tail_range)) {
      return nullptr;
    }
    AstNode* n = is_range ? make_ref3d_range(arena_, sheet_begin, sheet_end, first, last)
                          : make_ref3d(arena_, sheet_begin, sheet_end, first);
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(SpanRange(sheet_range, tail_range));
    return n;
  }

  // Five possibilities:
  //   1. CellRef: `Sheet1!A1`.
  //   2. Ident Colon Ident with matching column letters: `Sheet1!A:A`.
  //   3. Number Colon Number with matching row digits: `Sheet1!1:1`.
  //   4. `#REF!`: Excel's spelling after the sheet itself is deleted
  //      (`Sheet1!#REF!`). The whole reference has already collapsed to
  //      one error, so the sheet qualifier carries no surviving meaning.
  //   5. Any other Ident: a defined name looked up in that sheet's scope
  //      (`Sheet1!Rate`), the only spelling that reaches another sheet's
  //      local name. A following `(` calls it (`Sheet1!Fn(2)`); the
  //      postfix-call rule in the Pratt loop wraps it in a `LambdaCall`.
  // Anything else is an error.
  const TokenKind k = peek_kind();
  if (k == TokenKind::ErrorLiteral && peek().error_code == ErrorCode::Ref) {
    const Token& err_tok = advance();
    const TextRange range = consume_ref_error_glued_tail(SpanRange(sheet_range, err_tok.range));
    AstNode* n = make_error_literal(arena_, ErrorCode::Ref);
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(range);
    return n;
  }
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
  // `Sheet1!A:C` / `Sheet1!1:3`. The sheet qualifier stays on the left
  // endpoint, matching how `Sheet1!A1:B2` parses. Columns that fail to
  // decode fall through to the defined-name case below.
  const bool col_pair =
      k == TokenKind::Ident && peek_kind_at(1) == TokenKind::Colon && peek_kind_at(2) == TokenKind::Ident;
  const bool row_pair =
      k == TokenKind::Number && peek_kind_at(1) == TokenKind::Colon && peek_kind_at(2) == TokenKind::Number;
  if (col_pair || row_pair) {
    const Token& lhs_tok = peek();
    const Token& rhs_tok = peek_at(2);
    Reference lhs_ref;
    Reference rhs_ref;
    const bool decoded =
        col_pair ? decode_full_col_endpoint(lhs_tok, &lhs_ref) && decode_full_col_endpoint(rhs_tok, &rhs_ref)
                 : decode_full_row_endpoint(lhs_tok, &lhs_ref) && decode_full_row_endpoint(rhs_tok, &rhs_ref);
    if (decoded) {
      const TextRange lhs_range = SpanRange(sheet_range, lhs_tok.range);
      const TextRange rhs_range = rhs_tok.range;
      advance();
      advance();
      advance();
      lhs_ref.sheet = sheet;
      lhs_ref.sheet_quoted = quoted;
      return make_whole_axis_ref(lhs_ref, rhs_ref, lhs_range, rhs_range);
    }
  }
  if (k == TokenKind::Ident) {
    const Token& name = advance();
    AstNode* n = make_sheet_name_ref(arena_, sheet, name.lexeme, quoted);
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(SpanRange(sheet_range, name.range));
    return n;
  }
  record_error_with_token(ParseErrorCode::InvalidReference, peek().range, peek().lexeme);
  return nullptr;
}

// ---------------------------------------------------------------------------
// 3-D references
// ---------------------------------------------------------------------------

bool Parser::parse_3d_ref_tail(Reference* first, Reference* last, bool* is_range, TextRange* tail_range) {
  // Whole references are represented with the same flags as their single-
  // sheet counterparts; evaluation then expands each sheet in the span to
  // that sheet's populated extent.
  *is_range = false;
  if (peek_kind() == TokenKind::CellRef) {
    const Token& cell = advance();
    *tail_range = cell.range;
    if (!decode_cellref_lexeme(cell.lexeme, first)) {
      record_error_with_token(ParseErrorCode::InvalidReference, cell.range, cell.lexeme);
      return false;
    }
    if (peek_kind() == TokenKind::Colon && peek_kind_at(1) == TokenKind::CellRef) {
      advance();  // Colon
      const Token& tail = advance();
      *tail_range = tail.range;
      if (!decode_cellref_lexeme(tail.lexeme, last)) {
        record_error_with_token(ParseErrorCode::InvalidReference, tail.range, tail.lexeme);
        return false;
      }
      *is_range = true;
    }
    return true;
  }
  const bool col_pair =
      peek_kind() == TokenKind::Ident && peek_kind_at(1) == TokenKind::Colon && peek_kind_at(2) == TokenKind::Ident;
  const bool row_pair =
      peek_kind() == TokenKind::Number && peek_kind_at(1) == TokenKind::Colon && peek_kind_at(2) == TokenKind::Number;
  if (!col_pair && !row_pair) {
    record_error_with_token(ParseErrorCode::InvalidReference, peek().range, peek().lexeme);
    return false;
  }
  const Token& lhs = peek();
  const Token& rhs = peek_at(2);
  const bool decoded = col_pair ? decode_full_col_endpoint(lhs, first) && decode_full_col_endpoint(rhs, last)
                                : decode_full_row_endpoint(lhs, first) && decode_full_row_endpoint(rhs, last);
  if (!decoded) {
    record_error_with_token(ParseErrorCode::InvalidReference, lhs.range, lhs.lexeme);
    return false;
  }
  *tail_range = rhs.range;
  advance();
  advance();
  advance();
  *is_range = first->col != last->col || first->row != last->row;
  return true;
}

AstNode* Parser::parse_3d_ref(std::string_view sheet1, TextRange sheet1_range) {
  // The caller guarantees the current token is `Colon`, the next is an
  // `Ident` / `SheetName`, and the one after that is `Bang`.
  advance();  // Colon
  const Token& sheet2_tok = peek();
  std::string_view sheet2;
  if (sheet2_tok.kind == TokenKind::SheetName) {
    sheet2 = sheet2_tok.text;  // escape-resolved
  } else {
    sheet2 = sheet2_tok.lexeme;
  }
  advance();  // second sheet name
  advance();  // Bang

  Reference r;
  Reference r2;
  bool has_range_tail = false;
  TextRange tail_range;
  if (!parse_3d_ref_tail(&r, &r2, &has_range_tail, &tail_range)) {
    return nullptr;
  }
  // Range tail (`Sheet1:Sheet2!A1:B2`): a 3-D range over the sheet span,
  // built as a range `Ref3D`. Without this the outer Pratt `:` rule would
  // otherwise mis-assemble `RangeOp(Ref3D(A1), Ref(B2))`.
  if (!has_range_tail && peek_kind() == TokenKind::Colon && peek_kind_at(1) == TokenKind::CellRef) {
    advance();  // Colon
    const Token& tail = advance();
    if (!decode_cellref_lexeme(tail.lexeme, &r2)) {
      record_error_with_token(ParseErrorCode::InvalidReference, tail.range, tail.lexeme);
      return nullptr;
    }
    AstNode* range_node = make_ref3d_range(arena_, sheet1, sheet2, r, r2);
    if (range_node == nullptr) {
      return nullptr;
    }
    range_node->set_range(SpanRange(sheet1_range, tail.range));
    return range_node;
  }
  if (has_range_tail) {
    AstNode* range_node = make_ref3d_range(arena_, sheet1, sheet2, r, r2);
    if (range_node == nullptr) {
      return nullptr;
    }
    range_node->set_range(SpanRange(sheet1_range, tail_range));
    return range_node;
  }
  AstNode* n = make_ref3d(arena_, sheet1, sheet2, r);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(sheet1_range, tail_range));
  return n;
}

}  // namespace parser
}  // namespace formulon
