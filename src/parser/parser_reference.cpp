
#include <cstdint>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/parse_error.h"
#include "parser/parser.h"
#include "parser/parser_detail.h"
#include "parser/reference.h"
#include "parser/token.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {

using detail::DecodeDigitRunClamped;
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
    const bool ok = IsAsciiLetter(c) || IsAsciiDigit(c) || c == '_' || c == '.' || c == '?' || byte >= 0x80;
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
    return make_recovery_placeholder(SpanRange(call_start, rparen_empty.range));
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
    return make_recovery_placeholder(SpanRange(call_start, end_range));
  }

  AstNode* n = make_let_binding(arena_, names.data(), exprs.data(), static_cast<std::uint32_t>(names.size()), body);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(call_start, end_range));
  return n;
}

// ---------------------------------------------------------------------------
// LAMBDA parameters
// ---------------------------------------------------------------------------

AstNode* Parser::parse_lambda_call(const Token& name_tok) {
  const TextRange call_start = name_tok.range;
  advance();  // LAMBDA Ident
  advance();  // LParen

  // Collect parameter names then the body. Slot classification: each
  // non-final slot must be a bare Ident shape that passes the same
  // identifier rules used for LET binding names; the final slot is the
  // body. Excel's grammar requires at least one slot total — `LAMBDA()` is
  // an error. A single-slot form `LAMBDA(expr)` matches Mac Excel by
  // treating `expr` as the body of a zero-parameter lambda.
  //
  // Parameters introduced with bracket syntax `[name]` are optional. They
  // must be trailing (no required parameter may follow an optional one);
  // when omitted at the call site they bind to a sentinel that ISOMITTED
  // detects.
  std::vector<std::string_view> params;
  std::uint32_t optional_count = 0;
  AstNode* body = nullptr;

  // Empty arg list: LAMBDA requires the body slot.
  if (peek_kind() == TokenKind::RParen) {
    record_error_with_token(ParseErrorCode::LambdaEmpty, name_tok.range, name_tok.lexeme);
    const Token& rparen_empty = advance();
    return make_recovery_placeholder(SpanRange(call_start, rparen_empty.range));
  }

  // Slot-walk loop. Each iteration starts at the first token of the next
  // slot. A slot is a parameter name iff it is a well-shaped Ident followed
  // by a comma; otherwise it is the body. This matches the LET disambiguation
  // pattern: the body is always the final slot (never followed by a comma).
  while (true) {
    if (bailed_) {
      break;
    }
    // CellRef-shaped tokens (e.g. `A1`, `AA10`) sitting in a parameter slot
    // collide with the A1 cell they spell; emit the dedicated diagnostic so
    // siblings keep parsing.
    if (peek_kind() == TokenKind::CellRef && peek_kind_at(1) == TokenKind::Comma) {
      const Token& bad = peek();
      record_error_with_token(ParseErrorCode::LambdaInvalidParam, bad.range, bad.lexeme);
      advance();  // consume the cell-ref
      if (peek_kind() == TokenKind::Comma) {
        advance();
      }
      continue;
    }
    // Bracketed optional-parameter shape: `[name]` followed by either a
    // comma (more slots ahead) or `)` (this was actually the body slot —
    // a `[ref]` structured reference inside the body parses there). We
    // only treat it as a param when the closing bracket is followed by a
    // comma; otherwise let the body branch handle it.
    const bool is_optional_param_slot = (peek_kind() == TokenKind::LBracket) && (peek_kind_at(1) == TokenKind::Ident) &&
                                        IsLetNameShape(peek_at(1).lexeme) && !LooksLikeCellRef(peek_at(1).lexeme) &&
                                        (peek_kind_at(2) == TokenKind::RBracket) &&
                                        (peek_kind_at(3) == TokenKind::Comma);
    const bool is_param_slot = (peek_kind() == TokenKind::Ident) && IsLetNameShape(peek().lexeme) &&
                               !LooksLikeCellRef(peek().lexeme) && peek_kind_at(1) == TokenKind::Comma;
    if (is_optional_param_slot || is_param_slot) {
      const bool optional = is_optional_param_slot;
      const Token& tok = optional ? peek_at(1) : peek();
      const std::string_view pname = tok.lexeme;
      // Reject duplicates within a single LAMBDA: the second occurrence
      // would shadow the first at runtime, which is almost certainly a bug
      // and which Excel itself rejects.
      bool duplicate = false;
      for (const auto& existing : params) {
        if (strings::case_insensitive_eq(existing, pname)) {
          duplicate = true;
          break;
        }
      }
      if (duplicate) {
        record_error_with_token(ParseErrorCode::LambdaDuplicateParam, tok.range, tok.lexeme);
        if (optional) {
          advance();  // LBracket
          advance();  // Ident
          advance();  // RBracket
        } else {
          advance();  // Ident
        }
        if (peek_kind() == TokenKind::Comma) {
          advance();
        }
        continue;
      }
      params.push_back(pname);
      if (optional) {
        ++optional_count;
        advance();  // LBracket
        advance();  // Ident
        advance();  // RBracket
        advance();  // Comma
      } else {
        // A required parameter is illegal once we've started accepting
        // optional ones: Excel only allows trailing optionals.
        if (optional_count > 0) {
          record_error_with_token(ParseErrorCode::LambdaInvalidParam, tok.range, tok.lexeme);
          // Treat as optional anyway so we don't lose the slot — but the
          // diagnostic is what matters; the AST will still produce a name.
          ++optional_count;
        }
        advance();  // Ident
        advance();  // Comma
      }
      continue;
    }
    // Non-param-slot: this is either the body (final slot, followed by `)`)
    // or a malformed param slot (not an Ident, but followed by a comma).
    // Detect the malformed case by looking ahead: if the slot is *not* the
    // final one (i.e. there will be a comma after the expression) and the
    // current token is not a valid bare-Ident param shape, surface the
    // dedicated diagnostic before parsing the slot's expression.
    //
    // This catches `LAMBDA(x, x+1, y)`: the middle slot is `x+1` which is
    // not a bare Ident-Comma shape but is followed by a comma after the
    // expression closes.
    //
    // We cannot know the "followed by comma after expression" answer
    // without parsing first, so the strategy is: tentatively parse the
    // slot as an expression; if a comma follows, we wrongly admitted a
    // non-Ident param slot — emit the diagnostic and discard the
    // expression. If `)` follows, the expression is the body.
    AstNode* slot = parse_expression(0, SyncContext::CallArg);
    if (slot == nullptr) {
      return nullptr;  // hard arena failure
    }
    if (peek_kind() == TokenKind::Comma) {
      // The slot we just parsed was not the final one, but it was not a
      // bare-Ident param shape either: the user wrote something like
      // `LAMBDA(x, x+1, y)` where `x+1` is illegal as a parameter name.
      record_error_with_token(ParseErrorCode::LambdaInvalidParam, slot->range(), std::string_view{});
      advance();  // consume the comma so the next iteration starts cleanly
      continue;
    }
    // Final slot: this is the body.
    body = slot;
    break;
  }

  // Expect `)`; emit a diagnostic and recover otherwise.
  TextRange end_range = call_start;
  if (peek_kind() == TokenKind::RParen) {
    const Token& rparen = advance();
    end_range = rparen.range;
  } else {
    record_error_with_token(ParseErrorCode::ExpectedCloseParen, call_start, name_tok.lexeme);
  }

  if (body == nullptr) {
    record_error_with_token(ParseErrorCode::LambdaEmpty, call_start, name_tok.lexeme);
    return make_recovery_placeholder(SpanRange(call_start, end_range));
  }

  AstNode* n = make_lambda(arena_, params.empty() ? nullptr : params.data(), static_cast<std::uint32_t>(params.size()),
                           optional_count, body);
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
    if (sheet_begin.empty() || sheet_end.empty() || peek_kind() != TokenKind::CellRef) {
      record_error_with_token(ParseErrorCode::InvalidReference, peek().range, peek().lexeme);
      return nullptr;
    }
    const Token& cell = advance();
    Reference r;
    if (!decode_cellref_lexeme(cell.lexeme, &r)) {
      record_error_with_token(ParseErrorCode::InvalidReference, cell.range, cell.lexeme);
      return nullptr;
    }
    // Range tail (`'Data:S2'!A1:B2`): a 3-D range over the sheet span.
    if (peek_kind() == TokenKind::Colon && peek_kind_at(1) == TokenKind::CellRef) {
      advance();  // Colon
      const Token& tail = advance();
      Reference r2;
      if (!decode_cellref_lexeme(tail.lexeme, &r2)) {
        record_error_with_token(ParseErrorCode::InvalidReference, tail.range, tail.lexeme);
        return nullptr;
      }
      AstNode* range_node = make_ref3d_range(arena_, sheet_begin, sheet_end, r, r2);
      if (range_node == nullptr) {
        return nullptr;
      }
      range_node->set_range(SpanRange(sheet_range, tail.range));
      return range_node;
    }
    AstNode* n = make_ref3d(arena_, sheet_begin, sheet_end, r);
    if (n == nullptr) {
      return nullptr;
    }
    n->set_range(SpanRange(sheet_range, cell.range));
    return n;
  }

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
    if (lhs_col != 0 && rhs_col != 0) {
      const TextRange end_range = rhs_tok.range;
      advance();
      advance();
      advance();
      Reference lhs_ref;
      lhs_ref.col = lhs_col - 1;
      lhs_ref.col_abs = lhs_abs;
      lhs_ref.is_full_col = true;
      lhs_ref.sheet = sheet;
      lhs_ref.sheet_quoted = quoted;
      if (lhs_col == rhs_col) {
        // `Sheet1!A:A`: a single whole-column Ref.
        AstNode* n = make_ref(arena_, lhs_ref);
        if (n == nullptr) {
          return nullptr;
        }
        n->set_range(SpanRange(sheet_range, end_range));
        return n;
      }
      // `Sheet1!A:C`: a range spanning two whole-column Refs. The sheet
      // qualifier stays on the left endpoint, matching how `Sheet1!A1:B2`
      // parses; `expand_range` inherits it for the whole rectangle.
      Reference rhs_ref;
      rhs_ref.col = rhs_col - 1;
      rhs_ref.col_abs = rhs_abs;
      rhs_ref.is_full_col = true;
      AstNode* lhs_node = make_ref(arena_, lhs_ref);
      AstNode* rhs_node = make_ref(arena_, rhs_ref);
      if (lhs_node == nullptr || rhs_node == nullptr) {
        return nullptr;
      }
      lhs_node->set_range(SpanRange(sheet_range, lhs_tok.range));
      rhs_node->set_range(rhs_tok.range);
      AstNode* n = make_range_op(arena_, lhs_node, rhs_node);
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
    // Each endpoint may carry a leading `$` (absolute row anchor); the
    // tokenizer keeps the `$` in the Number lexeme, so strip it here and
    // record the row_abs flag. `DecodeDigitRunClamped` stops accumulating
    // once past `kMaxRow`, so a pathological 20-digit row literal cannot
    // wrap a `std::uint64_t` back into the valid range.
    auto decode_row = [](std::string_view lex, bool* row_abs, std::uint64_t* out_row) -> bool {
      *row_abs = false;
      if (!lex.empty() && lex.front() == '$') {
        *row_abs = true;
        lex.remove_prefix(1);
      }
      if (lex.empty()) {
        return false;
      }
      for (char c : lex) {
        if (!IsAsciiDigit(c)) {
          return false;
        }
      }
      *out_row = DecodeDigitRunClamped(lex, kMaxRow);
      return true;
    };
    bool lhs_abs = false;
    bool rhs_abs = false;
    std::uint64_t lhs_row = 0;
    std::uint64_t rhs_row = 0;
    if (lhs_tok.is_integer && rhs_tok.is_integer && decode_row(lhs_tok.lexeme, &lhs_abs, &lhs_row) &&
        decode_row(rhs_tok.lexeme, &rhs_abs, &rhs_row) && lhs_row != 0 && rhs_row != 0 && lhs_row <= kMaxRow &&
        rhs_row <= kMaxRow) {
      {
        const TextRange end_range = rhs_tok.range;
        advance();
        advance();
        advance();
        Reference lhs_ref;
        lhs_ref.row = static_cast<std::uint32_t>(lhs_row - 1);
        lhs_ref.row_abs = lhs_abs;
        lhs_ref.is_full_row = true;
        lhs_ref.sheet = sheet;
        lhs_ref.sheet_quoted = quoted;
        if (lhs_row == rhs_row) {
          // `Sheet1!1:1`: a single whole-row Ref.
          AstNode* n = make_ref(arena_, lhs_ref);
          if (n == nullptr) {
            return nullptr;
          }
          n->set_range(SpanRange(sheet_range, end_range));
          return n;
        }
        // `Sheet1!1:3`: a range spanning two whole-row Refs. The sheet
        // qualifier stays on the left endpoint; `expand_range` inherits it.
        Reference rhs_ref;
        rhs_ref.row = static_cast<std::uint32_t>(rhs_row - 1);
        rhs_ref.row_abs = rhs_abs;
        rhs_ref.is_full_row = true;
        AstNode* lhs_node = make_ref(arena_, lhs_ref);
        AstNode* rhs_node = make_ref(arena_, rhs_ref);
        if (lhs_node == nullptr || rhs_node == nullptr) {
          return nullptr;
        }
        lhs_node->set_range(SpanRange(sheet_range, lhs_tok.range));
        rhs_node->set_range(rhs_tok.range);
        AstNode* n = make_range_op(arena_, lhs_node, rhs_node);
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

// ---------------------------------------------------------------------------
// 3-D references
// ---------------------------------------------------------------------------

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

  // Only a single cell reference is supported as the 3-D tail; whole-column /
  // whole-row 3-D ranges are out of scope.
  if (peek_kind() != TokenKind::CellRef) {
    record_error_with_token(ParseErrorCode::InvalidReference, peek().range, peek().lexeme);
    return nullptr;
  }
  const Token& cell = advance();
  Reference r;
  if (!decode_cellref_lexeme(cell.lexeme, &r)) {
    record_error_with_token(ParseErrorCode::InvalidReference, cell.range, cell.lexeme);
    return nullptr;
  }
  // Range tail (`Sheet1:Sheet2!A1:B2`): a 3-D range over the sheet span,
  // built as a range `Ref3D`. Without this the outer Pratt `:` rule would
  // otherwise mis-assemble `RangeOp(Ref3D(A1), Ref(B2))`.
  if (peek_kind() == TokenKind::Colon && peek_kind_at(1) == TokenKind::CellRef) {
    advance();  // Colon
    const Token& tail = advance();
    Reference r2;
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
  AstNode* n = make_ref3d(arena_, sheet1, sheet2, r);
  if (n == nullptr) {
    return nullptr;
  }
  n->set_range(SpanRange(sheet1_range, cell.range));
  return n;
}

}  // namespace parser
}  // namespace formulon
