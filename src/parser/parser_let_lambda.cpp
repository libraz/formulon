//
// LET / LAMBDA special-form parsing: binding-name validation and the
// `parse_let_call` / `parse_lambda_call` Parser members.

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/parse_error.h"
#include "parser/parser.h"
#include "parser/parser_detail.h"
#include "parser/token.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {

using detail::IsAsciiDigit;
using detail::IsAsciiLetter;
using detail::SpanRange;

namespace {

// Returns true iff `name` starts with `[A-Za-z_\]` or any non-ASCII (UTF-8
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
  if (!IsAsciiLetter(static_cast<char>(first)) && first != '_' && first != '\\' && first < 0x80) {
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

}  // namespace parser
}  // namespace formulon
