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
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace parser {

using detail::IsAsciiDigit;
using detail::kBpAtPrefix;
using detail::kBpUnaryPrefix;
using detail::kMaxRow;
using detail::SpanRange;

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
  // Top-level parenthesised comma list = reference union. Excel's
  // `=AREAS((A1,B2))` and `=SUM((A1:B2,C3:D4))` rely on this. The comma
  // here cannot be confused with a function-arg separator because that is
  // consumed inside the call parser; reaching `parse_paren_atom` means we
  // are in a grouped-expression context.
  if (peek_kind() == TokenKind::Comma) {
    std::vector<const AstNode*> children;
    children.push_back(inner);
    while (peek_kind() == TokenKind::Comma) {
      advance();  // consume `,`
      AstNode* next_child = parse_expression(0, SyncContext::Paren);
      if (next_child == nullptr) {
        return nullptr;
      }
      children.push_back(next_child);
    }
    if (peek_kind() != TokenKind::RParen) {
      record_error_with_token(ParseErrorCode::ExpectedCloseParen, lparen.range, lparen.lexeme);
      // Best-effort recovery: emit the union we have so far. Children are
      // already valid subtrees (possibly placeholders); make_union_op can
      // build a node from any vector with size >= 2.
      AstNode* u = make_union_op(arena_, children.data(), static_cast<std::uint32_t>(children.size()));
      if (u != nullptr) {
        u->set_range(lparen.range);
      }
      return u;
    }
    const Token& rparen = advance();
    AstNode* u = make_union_op(arena_, children.data(), static_cast<std::uint32_t>(children.size()));
    if (u == nullptr) {
      return nullptr;
    }
    u->set_range(SpanRange(lparen.range, rparen.range));
    return u;
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
      case TokenKind::String: {
        if (have_prefix) {
          // Unary -/+ is only meaningful on numeric literals inside an
          // array; a prefix on a string surfaces as a syntax error. Mirror
          // the recovery in the `default:` arm so siblings keep parsing.
          const Token& tok = peek();
          record_error_with_token(ParseErrorCode::ExpectedExpression, tok.range, tok.lexeme);
          skip_to_sync(SyncContext::ArrayElem);
          AstNode* placeholder = make_error_placeholder(arena_);
          if (placeholder != nullptr) {
            placeholder->set_range(tok.range);
          }
          return placeholder;
        }
        // The tokenizer's `text` field is the escape-resolved payload,
        // already interned in the tokenizer arena, so we can hand it
        // straight to `Value::text` without copying — same pattern as
        // `parse_string_atom` for the top-level grammar.
        const Token& st = advance();
        node = make_literal(arena_, Value::text(st.text));
        if (node != nullptr) {
          node->set_range(st.range);
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
        // Empty argument slot: Excel allows `FN(a,,b)` and treats the missing
        // slot as blank / the function's documented default. Inject a Blank
        // literal rather than invoking parse_expression (which has no grammar
        // production for an empty arg).
        AstNode* arg = nullptr;
        const TokenKind here = peek_kind();
        if (here == TokenKind::Comma || here == TokenKind::RParen) {
          arg = make_literal(arena_, Value::blank());
          if (arg != nullptr) {
            arg->set_range(peek().range);
          }
        } else {
          arg = parse_expression(0, SyncContext::CallArg);
        }
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

}  // namespace parser
}  // namespace formulon
