//
// Implementation of `SourceWhitespace`. See `io/xlsb/ptg_whitespace.h`.

#include "io/xlsb/ptg_whitespace.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/token.h"
#include "parser/tokenizer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace detail {

void SourceWhitespace::visit(const parser::AstNode& node, bool lead_allowed) {
  // A node's own `(` opens a position whitespace may follow.
  const bool inside = lead_allowed || node.paren_depth() != 0U;
  switch (node.kind()) {
    case parser::NodeKind::Literal:
      if (!node.as_literal().is_blank()) {
        lead(node, inside);
      }
      break;
    case parser::NodeKind::Ref:
    case parser::NodeKind::Ref3D:
    case parser::NodeKind::ExternalRef:
    case parser::NodeKind::NameRef:
    case parser::NodeKind::ErrorLiteral:
    case parser::NodeKind::ArrayLiteral:
      lead(node, inside);
      break;
    case parser::NodeKind::UnaryOp:
      if (node.as_unary_op() == parser::UnaryOp::Percent) {
        visit(node.as_unary_operand(), inside);
        before_operator(node, node.as_unary_operand(), parser::TokenKind::Percent);
      } else {
        lead(node, inside);
        visit(node.as_unary_operand(), true);
      }
      break;
    case parser::NodeKind::BinaryOp:
      visit(node.as_binary_operand(0), inside);
      for (std::uint32_t i = 1; i < node.as_binary_arity(); ++i) {
        before_operator(node, node.as_binary_operand(i - 1U), parser::TokenKind::Eof,
                        static_cast<std::uint16_t>(i - 1U));
        visit(node.as_binary_operand(i), true);
      }
      break;
    case parser::NodeKind::RangeOp:
      for (std::uint32_t i = 0; i < node.as_range_arity(); ++i) {
        visit(node.as_range_operand(i), i == 0U ? inside : false);
      }
      break;
    case parser::NodeKind::IntersectOp:
      visit(node.as_intersect_lhs(), inside);
      intersection(node);
      visit(node.as_intersect_rhs(), false);
      break;
    case parser::NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        visit(node.as_union_child(i), i == 0U ? inside : true);
      }
      break;
    case parser::NodeKind::Call:
      if (parser::trim_ref_call_mode(node) != parser::TrimRefMode::None) {
        visit(node.as_call_arg(0), inside);
        break;
      }
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        visit(node.as_call_arg(i), true);
      }
      call_close(node);
      break;
    case parser::NodeKind::SpillRef:
      if (const parser::AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        visit(*anchor, inside);
      }
      break;
    case parser::NodeKind::LetBinding:
      // Binding names take no whitespace; values and the body follow a `,`.
      for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
        visit(node.as_let_binding_expr(i), true);
      }
      visit(node.as_let_body(), true);
      break;
    case parser::NodeKind::ImplicitIntersection:
    case parser::NodeKind::Lambda:
    case parser::NodeKind::LambdaCall:
    case parser::NodeKind::StructuredRef:
    case parser::NodeKind::ErrorPlaceholder:
      parser::any_child_node(node, [this](const parser::AstNode& child) {
        visit(child, false);
        return false;
      });
      break;
  }
  paren_opens(node, lead_allowed);
  paren_closes(node);
}

// Index of the first token starting at or after `offset`.
std::size_t SourceWhitespace::token_at(std::uint32_t offset) const {
  const auto it = std::lower_bound(tokens_->begin(), tokens_->end(), offset,
                                   [](const parser::Token& t, std::uint32_t at) { return t.range.start < at; });
  return static_cast<std::size_t>(it - tokens_->begin());
}

// Index of the token ending `node`'s text, its own parentheses included.
std::size_t SourceWhitespace::last_token(const parser::AstNode& node) const {
  const std::size_t next = token_at(node.range().end);
  if (next == 0U || (*tokens_)[next - 1U].range.end != node.range().end) {
    return kNone;
  }
  return next - 1U;
}

// Index of the first token of `node`'s text inside its own parentheses.
std::size_t SourceWhitespace::first_bare_token(const parser::AstNode& node) const {
  std::size_t i = token_at(node.range().start);
  for (unsigned d = 0; d < node.paren_depth(); ++d) {
    if (!is(i, parser::TokenKind::LParen)) {
      return kNone;
    }
    ++i;
    if (is(i, parser::TokenKind::Whitespace)) {
      ++i;
    }
  }
  return i < tokens_->size() ? i : kNone;
}

// Records the whitespace token `i`, less `skip` characters, as `slot` of
// `node` when it is a run of spaces or, if allowed, of line feeds, or line
// feeds then spaces.
void SourceWhitespace::record(std::size_t i, const parser::AstNode& node, parser::WhitespaceSlot slot,
                              std::uint16_t index, bool line_feeds_allowed, std::size_t skip) {
  if (!is(i, parser::TokenKind::Whitespace)) {
    return;
  }
  std::string_view run = (*tokens_)[i].lexeme;
  if (run.size() <= skip) {
    return;
  }
  run.remove_prefix(skip);
  const std::size_t feeds = line_feeds_allowed ? std::min(run.find_first_not_of('\n'), run.size()) : 0U;
  const std::size_t spaces = run.size() - feeds;
  if (run.find_first_not_of(' ', feeds) != std::string_view::npos || feeds > 0xFFU || spaces > 0xFFU) {
    return;
  }
  if (feeds != 0U) {
    out_.push_back({&node, slot, index, '\n', static_cast<std::uint8_t>(feeds)});
  }
  if (spaces != 0U) {
    out_.push_back({&node, slot, index, ' ', static_cast<std::uint8_t>(spaces)});
  }
}

void SourceWhitespace::lead(const parser::AstNode& node, bool allowed) {
  const std::size_t first = first_bare_token(node);
  if (!allowed || first == kNone || first == 0U) {
    return;
  }
  record(first - 1U, node, parser::WhitespaceSlot::kBeforeText, 0, true);
}

// `op` is the operator token expected after the run (`Eof`: any binary
// one); `index` is a binary operator's position in its chain.
void SourceWhitespace::before_operator(const parser::AstNode& node, const parser::AstNode& lhs, parser::TokenKind op,
                                       std::uint16_t index) {
  const std::size_t i = token_at(lhs.range().end);
  if (op != parser::TokenKind::Eof && !is(i + 1U, op)) {
    return;
  }
  record(i, node, parser::WhitespaceSlot::kBeforeOperator, index, true);
}

void SourceWhitespace::intersection(const parser::AstNode& node) {
  const std::size_t i = token_at(node.as_intersect_lhs().range().end);
  if (i + 1U >= tokens_->size() || (*tokens_)[i + 1U].range.start != node.as_intersect_rhs().range().start) {
    return;
  }
  record(i, node, parser::WhitespaceSlot::kBeforeOperator, 0, false, /*skip=*/1);
}

// A run before a `)` closing after an operand, not after `(` or `,`.
void SourceWhitespace::close_run(std::size_t rparen, const parser::AstNode& node, std::uint8_t index) {
  if (rparen < 2U || !is(rparen - 1U, parser::TokenKind::Whitespace) || is(rparen - 2U, parser::TokenKind::Comma) ||
      is(rparen - 2U, parser::TokenKind::LParen)) {
    return;
  }
  record(rparen - 1U, node, parser::WhitespaceSlot::kBeforeClose, index, true);
}

// Spaces before each `(` of the node's own pairs; before the outermost one
// only where whitespace may precede an operand.
void SourceWhitespace::paren_opens(const parser::AstNode& node, bool lead_allowed) {
  std::size_t i = token_at(node.range().start);
  for (unsigned level = node.paren_depth(); level >= 1U; --level) {
    if (!is(i, parser::TokenKind::LParen)) {
      return;
    }
    if (i != 0U && (lead_allowed || level != node.paren_depth())) {
      record(i - 1U, node, parser::WhitespaceSlot::kBeforeOpen, static_cast<std::uint8_t>(level), false);
    }
    ++i;
    if (is(i, parser::TokenKind::Whitespace)) {
      ++i;
    }
  }
}

void SourceWhitespace::paren_closes(const parser::AstNode& node) {
  std::size_t k = node.paren_depth() != 0U ? last_token(node) : kNone;
  for (unsigned level = node.paren_depth(); level >= 1U && k != kNone; --level) {
    if (!is(k, parser::TokenKind::RParen)) {
      return;
    }
    close_run(k, node, static_cast<std::uint8_t>(level));
    k = k >= 1U && is(k - 1U, parser::TokenKind::Whitespace) ? k - 1U : k;
    k = k >= 1U ? k - 1U : kNone;
  }
}

void SourceWhitespace::call_close(const parser::AstNode& node) {
  const std::uint32_t arity = node.as_call_arity();
  if (arity == 0U) {
    // Whitespace inside an empty call: name, `(`, the run, `)`.
    const std::size_t name = first_bare_token(node);
    if (name != kNone && is(name + 1U, parser::TokenKind::LParen) && is(name + 3U, parser::TokenKind::RParen)) {
      record(name + 2U, node, parser::WhitespaceSlot::kBeforeClose, 0, true);
    }
    return;
  }
  const parser::AstNode& last = node.as_call_arg(arity - 1U);
  if (last.kind() == parser::NodeKind::Literal && last.as_literal().is_blank()) {
    return;
  }
  // The call's own `)` is the first token after its last argument.
  std::size_t k = token_at(last.range().end);
  if (is(k, parser::TokenKind::Whitespace)) {
    ++k;
  }
  if (is(k, parser::TokenKind::RParen)) {
    close_run(k, node, 0);
  }
}

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
