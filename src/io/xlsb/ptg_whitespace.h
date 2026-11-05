//
// Collects the whitespace runs a formula's source text carries between its
// tokens, keyed by AST node, for the encoder to store as `PtgAttrSpace`.

#ifndef FORMULON_IO_XLSB_PTG_WHITESPACE_H_
#define FORMULON_IO_XLSB_PTG_WHITESPACE_H_

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

/// Finds, in the text a formula was parsed from, the whitespace runs Excel
/// stores as `PtgAttrSpace` (see `encode_ptgs`), keyed by node.
class SourceWhitespace {
 public:
  SourceWhitespace(std::string_view source, parser::FormulaWhitespace& out) : tokenizer_(source), out_(out) {}

  void collect(const parser::AstNode& root) {
    tokens_ = &tokenizer_.tokens();
    visit(root, /*lead_allowed=*/true);
  }

 private:
  void visit(const parser::AstNode& node, bool lead_allowed);
  std::size_t token_at(std::uint32_t offset) const;
  bool is(std::size_t i, parser::TokenKind kind) const { return i < tokens_->size() && (*tokens_)[i].kind == kind; }
  std::size_t last_token(const parser::AstNode& node) const;
  std::size_t first_bare_token(const parser::AstNode& node) const;
  void record(std::size_t i, const parser::AstNode& node, parser::WhitespaceSlot slot, std::uint16_t index,
              bool line_feeds_allowed, std::size_t skip = 0);
  void lead(const parser::AstNode& node, bool allowed);
  void before_operator(const parser::AstNode& node, const parser::AstNode& lhs, parser::TokenKind op,
                       std::uint16_t index = 0);
  void intersection(const parser::AstNode& node);
  void close_run(std::size_t rparen, const parser::AstNode& node, std::uint8_t index);
  void paren_opens(const parser::AstNode& node, bool lead_allowed);
  void paren_closes(const parser::AstNode& node);
  void call_close(const parser::AstNode& node);

  static constexpr std::size_t kNone = static_cast<std::size_t>(-1);
  parser::Tokenizer tokenizer_;
  const std::vector<parser::Token>* tokens_ = nullptr;
  parser::FormulaWhitespace& out_;
};

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_WHITESPACE_H_
