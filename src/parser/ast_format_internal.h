//
// Internal declarations shared by the formula-text emitters (ast_format.cpp),
// the parenthesisation pass (ast_format_parens.cpp) and name spelling
// (ast_format_names.cpp). Not a public header; do not include from outside
// src/parser/.

#ifndef FORMULON_PARSER_AST_FORMAT_INTERNAL_H_
#define FORMULON_PARSER_AST_FORMAT_INTERNAL_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/ast_format.h"

namespace formulon {
namespace parser {
namespace detail {

// The LET binding / LAMBDA parameter names in scope, each as its
// `cell_shaped_name_bare`.
using NameScope = std::vector<std::string_view>;

bool HasParamPrefix(std::string_view name) noexcept;
void AppendParamName(std::string_view raw, std::string& out);
void AppendBindingName(std::string_view raw, bool storage, std::string& out);
bool AppendCellShapedCallName(std::string_view raw, const NameScope& scope, std::string& out);
void AppendNameRef(const AstNode& node, const NameScope& scope, std::string& out);

// The parentheses printed around `node`; `ParenCounts` is decided once by
// `CollectParens` for the formatters and the XLSB encoder alike.
inline std::uint16_t ParenCount(const ParenCounts& parens, const AstNode& node) {
  const auto it = parens.find(&node);
  return it == parens.end() ? 0U : it->second;
}

// Emits `count` items separated by commas; `emit_at(i)` writes item `i`.
template <typename EmitAt>
void EmitCommaList(std::uint32_t count, std::string& out, EmitAt&& emit_at) {
  for (std::uint32_t i = 0; i < count; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    emit_at(i);
  }
}

void CollectParens(const AstNode& node, int min_bp, ParenCounts& out);
bool PrintsSpliced(const AstNode& range, std::uint32_t i) noexcept;

}  // namespace detail
}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_FORMAT_INTERNAL_H_
