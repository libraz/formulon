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

bool InScope(const NameScope& scope, std::string_view bare);
bool HasParamPrefix(std::string_view name) noexcept;
void AppendParamName(std::string_view raw, std::string& out);
void AppendBindingName(std::string_view raw, bool storage, std::string& out);
bool AppendCellShapedCallName(std::string_view raw, const NameScope& scope, std::string& out);
void AppendNameRef(const AstNode& node, const NameScope& scope, std::string& out);

void CollectParens(const AstNode& node, int min_bp, ParenCounts& out);
bool PrintsSpliced(const AstNode& range, std::uint32_t i) noexcept;

}  // namespace detail
}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_FORMAT_INTERNAL_H_
