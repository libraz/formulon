//
// Internal declarations shared by the formula-text emitters (ast_format.cpp),
// the parenthesisation pass (ast_format_parens.cpp) and name spelling
// (ast_format_names.cpp). Not a public header; do not include from outside
// src/parser/.

#ifndef FORMULON_PARSER_AST_FORMAT_INTERNAL_H_
#define FORMULON_PARSER_AST_FORMAT_INTERNAL_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/formula_prefix.h"

namespace formulon {
namespace parser {
namespace detail {

// The LET binding / LAMBDA parameter names in scope, each as its
// `cell_shaped_name_bare`.
using NameScope = std::vector<std::string_view>;

bool HasParamPrefix(std::string_view name) noexcept;
void AppendParamName(std::string_view raw, std::string& out);
void AppendBindingName(std::string_view raw, bool storage, std::string& out);
// Appends a LAMBDA's parameter list (optional ones bracketed, a trailing comma
// when non-empty) and pushes each parameter onto `scope`.
void AppendLambdaParams(const AstNode& node, bool storage, NameScope& scope, std::string& out);
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

// Appends `[structured]`; nothing when `structured` is empty.
void AppendStructuredPayload(std::string_view structured, std::string& out);

// Appends the sheet span `begin:end`, as one quoted unit when `quote` is set.
void AppendSheetSpan(std::string_view begin, std::string_view end, bool quote, std::string& out);

// Emits a LAMBDA: `append_name()` spells the function name, `emit(node)` writes
// the body. The parameters are in `scope` for the body only.
template <typename AppendName, typename Emit>
void EmitLambda(const AstNode& node, bool storage, NameScope& scope, std::string& out, AppendName&& append_name,
                Emit&& emit) {
  append_name();
  out.push_back('(');
  const std::size_t base = scope.size();
  AppendLambdaParams(node, storage, scope, out);
  emit(node.as_lambda_body());
  out.push_back(')');
  scope.resize(base);
}

// Emits a LET: `append_name()` spells the function name, `emit(node)` writes a
// binding value or the body. A binding value sees the earlier bindings only.
template <typename AppendName, typename Emit>
void EmitLet(const AstNode& node, bool storage, NameScope& scope, std::string& out, AppendName&& append_name,
             Emit&& emit) {
  append_name();
  out.push_back('(');
  const std::uint32_t n = node.as_let_binding_count();
  const std::size_t base = scope.size();
  for (std::uint32_t i = 0; i < n; ++i) {
    const std::string_view name = node.as_let_binding_name(i);
    AppendBindingName(name, storage, out);
    out.push_back(',');
    emit(node.as_let_binding_expr(i));
    out.push_back(',');
    scope.push_back(cell_shaped_name_bare(name));
  }
  emit(node.as_let_body());
  out.push_back(')');
  scope.resize(base);
}

void CollectParens(const AstNode& node, int min_bp, ParenCounts& out);
// `CollectParens` for a whole formula. A root union with no written
// parentheses is a defined-name body's bare union and prints bare.
void CollectRootParens(const AstNode& root, ParenCounts& out);
bool PrintsSpliced(const AstNode& range, std::uint32_t i) noexcept;

// The text operand `i` of a `:` chain prints quoted, so that `Data:'Mid'!A1`
// never reads back as a span: the qualifier of an operand after an
// unqualified name, or `quoted_name(name)` for an unqualified name before a
// qualified operand. Empty when neither applies.
std::string RangeOperandQuotedUnit(const AstNode& range, std::uint32_t i,
                                   std::string_view (*quoted_name)(std::string_view) noexcept);

// Quotes the first bare occurrence of `unit` in `text`; one already quoted
// or behind a name prefix (`_xlnm.`) is left alone.
void QuoteBareUnit(std::string_view unit, std::string& text);

// True when operand `i` of a `:` chain is written in A1 with its qualifier quoted: another book's sheet-qualified
// reference (`A1:'[1]Nope'!B2`), a local sheet-qualified one after another book's reference
// (`[1]Nope!A1:'Data'!B2`), and in `storage` any local sheet-qualified one (`Sheet1!A1:'Data'!B2`), which the
// formula bar shows bare.
bool QuotesQualifiedRightOperand(const AstNode& range, std::uint32_t i, bool storage) noexcept;

// Quotes the bare qualifier `text` opens with, up to its `!`; text that opens quoted is left alone.
void QuoteLeadingQualifier(std::string& text);

}  // namespace detail
}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_FORMAT_INTERNAL_H_
