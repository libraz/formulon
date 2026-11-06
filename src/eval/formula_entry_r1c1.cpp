//
// R1C1 entry fallback declared in `eval/formula_entry_r1c1.h`: the triggers
// that send A1-parsed entry to R1C1, and the lexical R1C1 -> A1 rewrite.
//
// @size-budget: 15 KB (with the cell-shaped name spelling in parser/ast_format_names.cpp and parser/formula_prefix.cpp)

#include "eval/formula_entry_r1c1.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "eval/a1_parse.h"
#include "eval/builtin_names.h"
#include "parser/ast.h"
#include "parser/formula_prefix.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/arena.h"
#include "utils/formula_text.h"
#include "utils/strings.h"
#include "utils/text_ops.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

using parser::AstNode;
using parser::NodeKind;
using refs_internal::R1C1Endpoint;
using refs_internal::R1C1EndpointKind;

constexpr refs_internal::R1C1Letters kInvariantLetters{};
constexpr std::string_view kParamPrefix = "_xlpm.";
constexpr std::string_view kNamePrefix = "_xlnm.";

bool is_ident_byte(char c) noexcept {
  return parser::formula_prefix_is_ident_byte(c);
}

// The tokenizer's identifier start: a letter, `_`, `\` or a UTF-8 lead byte.
bool is_ident_start(char c) noexcept {
  const auto u = static_cast<unsigned char>(c);
  return (u >= 'A' && u <= 'Z') || (u >= 'a' && u <= 'z') || u == '_' || u == '\\' || u >= 0xC0;
}

bool is_space(char c) noexcept {
  return c == ' ' || c == '\n' || c == '\r';
}

bool is_r1c1_letter(char c) noexcept {
  const char u = strings::ascii_to_upper(c);
  return u == kInvariantLetters.row || u == kInvariantLetters.col;
}

std::size_t skip_spaces(std::string_view text, std::size_t i) {
  while (i < text.size() && is_space(text[i])) {
    ++i;
  }
  return i;
}

std::size_t ident_end(std::string_view text, std::size_t i) {
  ++i;
  while (i < text.size() && is_ident_byte(text[i])) {
    ++i;
  }
  return i;
}

// End of the run quoted by `text[i]` (`"` or `'`); a doubled quote escapes one.
std::size_t quoted_end(std::string_view text, std::size_t i) {
  const char quote = text[i];
  for (std::size_t j = i + 1; j < text.size(); ++j) {
    if (text[j] == quote) {
      if (j + 1 < text.size() && text[j + 1] == quote) {
        ++j;
        continue;
      }
      return j + 1;
    }
  }
  return text.size();
}

// End of the bracket group opening at `text[i]`, nested groups included; `'`
// escapes the next byte, as in a structured reference.
std::size_t bracket_end(std::string_view text, std::size_t i) {
  std::size_t depth = 0;
  for (std::size_t j = i; j < text.size(); ++j) {
    if (text[j] == '\'') {
      ++j;
    } else if (text[j] == '[') {
      ++depth;
    } else if (text[j] == ']' && --depth == 0U) {
      return j + 1;
    }
  }
  return text.size();
}

// A number literal and any identifier bytes glued to it, with an exponent sign.
std::size_t number_end(std::string_view text, std::size_t i) {
  while (i < text.size() && is_ident_byte(text[i])) {
    const bool exponent = strings::ascii_to_upper(text[i]) == 'E';
    ++i;
    if (exponent && i + 1 < text.size() && (text[i] == '+' || text[i] == '-') && strings::is_ascii_digit(text[i + 1])) {
      ++i;
    }
  }
  return i;
}

std::size_t error_literal_length(std::string_view text, std::size_t i) {
  for (const ErrorCodeInfo& info : kErrorTable) {
    const std::string_view name = info.display_name;
    if (text.size() - i >= name.size() && strings::case_insensitive_eq(text.substr(i, name.size()), name)) {
      return name.size();
    }
  }
  return 1;
}

// An endpoint stands as a reference unless an identifier byte, `(` or `!` follows it.
bool endpoint_stands(std::string_view text, std::size_t end) {
  return end >= text.size() || (!is_ident_byte(text[end]) && text[end] != '(' && text[end] != '!');
}

bool reads_r1c1_reference(std::string_view name) {
  if (name.empty() || !is_r1c1_letter(name.front())) {
    return false;
  }
  const R1C1Endpoint ep = refs_internal::parse_r1c1_endpoint_axes(name, 0, kInvariantLetters);
  return ep.kind == R1C1EndpointKind::kReference && ep.end == name.size();
}

bool whole_axis(const R1C1Endpoint& ep) {
  return ep.row.present != ep.col.present;
}

// Appends the A1 spelling of `ep` entered at (`row`, `col`), `$` on its
// absolute axes; a `lone` whole row or column spans itself (`$1:$1`).
bool append_endpoint(const R1C1Endpoint& ep, std::uint32_t row, std::uint32_t col, bool lone, std::string& out) {
  std::uint32_t r = 0;
  std::uint32_t c = 0;
  if ((ep.row.present && !refs_internal::resolve_r1c1_axis(ep.row, true, row, Sheet::kMaxRows, &r)) ||
      (ep.col.present && !refs_internal::resolve_r1c1_axis(ep.col, true, col, Sheet::kMaxCols, &c))) {
    return false;
  }
  std::string part;
  if (ep.col.present) {
    if (!ep.col.relative) {
      part.push_back('$');
    }
    a1::append_column_letters(part, c);
  }
  if (ep.row.present) {
    if (!ep.row.relative) {
      part.push_back('$');
    }
    part.append(std::to_string(r + 1U));
  }
  out.append(part);
  if (lone && whole_axis(ep)) {
    out.push_back(':');
    out.append(part);
  }
  return true;
}

// Appends the reference `ep` read at `*i`, joined with a following `:` endpoint
// naming the same single axis (`R1:R2` -> `$1:$2`), and advances `*i`.
bool append_reference(std::string_view text, const R1C1Endpoint& ep, std::uint32_t row, std::uint32_t col,
                      std::size_t* i, std::string& out) {
  const std::size_t colon = skip_spaces(text, ep.end);
  if (whole_axis(ep) && colon < text.size() && text[colon] == ':') {
    const std::size_t next = skip_spaces(text, colon + 1);
    if (next < text.size() && is_r1c1_letter(text[next])) {
      const R1C1Endpoint second = refs_internal::parse_r1c1_endpoint_axes(text, next, kInvariantLetters);
      if (second.kind == R1C1EndpointKind::kReference && endpoint_stands(text, second.end) &&
          second.row.present == ep.row.present && second.col.present == ep.col.present) {
        if (!append_endpoint(ep, row, col, false, out)) {
          return false;
        }
        out.append(text.substr(ep.end, next - ep.end));
        *i = second.end;
        return append_endpoint(second, row, col, false, out);
      }
    }
  }
  *i = ep.end;
  return append_endpoint(ep, row, col, true, out);
}

// What the last non-space output was, for the `name:(A)` rule.
enum class Last : std::uint8_t { kOther, kName, kColonAfterName };

std::optional<std::string> rewrite_r1c1(std::string_view text, std::uint32_t row, std::uint32_t col) {
  std::string out;
  out.reserve(text.size() + 16U);
  Last last = Last::kOther;
  std::size_t i = 0;
  while (i < text.size()) {
    const char c = text[i];
    if (is_space(c)) {
      out.push_back(c);
      ++i;
      continue;
    }
    if (c == '$') {
      return std::nullopt;
    }
    if (c == ':') {
      last = last == Last::kName ? Last::kColonAfterName : Last::kOther;
      out.push_back(c);
      ++i;
      continue;
    }
    const Last before = last;
    last = Last::kOther;
    std::size_t end = i + 1;
    if (c == '"' || c == '\'') {
      end = quoted_end(text, i);
    } else if (c == '[') {
      end = bracket_end(text, i);
    } else if (c == '#') {
      end = i + error_literal_length(text, i);
    } else if (strings::is_ascii_digit(c)) {
      end = number_end(text, i);
      const std::size_t after = skip_spaces(text, end);
      if (after < text.size() && text[after] == ':') {
        return std::nullopt;
      }
    } else if (is_ident_start(c)) {
      if (is_r1c1_letter(c)) {
        const R1C1Endpoint ep = refs_internal::parse_r1c1_endpoint_axes(text, i, kInvariantLetters);
        if (ep.kind == R1C1EndpointKind::kMalformed) {
          return std::nullopt;
        }
        if (ep.kind == R1C1EndpointKind::kReference && endpoint_stands(text, ep.end)) {
          if (!append_reference(text, ep, row, col, &i, out)) {
            return std::nullopt;
          }
          continue;
        }
      }
      end = ident_end(text, i);
      const std::string_view run = text.substr(i, end - i);
      const char next = end < text.size() ? text[end] : '\0';
      if (next == '!') {
        out.append(run);
      } else if (next == '(') {
        if (resolve_builtin_function_name(parser::strip_storage_prefixes(run)) == nullptr &&
            parser::is_a1_cell_shaped(run)) {
          out.append(kParamPrefix);
        }
        out.append(run);
      } else if (parser::is_a1_cell_shaped(run)) {
        out.append(kParamPrefix);
        out.append(run);
        last = Last::kName;
      } else if (before == Last::kColonAfterName && parser::is_a1_column_shaped(run)) {
        out.push_back('(');
        out.append(run);
        out.push_back(')');
      } else {
        out.append(run);
        last = Last::kName;
      }
      i = end;
      continue;
    }
    out.append(text.substr(i, end - i));
    i = end;
  }
  return out;
}

using Scope = std::vector<std::string_view>;

// Bindings compare by their bare spelling: `_xlpm.B1` binds `B1` and back.
bool binds(const Scope& scope, std::string_view name) {
  const std::string_view bare = parser::cell_shaped_name_bare(name);
  for (const std::string_view bound : scope) {
    if (strings::case_insensitive_eq(parser::cell_shaped_name_bare(bound), bare)) {
      return true;
    }
  }
  return false;
}

// Calls `visit(ctx, node, scope)` on every node outside a LET / LAMBDA head,
// with the names bound where the node sits: a LET value sees the earlier names.
using ScopedVisit = void (*)(void* ctx, const AstNode& node, const Scope& bound);

void walk_scoped(const AstNode& node, Scope& scope, ScopedVisit visit, void* ctx) {
  const std::size_t mark = scope.size();
  if (node.kind() == NodeKind::LetBinding) {
    for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
      walk_scoped(node.as_let_binding_expr(i), scope, visit, ctx);
      scope.push_back(node.as_let_binding_name(i));
    }
    walk_scoped(node.as_let_body(), scope, visit, ctx);
    scope.resize(mark);
    return;
  }
  if (node.kind() == NodeKind::Lambda) {
    for (std::uint32_t i = 0; i < node.as_lambda_param_count(); ++i) {
      scope.push_back(node.as_lambda_param(i));
    }
    walk_scoped(node.as_lambda_body(), scope, visit, ctx);
    scope.resize(mark);
    return;
  }
  visit(ctx, node, scope);
  parser::any_child_node(node, [&](const AstNode& child) {
    walk_scoped(child, scope, visit, ctx);
    return false;
  });
}

bool has_param_prefix(std::string_view name) {
  return parser::cell_shaped_name_prefix(name) == kParamPrefix;
}

// Byte offset in `body` of the `_xlpm.` that spells `node`'s name, or npos.
std::size_t param_prefix_site(std::string_view body, const AstNode& node) {
  const std::size_t from = utf16_to_byte_offset(body, node.range().start);
  const std::string_view span = body.substr(from, utf16_to_byte_offset(body, node.range().end) - from);
  std::size_t at = 0;
  if (node.kind() == NodeKind::NameRef && !node.as_name_sheet().empty()) {
    const std::size_t bang = span.rfind('!');
    at = skip_spaces(span, bang == std::string_view::npos ? 0U : bang + 1U);
  }
  if (span.size() - at < kParamPrefix.size() ||
      !strings::case_insensitive_eq(span.substr(at, kParamPrefix.size()), kParamPrefix)) {
    return std::string_view::npos;
  }
  return from + at;
}

}  // namespace

bool has_entered_cell_shaped_param(std::string_view text) {
  std::size_t i = 0;
  while (i < text.size()) {
    const char c = text[i];
    if (c == '"' || c == '\'') {
      i = quoted_end(text, i);
    } else if (c == '[') {
      i = bracket_end(text, i);
    } else if (is_ident_start(c)) {
      const std::size_t end = ident_end(text, i);
      if (has_param_prefix(text.substr(i, end - i))) {
        return true;
      }
      i = end;
    } else {
      ++i;
    }
  }
  return false;
}

bool needs_r1c1_entry(const AstNode& root, std::string_view text, const Workbook& book) {
  bool found = false;
  Scope scope;
  walk_scoped(
      root, scope,
      [](void* ctx, const AstNode& node, const Scope& bound) {
        if (node.kind() == NodeKind::NameRef && (!node.as_name_sheet().empty() || !binds(bound, node.as_name())) &&
            reads_r1c1_reference(node.as_name())) {
          *static_cast<bool*>(ctx) = true;
        }
      },
      &found);
  if (found) {
    return true;
  }
  std::size_t i = 0;
  while (i < text.size()) {
    const char c = text[i];
    if (c == '"' || c == '\'') {
      i = quoted_end(text, i);
      continue;
    }
    if (c == '[') {
      i = bracket_end(text, i);
      continue;
    }
    if (!is_ident_start(c)) {
      ++i;
      continue;
    }
    if (is_r1c1_letter(c)) {
      const R1C1Endpoint ep = refs_internal::parse_r1c1_endpoint_axes(text, i, kInvariantLetters);
      if (ep.bracketed && (ep.kind == R1C1EndpointKind::kMalformed || endpoint_stands(text, ep.end)) &&
          !book.is_table_name(text.substr(i, text.find(kInvariantLetters.open, i) - i))) {
        return true;
      }
    }
    i = ident_end(text, i);
  }
  return false;
}

std::optional<std::string> translate_r1c1_entry(std::string_view text, std::uint32_t row, std::uint32_t col) {
  std::optional<std::string> a1 = rewrite_r1c1(text, row, col);
  if (!a1) {
    return std::nullopt;
  }
  // A cell-shaped name no LET / LAMBDA in scope binds is an undefined name.
  const std::string_view body = strip_formula_prefix(*a1);
  const std::size_t body_at = a1->size() - body.size();
  Arena arena;
  const AstNode* root = parse_formula_entry(body, arena);
  if (root == nullptr) {
    return std::nullopt;
  }
  struct Sites {
    std::string_view body;
    std::vector<std::size_t> at;
  } sites{body, {}};
  Scope scope;
  const ScopedVisit visit = [](void* ctx, const AstNode& node, const Scope& bound) {
    Sites& out = *static_cast<Sites*>(ctx);
    std::string_view name;
    bool qualified = false;
    if (node.kind() == NodeKind::NameRef) {
      name = node.as_name();
      qualified = !node.as_name_sheet().empty();
    } else if (node.kind() == NodeKind::Call) {
      name = node.as_call_name();
    }
    // A sheet-qualified name is never a binding.
    if (!has_param_prefix(name) || (!qualified && binds(bound, name))) {
      return;
    }
    if (const std::size_t site = param_prefix_site(out.body, node); site != std::string_view::npos) {
      out.at.push_back(site);
    }
  };
  walk_scoped(*root, scope, visit, &sites);
  for (const std::size_t site : sites.at) {
    a1->replace(body_at + site, kNamePrefix.size(), kNamePrefix);
  }
  return a1;
}

}  // namespace eval
}  // namespace formulon
