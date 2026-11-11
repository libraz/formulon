//
// R1C1 entry fallback declared in `eval/formula_entry_r1c1.h`: the triggers
// that send A1-parsed entry to R1C1, and the lexical R1C1 -> A1 rewrite.
//
// @size-budget: 19 KB (with the cell-shaped name spelling in parser/ast_format_names.cpp and parser/formula_prefix.cpp)

#include "eval/formula_entry_r1c1.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/a1_parse.h"
#include "eval/builtin_names.h"
#include "eval/formula_localize.h"
#include "parser/ast.h"
#include "parser/formula_prefix.h"
#include "parser/parse_error.h"
#include "parser/parse_status.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/number_text.h"
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

bool is_r1c1_letter(char c, const refs_internal::R1C1Letters& letters) noexcept {
  const char u = strings::ascii_to_upper(c);
  return u == letters.row || u == letters.col;
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
  if (name.empty() || !is_r1c1_letter(name.front(), kInvariantLetters)) {
    return false;
  }
  const R1C1Endpoint ep = refs_internal::parse_r1c1_endpoint_axes(name, 0, kInvariantLetters);
  return ep.kind == R1C1EndpointKind::kReference && ep.end == name.size();
}

bool whole_axis(const R1C1Endpoint& ep) {
  return ep.row.present != ep.col.present;
}

// The A1 spelling of `ep` measured from the base, `$` on its absolute axes;
// false when a relative axis has no base.
bool a1_endpoint(const R1C1Endpoint& ep, const R1C1RewriteOptions& o, std::string& part) {
  std::uint32_t r = 0;
  std::uint32_t c = 0;
  if ((ep.row.present && !refs_internal::resolve_r1c1_axis(ep.row, o.has_base, o.base_row, Sheet::kMaxRows, &r)) ||
      (ep.col.present && !refs_internal::resolve_r1c1_axis(ep.col, o.has_base, o.base_col, Sheet::kMaxCols, &c))) {
    return false;
  }
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
    char digits[24];
    part.append(digits, static_cast<std::size_t>(format_signed(digits, sizeof digits, r + 1LL)));
  }
  return true;
}

// What the last non-space output was, for the `name:(A)` rule.
enum class Last : std::uint8_t { kOther, kName, kColonAfterName };

class R1C1Rewriter {
 public:
  R1C1Rewriter(std::string_view text, const R1C1RewriteOptions& options) : text_(text), o_(options) {
    out_.text.reserve(text.size() + 16U);
  }

  Expected<MappedFormulaText, parser::ParseFailure> run() {
    Last last = Last::kOther;
    int braces = 0;
    std::size_t i = 0;
    while (i < text_.size()) {
      const char c = text_[i];
      if (is_space(c)) {
        copy(i, i + 1U);
        ++i;
        continue;
      }
      if (c == '$' && o_.entry) {
        return fail(parser::ParseErrorCode::UnexpectedToken, i, i + 1U);
      }
      if (c == ':') {
        last = last == Last::kName ? Last::kColonAfterName : Last::kOther;
        copy(i, i + 1U);
        ++i;
        continue;
      }
      const Last before = last;
      last = Last::kOther;
      if (c == '{') {
        ++braces;
      } else if (c == '}' && braces > 0) {
        --braces;
      }
      if (c == '\'' && !o_.entry && append_qualifier(&i)) {
        continue;
      }
      std::size_t end = i + 1U;
      if (c == '"' || c == '\'') {
        end = quoted_end(text_, i);
      } else if (c == '[') {
        end = bracket_end(text_, i);
      } else if (c == '#') {
        end = i + error_literal_length(text_, i);
      } else if (strings::is_ascii_digit(c)) {
        end = number_end(text_, i);
        const std::size_t after = skip_spaces(text_, end);
        const bool row_separator = braces > 0 && o_.array_row_separator == ':';
        if (o_.entry && !row_separator && after < text_.size() && text_[after] == ':') {
          return fail(parser::ParseErrorCode::UnexpectedToken, i, end);
        }
      } else if (is_ident_start(c)) {
        if (is_r1c1_letter(c, o_.letters)) {
          const R1C1Endpoint ep = refs_internal::parse_r1c1_endpoint_axes(text_, i, o_.letters);
          if (ep.kind == R1C1EndpointKind::kMalformed) {
            return fail(parser::ParseErrorCode::InvalidReference, i, malformed_end(ep.end));
          }
          if (ep.kind == R1C1EndpointKind::kReference && stands(ep.end)) {
            if (Expected<void, parser::ParseFailure> written = append_reference(ep, &i); !written) {
              return written.error();
            }
            continue;
          }
        }
        if (!o_.entry && append_qualifier(&i)) {
          continue;
        }
        end = ident_end(text_, i);
        append_name(i, end, before, last);
        i = end;
        continue;
      }
      copy(i, end);
      i = end;
    }
    return std::move(out_);
  }

 private:
  void put(std::size_t begin, std::size_t end, std::string_view piece) {
    const std::size_t out_begin = out_.text.size();
    out_.text.append(piece);
    const bool verbatim = piece == text_.substr(begin, end - begin);
    std::vector<FormulaSourceSpan>& spans = out_.spans;
    if (verbatim && !spans.empty() && spans.back().verbatim && spans.back().in_end == begin) {
      spans.back().in_end = end;
      spans.back().out_end = out_.text.size();
      return;
    }
    spans.push_back(FormulaSourceSpan{out_begin, out_.text.size(), begin, end, verbatim});
  }

  void copy(std::size_t begin, std::size_t end) { put(begin, end, text_.substr(begin, end - begin)); }

  parser::ParseFailure fail(parser::ParseErrorCode code, std::size_t begin, std::size_t end) const {
    return parser::ParseFailure{code, formula_text_range(text_, begin, end), parser::default_message(code)};
  }

  // An endpoint stands before a trim operator too, outside entry.
  bool stands(std::size_t end) const {
    return endpoint_stands(text_, end) ||
           (!o_.entry && end + 1U < text_.size() && text_[end] == '.' && text_[end + 1U] == ':');
  }

  // End of a malformed endpoint's text: through its closing bracket when one follows.
  std::size_t malformed_end(std::size_t from) const {
    if (from > 0U && text_[from - 1U] == o_.letters.close) {
      return from;
    }
    std::size_t j = from;
    while (j < text_.size() && is_ident_byte(text_[j]) && text_[j] != o_.letters.close) {
      ++j;
    }
    return j < text_.size() && text_[j] == o_.letters.close ? j + 1U : j;
  }

  // Offset past the range operator that starts at `from`, or npos: `:` after
  // optional spaces, or outside entry a trim operator.
  std::size_t operator_end(std::size_t from) const {
    const bool trim = !o_.entry;
    std::size_t j = trim && from < text_.size() && text_[from] == '.' ? from + 1U : skip_spaces(text_, from);
    if (j >= text_.size() || text_[j] != ':') {
      return std::string_view::npos;
    }
    ++j;
    return trim && j < text_.size() && text_[j] == '.' ? j + 1U : j;
  }

  Expected<void, parser::ParseFailure> put_endpoint(const R1C1Endpoint& ep, std::size_t begin, bool lone) {
    std::string part;
    if (!a1_endpoint(ep, o_, part)) {
      return fail(parser::ParseErrorCode::InvalidReference, begin, ep.end);
    }
    // A lone whole row or column spans itself (`$1:$1`).
    if (lone && whole_axis(ep)) {
      const std::string axis = part;
      part.push_back(':');
      part.append(axis);
    }
    put(begin, ep.end, part);
    return {};
  }

  // Writes the reference `ep` read at `*i`, joined with a following endpoint
  // naming the same single axis (`R1:R2` -> `$1:$2`), and advances `*i`.
  Expected<void, parser::ParseFailure> append_reference(const R1C1Endpoint& ep, std::size_t* i) {
    const std::size_t begin = *i;
    const std::size_t op_end = whole_axis(ep) ? operator_end(ep.end) : std::string_view::npos;
    if (op_end != std::string_view::npos) {
      const std::size_t next = skip_spaces(text_, op_end);
      if (next < text_.size() && is_r1c1_letter(text_[next], o_.letters)) {
        const R1C1Endpoint second = refs_internal::parse_r1c1_endpoint_axes(text_, next, o_.letters);
        if (second.kind == R1C1EndpointKind::kReference && stands(second.end) && second.row.present == ep.row.present &&
            second.col.present == ep.col.present) {
          if (Expected<void, parser::ParseFailure> first = put_endpoint(ep, begin, false); !first) {
            return first;
          }
          copy(ep.end, next);
          *i = second.end;
          return put_endpoint(second, next, false);
        }
      }
    }
    *i = ep.end;
    return put_endpoint(ep, begin, true);
  }

  // End of the quoted or bare name at `j`, unescaped into `*name`; npos when
  // none starts there.
  std::size_t name_end(std::size_t j, std::string* name, bool* quoted) const {
    if (j >= text_.size()) {
      return std::string_view::npos;
    }
    if (text_[j] != '\'') {
      if (!is_ident_start(text_[j])) {
        return std::string_view::npos;
      }
      const std::size_t end = ident_end(text_, j);
      name->assign(text_.substr(j, end - j));
      return end;
    }
    const std::size_t end = quoted_end(text_, j);
    if (end - j < 2U || text_[end - 1U] != '\'') {
      return std::string_view::npos;
    }
    *quoted = true;
    for (std::size_t k = j + 1U; k + 1U < end; ++k) {
      name->push_back(text_[k]);
      if (text_[k] == '\'') {
        ++k;
      }
    }
    return end;
  }

  // A local sheet qualifier at `*i` (a quoted or bare name, or two joined by
  // `:`, then `!`), written with the stored quoting rule; a partly quoted 3-D
  // span stays as written. Advances `*i` to the `!`; false when none starts here.
  bool append_qualifier(std::size_t* i) {
    if (*i > 0U && text_[*i - 1U] == ']') {
      return false;  // the sheet of a cross-workbook qualifier
    }
    std::string first;
    std::string last;
    bool first_quoted = false;
    bool last_quoted = false;
    const std::size_t end = name_end(*i, &first, &first_quoted);
    if (end == std::string_view::npos || end >= text_.size()) {
      return false;
    }
    std::size_t bang = end;
    bool span = false;
    if (text_[end] == ':') {
      bang = name_end(end + 1U, &last, &last_quoted);
      if (bang == std::string_view::npos || bang >= text_.size() || text_[bang] != '!') {
        return false;
      }
      span = true;
    } else if (text_[end] != '!') {
      return false;
    }
    if (span && (first_quoted || last_quoted)) {
      copy(*i, bang);
      *i = bang;
      return true;
    }
    if (first_quoted && first.find('[') != std::string::npos) {
      return false;
    }
    if (const std::size_t colon = first.find(':'); !span && colon != std::string::npos) {
      last = first.substr(colon + 1U);
      first.resize(colon);
      span = true;
    }
    const bool quote =
        parser::local_sheet_needs_quoting_a1(first) || (span && parser::local_sheet_needs_quoting_a1(last));
    std::string piece;
    append_sheet_qualifier(piece, first, span ? std::string_view(last) : std::string_view(), quote);
    put(*i, bang, piece);
    *i = bang;
    return true;
  }

  // Writes the identifier `text_[begin, end)`, spelling cell- and
  // column-shaped names so the A1 reading keeps them names.
  void append_name(std::size_t begin, std::size_t end, Last before, Last& last) {
    const std::string_view run = text_.substr(begin, end - begin);
    const char next = end < text_.size() ? text_[end] : '\0';
    std::string_view prefix;
    bool parenthesise = false;
    if (next == '(') {
      if (resolve_builtin_function_name(parser::strip_storage_prefixes(run)) == nullptr &&
          parser::is_a1_cell_shaped(run)) {
        prefix = kParamPrefix;
      }
    } else if (next != '!') {
      if (parser::is_a1_cell_shaped(run)) {
        prefix = kParamPrefix;
        last = Last::kName;
      } else if (before == Last::kColonAfterName && parser::is_a1_column_shaped(run)) {
        parenthesise = true;
      } else {
        last = Last::kName;
      }
    }
    if (prefix.empty() && !parenthesise) {
      copy(begin, end);
      return;
    }
    std::string piece(prefix);
    if (parenthesise) {
      piece.push_back('(');
    }
    piece.append(run);
    if (parenthesise) {
      piece.push_back(')');
    }
    put(begin, end, piece);
  }

  std::string_view text_;
  const R1C1RewriteOptions& o_;
  MappedFormulaText out_;
};

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
    if (is_r1c1_letter(c, kInvariantLetters)) {
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

Expected<MappedFormulaText, parser::ParseFailure> rewrite_r1c1_text(std::string_view text,
                                                                    const R1C1RewriteOptions& options) {
  return R1C1Rewriter(text, options).run();
}

void respell_unbound_cell_names(const AstNode& root, std::string& text) {
  const std::string_view body = strip_formula_prefix(text);
  const std::size_t body_at = text.size() - body.size();
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
  walk_scoped(root, scope, visit, &sites);
  for (const std::size_t site : sites.at) {
    text.replace(body_at + site, kNamePrefix.size(), kNamePrefix);
  }
}

void for_each_unbound_name(const AstNode& root, void (*visit)(void* ctx, const AstNode& node), void* ctx) {
  struct Target {
    void (*visit)(void*, const AstNode&);
    void* ctx;
  } target{visit, ctx};
  Scope scope;
  walk_scoped(
      root, scope,
      [](void* raw, const AstNode& node, const Scope& bound) {
        if (node.kind() == NodeKind::NameRef && node.as_name_sheet().empty() && !binds(bound, node.as_name())) {
          const Target& t = *static_cast<const Target*>(raw);
          t.visit(t.ctx, node);
        }
      },
      &target);
}

std::optional<std::string> translate_r1c1_entry(std::string_view text, std::uint32_t row, std::uint32_t col) {
  R1C1RewriteOptions options;
  options.entry = true;
  options.has_base = true;
  options.base_row = row;
  options.base_col = col;
  Expected<MappedFormulaText, parser::ParseFailure> rewritten = rewrite_r1c1_text(text, options);
  if (!rewritten) {
    return std::nullopt;
  }
  // A cell-shaped name no LET / LAMBDA in scope binds is an undefined name.
  std::optional<std::string> a1 = std::move(rewritten.value().text);
  Arena arena;
  const AstNode* root = parse_formula_entry(strip_formula_prefix(*a1), arena);
  if (root == nullptr) {
    return std::nullopt;
  }
  respell_unbound_cell_names(*root, *a1);
  return a1;
}

}  // namespace eval
}  // namespace formulon
