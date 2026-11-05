//
// Display and bind passes over the structured references of formula text.
// See `structured_ref_spelling.h` for the contract.

#include "parser/structured_ref_spelling.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "parser/ast.h"
#include "parser/parser.h"
#include "parser/structured_ref_payload.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "utils/text_ops.h"

namespace formulon {
namespace parser {
namespace {

// Replaces body bytes [start, end) with `text`.
struct Edit {
  std::size_t start;
  std::size_t end;
  std::string text;
};

bool IsFormulaSpace(char c) noexcept {
  return c == ' ' || c == '\n' || c == '\r' || c == '\t';
}

// Byte span of `node` in `body`, without the parentheses written around it.
Edit NodeSpan(const AstNode& node, std::string_view body) {
  std::size_t start = utf16_to_byte_offset(body, node.range().start);
  std::size_t end = utf16_to_byte_offset(body, node.range().end);
  if (node.paren_depth() > 0U) {
    while (start < end && (body[start] == '(' || IsFormulaSpace(body[start]))) {
      ++start;
    }
    while (end > start && (body[end - 1U] == ')' || IsFormulaSpace(body[end - 1U]))) {
      --end;
    }
  }
  return Edit{start, end, {}};
}

std::size_t BodyOffset(std::string_view text) noexcept {
  return !text.empty() && text.front() == '=' ? 1U : 0U;
}

// Leaves are visited in source order, so the edits are ascending; applying
// them right to left keeps the earlier offsets valid.
std::string Splice(std::string_view text, std::size_t body_at, const std::vector<Edit>& edits) {
  std::string out(text);
  for (auto it = edits.rbegin(); it != edits.rend(); ++it) {
    out.replace(body_at + it->start, it->end - it->start, it->text);
  }
  return out;
}

void CollectDisplay(const AstNode& node, std::string_view body, std::string_view own_table, std::vector<Edit>* edits) {
  if (node.kind() == NodeKind::StructuredRef) {
    const std::optional<StructuredRefPayload> payload =
        parse_structured_ref_payload_text(node.as_structured_ref_column());
    if (!payload) {
      return;
    }
    const std::string_view table = node.as_structured_ref_table();
    const bool inside = !own_table.empty() && strings::case_insensitive_eq(table, own_table);
    Edit edit = NodeSpan(node, body);
    edit.text = display_structured_ref(table, *payload, inside);
    if (body.substr(edit.start, edit.end - edit.start) != edit.text) {
      edits->push_back(std::move(edit));
    }
    return;
  }
  any_child_node(node, [&](const AstNode& child) {
    CollectDisplay(child, body, own_table, edits);
    return false;
  });
}

struct BindWalk {
  std::string_view body;
  BareTableNamePredicate is_table_name;
  // LET names and LAMBDA parameters in scope.
  std::vector<std::string_view> bound;
  std::vector<Edit> edits;

  bool is_bound(std::string_view name) const noexcept {
    for (const std::string_view b : bound) {
      if (strings::case_insensitive_eq(b, name)) {
        return true;
      }
    }
    return false;
  }

  void bind_structured_ref(const AstNode& node) {
    Edit edit = NodeSpan(node, body);
    const std::string_view written = body.substr(edit.start, edit.end - edit.start);
    const std::string_view raw = node.as_structured_ref_column();
    const bool bare = !written.empty() && written.front() == '[';
    const std::string_view trimmed = strings::trim_blanks(raw);
    if (!bare && (trimmed.empty() || trimmed.front() != '@')) {
      return;
    }
    const std::optional<StructuredRefPayload> payload = parse_structured_ref_payload_text(raw);
    edit.text.append(node.as_structured_ref_table()).push_back('[');
    edit.text.append(payload ? stored_payload(*payload) : std::string(raw)).push_back(']');
    if (edit.text != written) {
      edits.push_back(std::move(edit));
    }
  }

  void visit(const AstNode& node) {
    switch (node.kind()) {
      case NodeKind::StructuredRef:
        bind_structured_ref(node);
        return;
      case NodeKind::NameRef:
        if (node.as_name_sheet().empty() && is_table_name.is_bare_table != nullptr && !is_bound(node.as_name()) &&
            is_table_name.is_bare_table(is_table_name.ctx, node.as_name())) {
          const std::size_t end = NodeSpan(node, body).end;
          edits.push_back(Edit{end, end, "[]"});
        }
        return;
      case NodeKind::LetBinding: {
        const std::size_t mark = bound.size();
        for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
          visit(node.as_let_binding_expr(i));
          bound.push_back(node.as_let_binding_name(i));
        }
        visit(node.as_let_body());
        bound.resize(mark);
        return;
      }
      case NodeKind::Lambda: {
        const std::size_t mark = bound.size();
        for (std::uint32_t i = 0; i < node.as_lambda_param_count(); ++i) {
          bound.push_back(node.as_lambda_param(i));
        }
        visit(node.as_lambda_body());
        bound.resize(mark);
        return;
      }
      default:
        any_child_node(node, [this](const AstNode& child) {
          visit(child);
          return false;
        });
        return;
    }
  }
};

}  // namespace

std::string display_structured_refs(std::string_view text, std::string_view own_table) {
  if (text.find('[') == std::string_view::npos) {
    return std::string(text);
  }
  const std::size_t body_at = BodyOffset(text);
  const std::string_view body = text.substr(body_at);
  Arena arena;
  const AstNode* root = parse_strict(body, arena);
  if (root == nullptr) {
    return std::string(text);
  }
  std::vector<Edit> edits;
  CollectDisplay(*root, body, own_table, &edits);
  return Splice(text, body_at, edits);
}

std::string bind_structured_refs(std::string_view text, std::string_view own_table,
                                 BareTableNamePredicate is_table_name) {
  if (text.find('[') == std::string_view::npos && is_table_name.is_bare_table == nullptr) {
    return std::string(text);
  }
  const std::size_t body_at = BodyOffset(text);
  const std::string_view body = text.substr(body_at);
  Arena arena;
  ParserOptions opts;
  opts.bare_structured_ref_table = own_table;
  const AstNode* root = parse_strict(body, arena, opts);
  if (root == nullptr) {
    return std::string(text);
  }
  BindWalk walk{body, is_table_name, {}, {}};
  walk.visit(*root);
  return Splice(text, body_at, walk.edits);
}

}  // namespace parser
}  // namespace formulon
