//
// Respelling of the `SINGLE` / `ANCHORARRAY` storage calls as the `@x` / `x#`
// operators; the contract is declared in `parser/formula_prefix.h`.

#include "parser/formula_prefix.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/index_sort.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {

namespace {

bool IsOperatorCall(const AstNode& node, std::string_view name) {
  return node.kind() == NodeKind::Call && node.as_call_arity() == 1U &&
         strings::case_insensitive_eq(node.as_call_name(), name);
}

bool IsStorageOperatorCall(const AstNode& node) {
  return IsOperatorCall(node, "SINGLE") || IsOperatorCall(node, "ANCHORARRAY");
}

bool HasStorageOperatorCall(const AstNode& node) {
  if (IsStorageOperatorCall(node)) {
    return true;
  }
  for (const AstNode* child : child_nodes(node)) {
    if (HasStorageOperatorCall(*child)) {
      return true;
    }
  }
  return false;
}

/// Rewrites `node`'s source text, copying every part without an operator call
/// verbatim so spacing and spelling survive.
class OperatorRespeller {
 public:
  explicit OperatorRespeller(std::string_view src) : src_(src) {}

  std::string text(const AstNode& node, bool tight_parent) const {
    const TextRange r = node.range();
    if (!HasStorageOperatorCall(node)) {
      return std::string(src_.substr(r.start, r.end - r.start));
    }
    if (IsStorageOperatorCall(node)) {
      const AstNode& arg = node.as_call_arg(0);
      if (IsOperatorCall(node, "SINGLE")) {
        // `@` binds looser than `:`, ` ` and `%`, tighter than every infix operator.
        const bool wrap_arg = arg.kind() == NodeKind::BinaryOp ||
                              (arg.kind() == NodeKind::UnaryOp && arg.as_unary_op() == UnaryOp::Percent);
        const std::string at = "@" + parenthesised(arg, wrap_arg);
        return tight_parent ? "(" + at + ")" : at;
      }
      const bool bare = arg.kind() == NodeKind::Ref || arg.kind() == NodeKind::Ref3D ||
                        arg.kind() == NodeKind::NameRef || arg.kind() == NodeKind::Call ||
                        arg.kind() == NodeKind::ExternalRef;
      return parenthesised(arg, !bare) + "#";
    }
    const std::vector<const AstNode*> children = child_nodes(node);
    const auto by_start = [&](std::uint32_t lhs, std::uint32_t rhs) {
      return children[lhs]->range().start < children[rhs]->range().start;
    };
    std::vector<std::uint32_t> order;
    sorted_index_order(order, static_cast<std::uint32_t>(children.size()), make_index_less(by_start));
    const bool tight = node.kind() == NodeKind::RangeOp || node.kind() == NodeKind::IntersectOp ||
                       node.kind() == NodeKind::SpillRef ||
                       (node.kind() == NodeKind::UnaryOp && node.as_unary_op() == UnaryOp::Percent);
    std::string out;
    std::uint32_t at = r.start;
    for (const std::uint32_t index : order) {
      const AstNode* child = children[index];
      if (!HasStorageOperatorCall(*child)) {
        continue;
      }
      out.append(src_.substr(at, child->range().start - at));
      out.append(text(*child, tight));
      at = child->range().end;
    }
    out.append(src_.substr(at, r.end - at));
    return out;
  }

 private:
  std::string parenthesised(const AstNode& node, bool wrap) const {
    std::string inner = text(node, false);
    return wrap && (inner.empty() || inner.front() != '(') ? "(" + inner + ")" : inner;
  }

  std::string_view src_;
};

}  // namespace

std::string spell_storage_operators(std::string_view formula) {
  const std::size_t body_at = !formula.empty() && formula.front() == '=' ? 1U : 0U;
  const std::string_view body = formula.substr(body_at);
  if (!strings::case_insensitive_contains(body, "SINGLE(") &&
      !strings::case_insensitive_contains(body, "ANCHORARRAY(")) {
    return std::string(formula);
  }
  Arena arena;
  const AstNode* root = parse_strict(body, arena);
  if (root == nullptr || !HasStorageOperatorCall(*root)) {
    return std::string(formula);
  }
  const TextRange r = root->range();
  const std::string respelled = OperatorRespeller(body).text(*root, false);
  return std::string(formula.substr(0, body_at + r.start)) + respelled + std::string(body.substr(r.end));
}

}  // namespace parser
}  // namespace formulon
