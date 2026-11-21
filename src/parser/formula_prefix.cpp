//
// Respelling of the `SINGLE` / `ANCHORARRAY` storage calls as the `@x` / `x#`
// operators, and the cell-shaped name prefixes; the contracts are declared in
// `parser/formula_prefix.h`.

#include "parser/formula_prefix.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
#include "utils/strings.h"
#include "utils/text_ops.h"

namespace formulon {
namespace parser {

namespace {

bool IsOperatorCall(const AstNode& node, std::string_view name) {
  return node.kind() == NodeKind::Call && node.as_call_arity() == 1U &&
         strings::case_insensitive_eq(node.as_call_name(), name);
}

bool IsStorageOperatorCall(const AstNode& node) {
  return IsOperatorCall(node, kSingleFn) || IsOperatorCall(node, kAnchorArrayFn) ||
         trim_ref_call_mode(node) != TrimRefMode::None;
}

bool HasStorageOperatorCall(const AstNode& node) {
  if (IsStorageOperatorCall(node)) {
    return true;
  }
  if (node.kind() == NodeKind::ArrayLiteral) {
    return false;
  }
  return any_child_node(node, [](const AstNode& child) { return HasStorageOperatorCall(child); });
}

/// Rewrites `node`'s source text, copying every part without an operator call
/// verbatim so spacing and spelling survive.
class OperatorRespeller {
 public:
  explicit OperatorRespeller(std::string_view src) : src_(src) {}

  std::string text(const AstNode& node, bool tight_parent) const {
    const TextRange r = node.range();
    if (!HasStorageOperatorCall(node)) {
      return std::string(slice(r.start, r.end));
    }
    if (IsStorageOperatorCall(node)) {
      const AstNode& arg = node.as_call_arg(0);
      if (const TrimRefMode mode = trim_ref_call_mode(node); mode != TrimRefMode::None) {
        return trim_ref(node, arg, mode);
      }
      if (IsOperatorCall(node, kSingleFn)) {
        // `@` binds looser than `:`, ` ` and `%`, tighter than every infix operator.
        const bool wrap_arg = arg.kind() == NodeKind::BinaryOp ||
                              (arg.kind() == NodeKind::UnaryOp && arg.as_unary_op() == UnaryOp::Percent);
        const std::string at = "@" + parenthesised(arg, wrap_arg);
        return tight_parent ? "(" + at + ")" : at;
      }
      // An anchor `#` cannot follow keeps the call spelling, which parses.
      if (!is_spill_anchor(arg)) {
        return std::string(slice(r.start, arg.range().start)) + text(arg, false) +
               std::string(slice(arg.range().end, r.end));
      }
      return text(arg, false) + "#";
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
      out.append(slice(at, child->range().start));
      out.append(text(*child, tight));
      at = child->range().end;
    }
    out.append(slice(at, r.end));
    return out;
  }

  /// The source text between two token offsets, which count UTF-16 code
  /// units while `src_` is UTF-8.
  std::string_view slice(std::uint32_t start, std::uint32_t end) const {
    const std::size_t from = utf16_to_byte_offset(src_, start);
    return src_.substr(from, utf16_to_byte_offset(src_, end) - from);
  }

 private:
  // `_TRO_TRAILING(A1:A10)` is `A1:.A10`; the range's own `:` takes the
  // operator. An argument that is no range keeps the call spelling.
  std::string trim_ref(const AstNode& call, const AstNode& arg, TrimRefMode mode) const {
    if (arg.paren_depth() == 0U && arg.kind() == NodeKind::RangeOp) {
      // The chain's last `:` takes the operator.
      const std::uint32_t n = arg.as_range_arity();
      std::string out;
      for (std::uint32_t i = 0; i < n; ++i) {
        if (i > 0U) {
          out.append(i + 1U == n ? std::string_view(trim_ref_operator(mode)) : std::string_view(":"));
        }
        out.append(text(arg.as_range_operand(i), true));
      }
      return out;
    }
    if (arg.paren_depth() == 0U && arg.kind() == NodeKind::Ref) {
      const std::string ref(slice(arg.range().start, arg.range().end));
      if (const std::size_t colon = ref.rfind(':'); colon != std::string::npos) {
        return ref.substr(0, colon) + trim_ref_operator(mode) + ref.substr(colon + 1);
      }
    }
    const TextRange r = call.range();
    return std::string(slice(r.start, r.end));
  }

  std::string parenthesised(const AstNode& node, bool wrap) const {
    std::string inner = text(node, false);
    return wrap && (inner.empty() || inner.front() != '(') ? "(" + inner + ")" : inner;
  }

  std::string_view src_;
};

// Copies one structured-reference escape, quoted run or bracket, tracking
// `bracket_depth`; false when `formula[i]` starts none of them.
bool CopyOpaqueRun(std::string_view formula, std::size_t& i, std::size_t& bracket_depth, std::string& out) {
  const char c = formula[i];
  if (c == '\'' && bracket_depth > 0U) {
    out.append(formula.substr(i, 2));
    i += 2;
    return true;
  }
  if (c == '"' || c == '\'') {
    i = CopyQuotedRun(formula, i, c, out);
    return true;
  }
  if (c == '[') {
    ++bracket_depth;
  } else if (c == ']' && bracket_depth > 0U) {
    --bracket_depth;
  } else {
    return false;
  }
  out.push_back(c);
  ++i;
  return true;
}

}  // namespace

std::string spell_storage_operators(std::string_view formula) {
  const std::string_view body = strip_formula_prefix(formula);
  const std::size_t body_at = formula_prefix_len(formula);
  if (!strings::case_insensitive_contains(body, "SINGLE(") &&
      !strings::case_insensitive_contains(body, "ANCHORARRAY(") && !strings::case_insensitive_contains(body, "_TRO_")) {
    return std::string(formula);
  }
  Arena arena;
  const AstNode* root = parse_strict(body, arena);
  if (root == nullptr || !HasStorageOperatorCall(*root)) {
    return std::string(formula);
  }
  const OperatorRespeller respeller(body);
  const TextRange r = root->range();
  const std::string_view whole = respeller.slice(0, r.end);
  const std::string_view lead = respeller.slice(0, r.start);
  return std::string(formula.substr(0, body_at + lead.size())) + respeller.text(*root, false) +
         std::string(body.substr(whole.size()));
}

std::string_view cell_shaped_name_prefix(std::string_view name) noexcept {
  static constexpr std::array<std::string_view, 2> kPrefixes = {kParamPrefix, kNamePrefix};
  for (const std::string_view prefix : kPrefixes) {
    if (name.size() > prefix.size() && strings::case_insensitive_starts_with(name, prefix) &&
        is_a1_cell_shaped(name.substr(prefix.size()))) {
      return prefix;
    }
  }
  return {};
}

std::string_view cell_shaped_name_bare(std::string_view name) noexcept {
  return name.substr(cell_shaped_name_prefix(name).size());
}

std::string strip_storage_prefixes(std::string_view formula, bool (*known)(std::string_view)) {
  // Longest first so `_xlfn._xlws.` is consumed as a unit before `_xlfn.`.
  static constexpr std::array<std::string_view, 5> kPrefixes = {kFnWsPrefix, kFnPrefix, kWsPrefix, kParamPrefix,
                                                                kEtaPrefix};
  std::string out;
  out.reserve(formula.size());
  for (std::size_t i = 0; i < formula.size();) {
    const char c = formula[i];
    if (c == '"') {
      const std::size_t end = quoted_run_end(formula, i, '"');
      out.append(formula.substr(i, end - i));
      i = end;
      continue;
    }
    if (c == '_') {
      const bool at_boundary = out.empty() || !formula_prefix_is_ident_byte(out.back());
      if (at_boundary && strings::case_insensitive_starts_with(formula.substr(i), kOptionalParamPrefix)) {
        const std::size_t end = ident_run_end(formula, i + kOptionalParamPrefix.size());
        const std::string_view name =
            formula.substr(i + kOptionalParamPrefix.size(), end - i - kOptionalParamPrefix.size());
        out.push_back('[');
        if (is_a1_cell_shaped(name)) {
          out.append(kParamPrefix);
        }
        out.append(name);
        out.push_back(']');
        i = end;
        continue;
      }
      if (at_boundary) {
        bool matched = false;
        for (const std::string_view prefix : kPrefixes) {
          if (strings::case_insensitive_starts_with(formula.substr(i), prefix)) {
            const std::size_t end = ident_run_end(formula, i + prefix.size());
            const std::string_view name = formula.substr(i + prefix.size(), end - i - prefix.size());
            if (prefix == kParamPrefix && is_a1_cell_shaped(name)) {
              break;
            }
            if (prefix != kParamPrefix && prefix != kEtaPrefix && known != nullptr && !known(name)) {
              break;
            }
            i += prefix.size();
            matched = true;
            break;
          }
        }
        if (matched) {
          continue;
        }
      }
    }
    out.push_back(c);
    ++i;
  }
  return out;
}

std::string display_cell_shaped_names(std::string_view formula) {
  std::string out;
  out.reserve(formula.size());
  std::size_t bracket_depth = 0;
  for (std::size_t i = 0; i < formula.size();) {
    if (CopyOpaqueRun(formula, i, bracket_depth, out)) {
      continue;
    }
    const char c = formula[i];
    if (c == '_' && (out.empty() || !formula_prefix_is_ident_byte(out.back()))) {
      const std::size_t end = ident_run_end(formula, i);
      const std::string_view run = formula.substr(i, end - i);
      if (end < formula.size() && formula[end] == '(' && run.size() > kUdfPrefix.size() &&
          strings::case_insensitive_starts_with(run, kUdfPrefix)) {
        out.append(run.substr(kUdfPrefix.size()));
        i = end;
        continue;
      }
      const std::string_view prefix = cell_shaped_name_prefix(run);
      if (!prefix.empty()) {
        const std::string_view bare = run.substr(prefix.size());
        if (prefix == kNamePrefix) {
          out.push_back('\'');
          out.append(bare);
          out.push_back('\'');
        } else {
          out.append(bare);
        }
        i = end;
        continue;
      }
      if (bracket_depth == 0U && run.size() > kNamePrefix.size() &&
          strings::case_insensitive_starts_with(run, kNamePrefix)) {
        out.append(run.substr(kNamePrefix.size()));
        i = end;
        continue;
      }
    }
    out.push_back(c);
    ++i;
  }
  return out;
}

std::string spell_bare_calls_as_udf(std::string_view formula, bool (*reads_as_udf)(std::string_view)) {
  std::string out;
  out.reserve(formula.size());
  std::size_t bracket_depth = 0;
  for (std::size_t i = 0; i < formula.size();) {
    if (CopyOpaqueRun(formula, i, bracket_depth, out)) {
      continue;
    }
    const char c = formula[i];
    if (bracket_depth == 0U && formula_prefix_is_ident_byte(c) &&
        (out.empty() || (!formula_prefix_is_ident_byte(out.back()) && out.back() != '!'))) {
      const std::size_t end = ident_run_end(formula, i);
      const std::string_view run = formula.substr(i, end - i);
      if (end < formula.size() && formula[end] == '(' && run.front() != '_' && reads_as_udf(run)) {
        out.append(kUdfPrefix);
        out.append(strings::to_ascii_upper(run));
      } else {
        out.append(run);
      }
      i = end;
      continue;
    }
    out.push_back(c);
    ++i;
  }
  return out;
}

}  // namespace parser
}  // namespace formulon
