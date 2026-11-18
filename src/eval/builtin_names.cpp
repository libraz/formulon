#include "eval/builtin_names.h"

#include <cstdint>
#include <string>

#include "eval/function_registry.h"
#include "eval/special_forms_catalog.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/formula_prefix.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/name_syntax.h"
#include "utils/strings.h"

namespace formulon {
namespace eval {

namespace {

using parser::AstNode;
using parser::NodeKind;

const char* find_in(const char* const* names, std::string_view name) {
  for (const char* const* p = names; p != nullptr && *p != nullptr; ++p) {
    if (strings::case_insensitive_eq(name, *p)) {
      return *p;
    }
  }
  return nullptr;
}

bool is_qualified_builtin_callee(const AstNode& callee) {
  std::string_view name;
  if (callee.kind() == NodeKind::NameRef && !callee.as_name_sheet().empty()) {
    name = callee.as_name();
  } else if (parser::is_self_book_name_ref(callee)) {
    name = callee.as_external_ref_name();
  } else {
    return false;
  }
  return resolve_builtin_function_name(name) != nullptr;
}

// Whether parameter `index` of the lazy form `name` takes only a reference.
bool takes_only_reference(std::string_view name, std::uint32_t index) {
  const LazyEntry* entry = find_lazy_entry(name);
  return entry != nullptr && param_in_set(entry->ref_params, index);
}

// Built-ins whose arguments come in pairs: a count within the arity range is
// accepted when it equals `also`, or when the count after the first
// `pairs_from` arguments is even. Measured in Excel 365 by entering each
// built-in with every count up to 21.
struct ArgumentPairEntry {
  const char* name;
  std::uint32_t pairs_from;
  std::uint32_t also;
};

constexpr ArgumentPairEntry kArgumentPairs[] = {
    {"AVERAGEIFS", 1U, 0U}, {"COUNTIFS", 0U, 0U}, {"GETPIVOTDATA", 0U, 3U}, {"IFS", 0U, 0U},
    {"MAXIFS", 1U, 0U},     {"MINIFS", 1U, 0U},   {"SORTBY", 1U, 2U},       {"SUMIFS", 1U, 0U},
};

bool takes_argument_count(std::string_view name, std::uint32_t count) {
  std::uint32_t min_arity = 0U;
  std::uint32_t max_arity = 0U;
  if (!builtin_arity(name, &min_arity, &max_arity)) {
    return true;
  }
  if (count < min_arity || count > max_arity) {
    return false;
  }
  for (const ArgumentPairEntry& entry : kArgumentPairs) {
    if (strings::case_insensitive_eq(name, entry.name)) {
      return count == entry.also || (count - entry.pairs_from) % 2U == 0U;
    }
  }
  return true;
}

bool is_reference_expression(const AstNode& node) {
  switch (node.kind()) {
    case NodeKind::Ref:
    case NodeKind::Ref3D:
    case NodeKind::ExternalRef:
    case NodeKind::StructuredRef:
    case NodeKind::NameRef:
    case NodeKind::SpillRef:
    case NodeKind::LambdaCall:
    case NodeKind::LetBinding:
      return true;
    case NodeKind::ErrorLiteral:
      return node.as_error_literal() == ErrorCode::Ref;
    case NodeKind::ImplicitIntersection:
      return is_reference_expression(node.as_implicit_intersection_operand());
    case NodeKind::RangeOp:
      for (std::uint32_t i = 0; i < node.as_range_arity(); ++i) {
        if (!is_reference_expression(node.as_range_operand(i))) {
          return false;
        }
      }
      return true;
    case NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (!is_reference_expression(node.as_union_child(i))) {
          return false;
        }
      }
      return true;
    case NodeKind::IntersectOp:
      return is_reference_expression(node.as_intersect_lhs()) && is_reference_expression(node.as_intersect_rhs());
    case NodeKind::Call: {
      const std::string name = parser::strip_storage_prefixes(node.as_call_name());
      if (resolve_builtin_function_name(name) == nullptr) {
        return true;
      }
      const LazyEntry* entry = find_lazy_entry(name);
      return entry != nullptr && entry->ref_result != RefResult::kNone;
    }
    default:
      return false;
  }
}

}  // namespace

const char* resolve_builtin_function_name(std::string_view name) {
  if (const FunctionDef* def = default_registry().lookup(name); def != nullptr) {
    return def->canonical_name.data();
  }
  if (const char* lazy = find_in(lazy_form_names(), name); lazy != nullptr) {
    return lazy;
  }
  return find_in(parser_special_form_names(), name);
}

bool builtin_arity(std::string_view name, std::uint32_t* min_arity, std::uint32_t* max_arity) {
  if (find_lazy_arity(name, min_arity, max_arity)) {
    return true;
  }
  if (const FunctionDef* def = default_registry().lookup(name); def != nullptr) {
    *min_arity = def->min_arity;
    *max_arity = def->max_arity;
    return true;
  }
  return find_special_form_arity(name, min_arity, max_arity);
}

const AstNode* find_qualified_builtin_call(const AstNode& root) {
  switch (root.kind()) {
    case NodeKind::Literal:
    case NodeKind::Ref:
    case NodeKind::Ref3D:
    case NodeKind::ExternalRef:
    case NodeKind::StructuredRef:
    case NodeKind::NameRef:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
      return nullptr;
    case NodeKind::LambdaCall: {
      const AstNode& callee = root.as_lambda_call_callee();
      if (is_qualified_builtin_callee(callee)) {
        return &callee;
      }
      break;
    }
    default:
      break;
  }

  const AstNode* hit = nullptr;
  parser::any_child_node(root, [&](const AstNode& child) {
    hit = find_qualified_builtin_call(child);
    return hit != nullptr;
  });
  return hit;
}

const AstNode* find_non_reference_argument(const AstNode& root) {
  if (root.kind() == NodeKind::Call) {
    const std::string name = parser::strip_storage_prefixes(root.as_call_name());
    for (std::uint32_t i = 0; i < root.as_call_arity(); ++i) {
      const AstNode& arg = root.as_call_arg(i);
      if (takes_only_reference(name, i) && !is_reference_expression(arg)) {
        return &arg;
      }
    }
  }
  const AstNode* hit = nullptr;
  parser::any_child_node(root, [&](const AstNode& child) {
    hit = find_non_reference_argument(child);
    return hit != nullptr;
  });
  return hit;
}

const AstNode* find_wrong_arity_call(const AstNode& root) {
  if (root.kind() == NodeKind::Call) {
    const std::uint32_t count = root.as_call_arity();
    if (count > kMaxCallArguments ||
        !takes_argument_count(parser::strip_storage_prefixes(root.as_call_name()), count)) {
      return &root;
    }
  }
  const AstNode* hit = nullptr;
  parser::any_child_node(root, [&](const AstNode& child) {
    hit = find_wrong_arity_call(child);
    return hit != nullptr;
  });
  return hit;
}

const AstNode* find_refused_error_literal(const AstNode& root) {
  if (root.kind() == NodeKind::ErrorLiteral) {
    switch (root.as_error_literal()) {
      case ErrorCode::Calc:
      case ErrorCode::Spill:
      case ErrorCode::Field:
      case ErrorCode::Blocked:
      case ErrorCode::Connect:
      case ErrorCode::Busy:
      case ErrorCode::Unknown:
        return &root;
      default:
        break;
    }
  }
  const AstNode* hit = nullptr;
  parser::any_child_node(root, [&](const AstNode& child) {
    hit = find_refused_error_literal(child);
    return hit != nullptr;
  });
  return hit;
}

const AstNode* find_anchor_array_call(const AstNode& root) {
  if (root.kind() == NodeKind::Call &&
      strings::case_insensitive_eq(parser::strip_storage_prefixes(root.as_call_name()), parser::kAnchorArrayFn)) {
    return &root;
  }
  const AstNode* hit = nullptr;
  parser::any_child_node(root, [&](const AstNode& child) {
    hit = find_anchor_array_call(child);
    return hit != nullptr;
  });
  return hit;
}

const AstNode* find_refused_table_reference(const AstNode& root, char row_letter, char col_letter,
                                            bool (*is_table)(const void* ctx, std::string_view name), const void* ctx) {
  if (root.kind() == NodeKind::StructuredRef &&
      opens_with_r1c1(root.as_structured_ref_table(), row_letter, col_letter) &&
      is_table(ctx, root.as_structured_ref_table())) {
    return &root;
  }
  const AstNode* hit = nullptr;
  parser::any_child_node(root, [&](const AstNode& child) {
    hit = find_refused_table_reference(child, row_letter, col_letter, is_table, ctx);
    return hit != nullptr;
  });
  return hit;
}

namespace {

AstNode* parse_entry(std::string_view src, Arena& arena, parser::ParserOptions opts) {
  AstNode* root = parser::parse_strict(src, arena, opts);
  if (root == nullptr || find_qualified_builtin_call(*root) != nullptr) {
    return nullptr;
  }
  return root;
}

AstNode* parse_entry_reporting(std::string_view src, Arena& arena, parser::ParseFailure* out,
                               parser::ParserOptions opts) {
  AstNode* root = parser::parse_entry_with_diagnostics(src, arena, out, opts);
  if (root == nullptr) {
    return nullptr;
  }
  if (const AstNode* callee = find_qualified_builtin_call(*root); callee != nullptr) {
    out->code = parser::ParseErrorCode::EntryQualifiedBuiltin;
    out->range = callee->range();
    out->message = parser::default_message(out->code);
    return nullptr;
  }
  return root;
}

}  // namespace

AstNode* parse_formula_entry(std::string_view src, Arena& arena) {
  return parse_entry(src, arena, parser::ParserOptions{});
}

AstNode* parse_formula_entry_with_diagnostics(std::string_view src, Arena& arena, parser::ParseFailure* out) {
  return parse_entry_reporting(src, arena, out, parser::ParserOptions{});
}

AstNode* parse_name_body(std::string_view src, Arena& arena) {
  return parse_entry(src, arena, parser::name_body_options());
}

AstNode* parse_name_body_with_diagnostics(std::string_view src, Arena& arena, parser::ParseFailure* out) {
  return parse_entry_reporting(src, arena, out, parser::name_body_options());
}

}  // namespace eval
}  // namespace formulon
