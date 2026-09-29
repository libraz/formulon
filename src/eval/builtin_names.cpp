#include "eval/builtin_names.h"

#include <cstdint>

#include "eval/function_registry.h"
#include "eval/special_forms_catalog.h"
#include "eval/tree_walker.h"
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

const AstNode* find_qualified_builtin_call(const AstNode& root) {
  auto first = [](const AstNode& a, const AstNode& b) {
    const AstNode* hit = find_qualified_builtin_call(a);
    return hit != nullptr ? hit : find_qualified_builtin_call(b);
  };
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
    case NodeKind::SpillRef: {
      const AstNode* anchor = root.as_spill_ref_anchor_expr();
      return anchor != nullptr ? find_qualified_builtin_call(*anchor) : nullptr;
    }
    case NodeKind::UnaryOp:
      return find_qualified_builtin_call(root.as_unary_operand());
    case NodeKind::ImplicitIntersection:
      return find_qualified_builtin_call(root.as_implicit_intersection_operand());
    case NodeKind::BinaryOp:
      return first(root.as_binary_lhs(), root.as_binary_rhs());
    case NodeKind::RangeOp:
      return first(root.as_range_lhs(), root.as_range_rhs());
    case NodeKind::IntersectOp:
      return first(root.as_intersect_lhs(), root.as_intersect_rhs());
    case NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < root.as_union_arity(); ++i) {
        if (const AstNode* hit = find_qualified_builtin_call(root.as_union_child(i)); hit != nullptr) {
          return hit;
        }
      }
      return nullptr;
    case NodeKind::Call:
      for (std::uint32_t i = 0; i < root.as_call_arity(); ++i) {
        if (const AstNode* hit = find_qualified_builtin_call(root.as_call_arg(i)); hit != nullptr) {
          return hit;
        }
      }
      return nullptr;
    case NodeKind::ArrayLiteral:
      for (std::uint32_t r = 0; r < root.as_array_rows(); ++r) {
        for (std::uint32_t c = 0; c < root.as_array_cols(); ++c) {
          if (const AstNode* hit = find_qualified_builtin_call(root.as_array_element(r, c)); hit != nullptr) {
            return hit;
          }
        }
      }
      return nullptr;
    case NodeKind::Lambda:
      return find_qualified_builtin_call(root.as_lambda_body());
    case NodeKind::LetBinding:
      for (std::uint32_t i = 0; i < root.as_let_binding_count(); ++i) {
        if (const AstNode* hit = find_qualified_builtin_call(root.as_let_binding_expr(i)); hit != nullptr) {
          return hit;
        }
      }
      return find_qualified_builtin_call(root.as_let_body());
    case NodeKind::LambdaCall: {
      const AstNode& callee = root.as_lambda_call_callee();
      if (is_qualified_builtin_callee(callee)) {
        return &callee;
      }
      if (const AstNode* hit = find_qualified_builtin_call(callee); hit != nullptr) {
        return hit;
      }
      for (std::uint32_t i = 0; i < root.as_lambda_call_arity(); ++i) {
        if (const AstNode* hit = find_qualified_builtin_call(root.as_lambda_call_arg(i)); hit != nullptr) {
          return hit;
        }
      }
      return nullptr;
    }
  }
  return nullptr;
}

}  // namespace eval
}  // namespace formulon
