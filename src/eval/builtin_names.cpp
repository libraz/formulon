#include "eval/builtin_names.h"

#include "eval/function_registry.h"
#include "eval/special_forms_catalog.h"
#include "eval/tree_walker.h"
#include "parser/parser.h"
#include "utils/arena.h"
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

AstNode* parse_formula_entry(std::string_view src, Arena& arena) {
  AstNode* root = parser::parse_strict(src, arena);
  if (root == nullptr || find_qualified_builtin_call(*root) != nullptr) {
    return nullptr;
  }
  return root;
}

}  // namespace eval
}  // namespace formulon
