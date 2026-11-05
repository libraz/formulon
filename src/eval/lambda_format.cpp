//
// Implementation of `format_lambda_value`. See the header for contract.

#include "eval/lambda_format.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/lambda_value.h"
#include "eval/name_env.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/formula_prefix.h"

namespace formulon {
namespace eval {

namespace {

// Adds each name `body` uses that `env` binds, so a captured LET name keeps
// its bound spelling.
void add_captured_names(const parser::AstNode& body, const NameEnv& env, std::vector<std::string_view>& scope) {
  std::string_view name;
  if (body.kind() == parser::NodeKind::NameRef && body.as_name_sheet().empty()) {
    name = body.as_name();
  } else if (body.kind() == parser::NodeKind::Call) {
    name = body.as_call_name();
  }
  if (!name.empty() && env.lookup(name) != nullptr) {
    scope.push_back(name);
  }
  parser::any_child_node(body, [&](const parser::AstNode& child) {
    add_captured_names(child, env, scope);
    return false;
  });
}

}  // namespace

std::string format_lambda_value(const LambdaValue& lv) {
  if (!lv.builtin.empty()) {
    return std::string(lv.builtin);
  }
  std::string out;
  out.append("LAMBDA(");
  const std::uint32_t first_optional = lv.param_count - lv.optional_count;
  std::vector<std::string_view> scope;
  for (std::uint32_t i = 0; i < lv.param_count; ++i) {
    if (i > 0U) {
      out.push_back(',');
    }
    scope.push_back(lv.params[i]);
    // The formula-bar spelling: a cell-shaped parameter shows bare.
    const std::string_view param = parser::cell_shaped_name_bare(lv.params[i]);
    if (i >= first_optional) {
      out.push_back('[');
      out.append(param);
      out.push_back(']');
    } else {
      out.append(param);
    }
  }
  if (lv.param_count > 0U) {
    out.push_back(',');
  }
  if (lv.body != nullptr) {
    if (lv.captured_env != nullptr) {
      add_captured_names(*lv.body, *lv.captured_env, scope);
    }
    out.append(parser::display_cell_shaped_names(parser::format_formula_in_scope(*lv.body, scope)));
  }
  out.push_back(')');
  return out;
}

}  // namespace eval
}  // namespace formulon
