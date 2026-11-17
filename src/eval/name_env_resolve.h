//
// Helper that "looks through" a LET / LAMBDA-bound `NameRef` to the `Literal`
// holding the `Ref` it stands for, for the sites that decide argument handling
// by `node.kind()`. Side-effect free and cheap: one walk of the `NameEnv`
// chain. Any other node, and a name bound to anything but a reference, passes
// through unchanged.

#ifndef FORMULON_EVAL_NAME_ENV_RESOLVE_H_
#define FORMULON_EVAL_NAME_ENV_RESOLVE_H_

#include "eval/name_env.h"
#include "parser/ast.h"
#include "utils/strings.h"

namespace formulon {
namespace eval {

/// Returns the `Literal` holding the `Ref` a LET / LAMBDA-bound `NameRef` `node` stands for; any other node, and a
/// name bound to anything but a reference, unchanged. A sheet-qualified name (`Sheet1!r`) is a workbook name, never
/// a binding.
inline const parser::AstNode& resolve_name_ast(const parser::AstNode& node, const NameEnv* env) noexcept {
  if (env == nullptr || node.kind() != parser::NodeKind::NameRef || !node.as_name_sheet().empty()) {
    return node;
  }
  const parser::AstNode* bound = env->lookup_ast(node.as_name());
  return bound != nullptr ? *bound : node;
}

/// True when `node` is one of the AST shapes read as range-producing: a
/// `RangeOp` (`A1:B2`), an `ArrayLiteral` (`{1,2;3,4}`), a `SpillRef` (`A1#`),
/// or one of the reference-producing calls (`OFFSET`, `CHOOSE`, `INDIRECT`,
/// `IF`). A single-cell `Ref` is excluded: it keeps its scalar provenance.
inline bool is_range_shaped_ast(const parser::AstNode& node) noexcept {
  switch (node.kind()) {
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::ArrayLiteral:
    case parser::NodeKind::SpillRef:
      return true;
    case parser::NodeKind::Call: {
      const auto name = node.as_call_name();
      return strings::case_insensitive_eq(name, "OFFSET") || strings::case_insensitive_eq(name, "CHOOSE") ||
             strings::case_insensitive_eq(name, "INDIRECT") || strings::case_insensitive_eq(name, "IF");
    }
    default:
      return false;
  }
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_NAME_ENV_RESOLVE_H_
