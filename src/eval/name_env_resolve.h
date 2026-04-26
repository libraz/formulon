// Copyright 2026 libraz. Licensed under the MIT License.
//
// Helper that "looks through" a `NameRef` AST node to its LET-bound source
// AST when the binding has range shape. Used by every site that decides
// argument expansion based on `node.kind()` (the eager dispatcher's
// `accepts_ranges` paths, `resolve_range_arg`, and the `eval_count_lazy`
// special form). Without it, the dispatch sees `NodeKind::NameRef` for a
// LET binding such as `=LET(r, A1:A3, SUM(r))` and falls back to the
// scalar-Value path, collapsing the range to its spill-anchor cell.
//
// The helper is intentionally side-effect free and cheap: a single linked
// walk of the `NameEnv` chain followed by a kind check. It is safe to call
// on any AST node — non-NameRef inputs (and unbound / Value-only bindings)
// pass through unchanged.

#ifndef FORMULON_EVAL_NAME_ENV_RESOLVE_H_
#define FORMULON_EVAL_NAME_ENV_RESOLVE_H_

#include "eval/name_env.h"
#include "parser/ast.h"
#include "utils/strings.h"

namespace formulon {
namespace eval {

/// Returns the effective AST node for `node` after walking through any
/// LET-bound `NameRef -> AST` chain. Stops as soon as the current node is
/// not a `NameRef`, the name is unbound, or the binding has no AST (the
/// scalar / Value-only case). The returned reference always refers to a
/// node that outlives the LET scope: bound ASTs come from the same parser
/// arena that owns the LET node itself.
inline const parser::AstNode& resolve_name_ast(const parser::AstNode& node, const NameEnv* env) noexcept {
  if (env == nullptr) {
    return node;
  }
  const parser::AstNode* current = &node;
  // Bound chains are practically shallow (LET depth in user formulas peaks
  // in the single digits). A bounded loop also defends against a future
  // pathological `LET(a, b, LET(b, a, ...))` without per-call cycle state.
  for (int hops = 0; hops < 32; ++hops) {
    if (current->kind() != parser::NodeKind::NameRef) {
      return *current;
    }
    const parser::AstNode* next = env->lookup_ast(current->as_name());
    if (next == nullptr) {
      return *current;
    }
    current = next;
  }
  return *current;
}

/// True when `node` is one of the AST shapes that the eager dispatcher and
/// `resolve_range_arg` treat as range-producing: a `RangeOp` (`A1:B2`), an
/// `ArrayLiteral` (`{1,2;3,4}`), or one of the reference-producing calls
/// (`OFFSET`, `CHOOSE`, `INDIRECT`, `IF`). Single-cell `Ref` is intentionally
/// excluded so that LET passthrough does not silently change the
/// scalar-vs-range provenance of a `=LET(r, A1, SUM(r,B1))` formula.
///
/// `IF` is included because Mac Excel preserves reference-shape through the
/// picked branch: `=LET(r, IF(TRUE, A1:A3, B1:B3), SUM(r))` evaluates as if
/// `r` were `A1:A3`. The companion logic in `resolve_range_arg` short-
/// circuits the condition and recurses into the chosen branch.
inline bool is_range_shaped_ast(const parser::AstNode& node) noexcept {
  switch (node.kind()) {
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::ArrayLiteral:
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
