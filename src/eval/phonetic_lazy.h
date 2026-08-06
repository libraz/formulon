//
// Lazy impl for PHONETIC. Mac Excel reads the IME-typed kana from the
// OOXML `<rPh>` annotation attached to the source cell (either via the
// shared-strings entry the cell points at or via the cell's own inline-
// string block). Formulon plumbs that annotation through to the cell's
// `phonetic_text` field; PHONETIC's lazy impl looks it up directly off
// the un-evaluated argument's `Reference` AST so the eager dispatcher
// cannot flatten the cell to a Value before the kana is read.
//
// For non-Ref arguments PHONETIC eagerly evaluates the subtree and
// applies Mac's strict-text passthrough surface: text passes through
// unchanged, blank yields "", numeric / boolean / array / error values
// surface #N/A. Errors propagate through the eager arm before the
// passthrough fires.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// dispatch-table contract in `tree_walker.cpp`.

#ifndef FORMULON_EVAL_PHONETIC_LAZY_H_
#define FORMULON_EVAL_PHONETIC_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `PHONETIC(reference)` — returns the IME-typed kana annotation
/// attached to the referenced cell, or the cell's surface text when no
/// annotation is present. Non-text values surface `#N/A`; blanks
/// surface `""`. Whole-row / whole-column references surface `#VALUE!`.
/// Non-Ref arguments are eagerly evaluated and the same text /
/// blank / #N/A passthrough is applied to the resulting scalar (errors
/// propagate).
Value eval_phonetic_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_PHONETIC_LAZY_H_
