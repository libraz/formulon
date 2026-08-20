//
// Lazy impl for PHONETIC. Mac Excel reads the IME-typed kana from the
// OOXML `<rPh>` annotation attached to the source cell (either via the
// shared-strings entry the cell points at or via the cell's own inline-
// string block). Formulon plumbs that annotation through to the cell's
// `phonetic_runs` field; PHONETIC's lazy impl looks it up directly off
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

#include <string>
#include <string_view>
#include <vector>

#include "phonetic.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Returns what Excel's `PHONETIC` surfaces for the string `surface`
/// annotated with `runs`: each run's span is replaced by that run's kana
/// and everything outside every span is copied through unchanged. A
/// whole-string run therefore yields the kana alone, while a run covering
/// a prefix leaves the remainder in place -- `東京都` annotated with one
/// `<rPh sb="0" eb="2">トウキョウ</rPh>` surfaces `トウキョウ都`.
///
/// `runs` is expected in document order (non-decreasing `sb`), which is
/// what Excel emits and what the readers preserve. Malformed input is
/// absorbed rather than rejected: a run starting at or before the current
/// position has its kana emitted where the walk stands, an `eb` past the
/// end of `surface` consumes what is left, and a run whose span is empty
/// acts as an insertion.
std::string compose_phonetic(std::string_view surface, const std::vector<PhoneticRun>& runs);

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
