//
// Resolution of `NodeKind::ExternalRef` — a reference into a supporting
// workbook — against the values Excel cached in the external link part.
//
// Nothing here opens another file. Excel caches the cells a workbook
// actually references, so the cache holds an answer for every reference
// the file was saved with, and reading it reproduces what Excel itself
// shows when the source workbook is closed.
//
// Design references:
//   * `io/external_book.h` for the cache model and its lookup rules.

#ifndef FORMULON_EVAL_EXTERNAL_REF_H_
#define FORMULON_EVAL_EXTERNAL_REF_H_

#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

class EvalContext;

/// Resolves `node` (which must be `NodeKind::ExternalRef`) to its value.
///
/// A single cell yields a scalar; a rectangle — written out, or named by
/// a supporting-workbook defined name — yields an `Array`, so the
/// ordinary dynamic-array machinery spills it and range-taking functions
/// consume it without a dedicated path.
///
/// Error mapping:
///
/// | Condition                                        | Result    |
/// |--------------------------------------------------|-----------|
/// | No workbook bound to the context                 | `#REF!`   |
/// | `[N]` names no external link                     | `#REF!`   |
/// | Sheet absent from the supporting workbook        | `#REF!`   |
/// | Name absent from the supporting workbook         | `#NAME?`  |
/// | Name present but its target is not a rectangle   | `#REF!`   |
/// | Rectangle too large to materialise               | `#NUM!`   |
/// | Address the cache does not hold                  | `0`       |
///
/// The final row is Excel's own behaviour rather than a fallback: a
/// reference into a supporting workbook whose value Excel does not hold
/// reads as zero.
///
/// Text results are interned into `arena`, so the returned `Value` does
/// not borrow the workbook's cache.
Value resolve_external_ref(const parser::AstNode& node, Arena& arena, const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_EXTERNAL_REF_H_
