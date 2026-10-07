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
//   * `external_book.h` for the cache model and its lookup rules.

#ifndef FORMULON_EVAL_EXTERNAL_REF_H_
#define FORMULON_EVAL_EXTERNAL_REF_H_

#include <vector>

#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

class EvalContext;

/// Resolves `node` (which must be `NodeKind::ExternalRef`) to its value,
/// reading the cache of the link `Workbook::find_external_link` matches to
/// its book.
///
/// A single cell yields a scalar; a rectangle — written out, a whole column
/// or row clipped to the sheet's cached extent, or named by a
/// supporting-workbook defined name — yields an `Array`, so the ordinary
/// dynamic-array machinery spills it and range-taking functions consume it
/// without a dedicated path. A 3-D span in this scalar context reads as a
/// local 3-D reference does.
///
/// Error mapping:
///
/// | Condition                                        | Result    |
/// |--------------------------------------------------|-----------|
/// | No workbook bound to the context                 | `#REF!`   |
/// | The book matches no external link                | `#REF!`   |
/// | Sheet absent from the supporting workbook        | `#REF!`   |
/// | Sheet listed but without cached data             | `#REF!`   |
/// | 3-D span whose every sheet is cached             | `#VALUE!` |
/// | Name absent from the book or that sheet's scope  | `#NAME?`  |
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

/// For a 3-D `node` (`[Book]S1:S2!tail`), appends the tail's cells on every
/// sheet of the span in the link's sheet order, the way a range-taking
/// function reads a local 3-D reference. False (`#REF!`) when the book or an
/// endpoint sheet is unknown, or a sheet in the span has no cached data.
bool collect_external_ref3d_cells(const parser::AstNode& node, Arena& arena, const EvalContext& ctx,
                                  std::vector<Value>* out);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_EXTERNAL_REF_H_
