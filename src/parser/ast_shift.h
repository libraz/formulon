// Copyright 2026 libraz. Licensed under the MIT License.
//
// Relative-reference shifter for parsed formula ASTs.
//
// Excel formula sources that "anchor" at one cell and apply to a
// rectangle of others — conditional-format expression rules and shared
// formulas — need their A1 references re-aimed when applied to a cell
// other than the anchor. `shift_relative_refs` walks an `AstNode` tree
// and returns a new tree where every relative cell / range reference
// has been adjusted by `(row_delta, col_delta)`. Absolute references
// (`$A$1`, `$A1`, `A$1`) are preserved verbatim.
//
// Out-of-bounds shifts (a relative coordinate that would land before
// row 1, before column A, past row 1048576, or past column XFD)
// collapse the offending reference to `ErrorLiteral(#REF!)` so the
// evaluator surfaces the broken-reference error exactly as Excel does
// when a relative reference is shifted off the grid.
//
// Whole-column (`A:A`) and whole-row (`1:1`) references shift only
// along their meaningful axis — a whole-column ref shifts horizontally
// only, ignoring `row_delta`; a whole-row ref shifts vertically only.
//
// External-workbook references and structured (table) references are
// returned unchanged: they are addressed by name, not by relative
// coordinates. Lambda / Let bodies are walked recursively, but in
// practice CF and shared formulas do not use them.

#ifndef FORMULON_PARSER_AST_SHIFT_H_
#define FORMULON_PARSER_AST_SHIFT_H_

#include <cstdint>

#include "parser/ast.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {

/// Returns a new AST in which every relative reference inside `root`
/// has been shifted by `(row_delta, col_delta)`. The original tree is
/// not mutated. All allocations live in `arena`, which must outlive
/// the returned tree.
///
/// Returns `nullptr` only on arena exhaustion.
const AstNode* shift_relative_refs(const AstNode& root, Arena& arena, std::int32_t row_delta, std::int32_t col_delta);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_SHIFT_H_
