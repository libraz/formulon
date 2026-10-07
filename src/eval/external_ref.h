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

#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/declared_rect.h"
#include "eval/tail_array.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

struct ExternalBook;

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Resolves `node` (which must be `NodeKind::ExternalRef`) to its value,
/// reading the cache of the link `Workbook::find_external_link` matches to
/// its book.
///
/// A single cell yields a scalar; a rectangle — written out, a whole column
/// or row at its declared size, or named by a
/// supporting-workbook defined name — yields an `Array`, so the ordinary
/// dynamic-array machinery spills it and range-taking functions consume it
/// without a dedicated path. A 3-D span in this scalar context is `#REF!`,
/// as a local 3-D reference is.
///
/// Error mapping:
///
/// | Condition                                        | Result    |
/// |--------------------------------------------------|-----------|
/// | No workbook bound to the context                 | `#REF!`   |
/// | The book matches no external link                | `#REF!`   |
/// | Sheet absent from the supporting workbook        | `#REF!`   |
/// | Sheet listed but without cached data             | `#REF!`   |
/// | 3-D span                                         | `#REF!`   |
/// | Name absent from the book or that sheet's scope  | `#NAME?`  |
/// | Name present but its target is not a rectangle   | `#REF!`   |
/// | Rectangle too large to materialise               | `#NUM!`   |
/// | Address the cache does not hold                  | blank     |
///
/// The final row is Excel's own behaviour rather than a fallback: a
/// reference into a supporting workbook whose value Excel does not hold
/// is an empty cell (a formula result displays it as `0`).
///
/// Text results are interned into `arena`, so the returned `Value` does
/// not borrow the workbook's cache.
Value resolve_external_ref(const parser::AstNode& node, Arena& arena, const EvalContext& ctx);

/// `resolve_external_ref` in the evaluator's internal result type. A
/// single-sheet whole column or row is a `TailArray` at its declared size:
/// the cached extent as the head, a reference-grid blank as the tail (head 0
/// when nothing is cached).
Shaped resolve_external_ref_shaped(const parser::AstNode& node, Arena& arena, const EvalContext& ctx);

/// A single-sheet cell-form `ExternalRef` read by coordinate rather than
/// materialised: the cached sheet it reads, the rectangle it declares (a
/// whole column spans every row), and the part of it a full read walks.
struct ExternalRect {
  const ExternalBook* book = nullptr;
  std::uint32_t sheet = 0;
  DeclaredRect declared;
  /// `declared` with a whole column or row clipped to the sheet's cached
  /// extent, or to its first line when the sheet caches no cell.
  DeclaredRect walked;
};

/// Resolves `node`, for which `external_ref_declared_endpoints` holds, to its
/// `ExternalRect`, with the book and sheet rows of the `resolve_external_ref`
/// error mapping.
Expected<ExternalRect, ErrorCode> resolve_external_rect(const parser::AstNode& node, const EvalContext& ctx);

/// Reads the cell at (`row`, `col`) of `book`'s sheet `sheet`, interning a
/// Text result into `arena`. An address the cache does not hold reads blank.
Value read_external_cell(const ExternalBook& book, std::uint32_t sheet, std::uint32_t row, std::uint32_t col,
                         Arena& arena);

/// Resolves the defined name `name` of the supporting workbook `book` in
/// `scope_sheet` (`ExternalBook::kNoSheet` for book scope), with the name
/// rows of the `resolve_external_ref` error mapping. A name the link records
/// as absent (`ExternalBookName::exists` false) is `#NAME?`.
Value resolve_external_book_name(const ExternalBook& book, std::uint32_t scope_sheet, std::string_view name,
                                 Arena& arena);

/// True for a 3-D reference: a local `Ref3D` or an `ExternalRef` spanning
/// sheets (`[Book]S1:S2!A1`). The shape-reading functions reject one with
/// `#VALUE!`.
bool is_three_d_reference(const parser::AstNode& node) noexcept;

/// Evaluates `arm`, the argument `IF` or `CHOOSE` selected. A 3-D reference
/// there is `#VALUE!` rather than the `#REF!` it reads as on its own, as in
/// Excel's dynamic-array entry of `=IF(TRUE,Sa:Sb!A1)`.
Value eval_selected_arm(const parser::AstNode& arm, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// For a 3-D `node` (`[Book]S1:S2!tail`), appends the tail's cells on every
/// sheet of the span in the link's sheet order, the way a range-taking
/// function reads a local 3-D reference. False (`#REF!`) when the book or an
/// endpoint sheet is unknown, or a sheet in the span has no cached data.
bool collect_external_ref3d_cells(const parser::AstNode& node, Arena& arena, const EvalContext& ctx,
                                  std::vector<Value>* out);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_EXTERNAL_REF_H_
