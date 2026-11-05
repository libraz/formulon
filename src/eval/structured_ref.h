//
// Structured-reference resolver. Translates the bracket payload of an Excel
// table reference (`Table[Col]`, `Table[@Col]`, `Table[#All]`,
// `Table[[#Headers],[Col]]`, `Table[ColA]:Table[ColB]`, ...) into a concrete
// rectangle on a workbook sheet. The parser captures the raw bracket payload
// verbatim into a `StructuredRef` AST node; the evaluator hands that text and
// the table name to this resolver to produce the (sheet, row_first, col_first,
// row_last, col_last) span the rest of the engine consumes.
//
// Out of scope (deferred follow-up):
//   * Localised specifier names (Excel ja-JP `[#見出し]` etc.). The OOXML
//     formula text always carries the English specifiers; localisation is a
//     UI concern, not a formula-text one.
//   * Cross-table row context inheritance. `Table[@Col]` resolves against the
//     evaluator's `current_cell()` row regardless of which table currently
//     owns that cell.
//   * Calculated columns spilling across a table column. Tables read by
//     `tables_reader` carry `display_name` and `ref` only; the formula
//     contents of calculated columns are tracked at the cell layer.
//
// See ECMA-376 §18.5.1.10 description of structured references for the
// full grammar.

#ifndef FORMULON_EVAL_STRUCTURED_REF_H_
#define FORMULON_EVAL_STRUCTURED_REF_H_

#include <cstdint>
#include <string>
#include <string_view>

#include "parser/structured_ref_payload.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

class Workbook;
struct TableMetadata;

namespace eval {

/// Bitmask of the area selectors attached to a structured reference. A
/// reference with no `#`-prefixed selector and no leading `@` defaults to
/// `kData`.
using parser::StructuredRefSpecifiers;

/// A structured reference to resolve: the table identifier preceding the
/// brackets and the parsed bracket payload. `table_name` borrows from
/// caller-owned storage and must outlive the resolver call. Empty
/// `payload.specifiers` means data rows; empty `payload.columns` means every
/// column; two columns are the ends of an inclusive range.
struct StructuredRefSelector {
  std::string_view table_name;
  parser::StructuredRefPayload payload;
};

/// Resolved rectangle on the table's home sheet. Coordinates are 0-based
/// and inclusive on both ends. `sheet_name` is the display name of the
/// owning sheet (case-preserving copy of `Sheet::name()`).
struct StructuredRefRange {
  std::string sheet_name;
  std::uint32_t sheet_index = 0;
  std::uint32_t row_first = 0;
  std::uint32_t col_first = 0;
  std::uint32_t row_last = 0;
  std::uint32_t col_last = 0;
};

/// Parses the text between a structured reference's outer brackets into a
/// selector (`table_name` left empty); see
/// `parser::parse_structured_ref_payload_text` for the accepted grammar. A
/// malformed payload returns `ErrorCode::Name` (Excel surfaces `#NAME?`).
Expected<StructuredRefSelector, ErrorCode> parse_structured_ref_payload(std::string_view payload);

/// Looks up `selector.table_name` in `wb.tables()` (case-insensitive),
/// computes the row span from `selector.payload.specifiers`, computes the column
/// span from `selector.payload.columns`, and returns the resolved
/// rectangle.
///
/// `current_sheet_index` is the 0-based workbook sheet index of the cell
/// currently being evaluated. `current_row` is its 0-based row, used only
/// when `kThisRow` is set; when `kThisRow` is set but `current_row` is
/// outside the table's data row span, returns `ErrorCode::Value`
/// (matches Excel's `#VALUE!`).
///
/// Error mapping:
///   * Unknown `table_name`                     -> `ErrorCode::Name`
///   * Unknown column name                      -> `ErrorCode::Ref`
///   * `kHeaders` on a table with no header row -> `ErrorCode::Ref`
///   * `kTotals` on a table with no totals row  -> `ErrorCode::Ref`
///   * `kThisRow` outside the table's data span -> `ErrorCode::Value`
///   * Malformed `ref` attribute on the table   -> `ErrorCode::Ref`
Expected<StructuredRefRange, ErrorCode> resolve_structured_ref(const StructuredRefSelector& selector,
                                                               const Workbook& wb, std::uint32_t current_sheet_index,
                                                               std::uint32_t current_row);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STRUCTURED_REF_H_
