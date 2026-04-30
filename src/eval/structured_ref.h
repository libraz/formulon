// Copyright 2026 libraz. Licensed under the MIT License.
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
// See `backup/plans/04-xlsx-io.md` (table parts) and the ECMA-376 §18.5.1.10
// description of structured references for the full grammar.

#ifndef FORMULON_EVAL_STRUCTURED_REF_H_
#define FORMULON_EVAL_STRUCTURED_REF_H_

#include <cstdint>
#include <string>
#include <string_view>

#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

class Workbook;

namespace io {
struct TableMetadata;
}

namespace eval {

/// Bitmask of the area selectors (`#All`, `#Data`, `#Headers`, `#Totals`,
/// `#This Row` aka `@`) attached to a structured reference. A reference may
/// carry more than one selector (e.g. `Table[[#Headers],[Col]]`); a reference
/// with no `#`-prefixed selector and no leading `@` defaults to `kData`.
struct StructuredRefSpecifiers {
  static constexpr std::uint8_t kAll = 1u << 0u;
  static constexpr std::uint8_t kData = 1u << 1u;
  static constexpr std::uint8_t kHeaders = 1u << 2u;
  static constexpr std::uint8_t kTotals = 1u << 3u;
  static constexpr std::uint8_t kThisRow = 1u << 4u;
};

/// Parsed shape of a structured reference's bracket payload.
///
/// `table_name` is the table identifier preceding the brackets (or empty
/// when this selector represents the table-name slot of a cross-table range
/// like `Table[ColA]:Table[ColB]`). The view borrows from caller-owned
/// storage and must outlive the resolver call.
///
/// `column_first` and `column_last` capture the column slice. Both empty
/// (`column_first.empty()` is true) means "every column"; only
/// `column_first` set means a single-column slice; both set means the
/// inclusive `[first..last]` slice. Like `table_name`, both views borrow
/// from caller-owned storage. `column_last_set` distinguishes the
/// "no last column" case from the (rare but legal) empty-string column
/// name slot.
///
/// `specifiers` is the bitwise OR of the `StructuredRefSpecifiers::k*`
/// constants. The default (no `#` selector, no `@`) is `kData`.
struct StructuredRefSelector {
  std::string_view table_name;
  std::string_view column_first;
  std::string_view column_last;
  bool column_first_set = false;
  bool column_last_set = false;
  std::uint8_t specifiers = StructuredRefSpecifiers::kData;
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

/// Parses the bracket payload of a structured reference (the text between
/// the outermost `[` and `]`, exclusive). The payload is allowed to be
/// empty — that is the whole-table form `Table[]` which Excel accepts and
/// treats as `Table[#Data]`.
///
/// Recognised shapes:
///   * Empty                                     -> default (`kData`)
///   * `ColumnName`                              -> single column, default
///   * `@`                                       -> `kThisRow`, no column
///   * `@ColumnName`                             -> `kThisRow` + column
///   * `#All` / `#Data` / `#Headers` / `#Totals` -> selector only
///   * `#This Row`                               -> `kThisRow`
///   * `[#Headers],[Col]`                        -> selector + column
///   * `[#Headers],[ColA]:[ColB]`                -> selector + range
///   * `@[Col]` / `@[ColA]:[ColB]`               -> row-implicit range
///   * `[Col]` / `[ColA]:[ColB]`                 -> bracketed column form
///
/// On unknown specifier names, mismatched brackets, or any malformed
/// payload, returns `ErrorCode::Name` (Excel surfaces `#NAME?` for
/// unparseable structured references).
Expected<StructuredRefSelector, ErrorCode> parse_structured_ref_payload(std::string_view payload);

/// Looks up `selector.table_name` in `wb.tables()` (case-insensitive),
/// computes the row span from `selector.specifiers`, computes the column
/// span from `selector.column_first/column_last`, and returns the resolved
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
