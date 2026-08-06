//
// Workbook-oracle test harness: turns a declarative workbook + pivot spec
// (the `spec` block of a workbook golden case) into in-memory engine
// objects so the pivot evaluator / layout pipeline can be exercised
// without a Windows-Excel golden being present.
//
// This file is deliberately NOT in `src/`: it is a test helper. It depends
// on `tests/oracle/json_reader.h` for its declarative input, which is the
// tiny in-test JSON parser and must not leak into the engine.
//
// Declarative `pivot` block shape
// -------------------------------
// A workbook case `spec` may carry an optional `pivot` object. When it
// does, `build_pivot_from_spec` understands the following fields:
//
//   {
//     "source":   "Data!A1:C13",       // sheet-qualified A1 range; row 0 is
//                                      // the header row, every later row is
//                                      // one source record.
//     "anchor":   "Sheet2!A1",         // sheet-qualified A1 top-left of the
//                                      // rendered pivot.
//     "row_fields":  ["Region"],       // source header names on the row axis
//     "col_fields":  ["Product"],      // source header names on the col axis
//     "data_fields": [                 // value-axis aggregations
//       {"field": "Amount", "agg": "Sum"}
//     ],
//     "layout":   "Compact",           // Compact | Tabular | Outline
//     "grand_totals": {"rows": true, "cols": true},
//     "filters":  [                    // optional manual item filters
//       {"field": "Region", "hide": ["South"]}
//     ]
//   }
//
// `agg` accepts: Sum, Count, Average, Max, Min, Product, CountNumbers,
// StdDev, StdDevP, Var, VarP. All fields except `source`, `anchor`,
// `data_fields` are optional; `row_fields` / `col_fields` / `filters`
// default to empty, `layout` to Compact, `grand_totals` to {true,true}.
//
// Declarative `print` block shape
// -------------------------------
// A workbook case `spec` may carry an optional `print` object. When it
// does, `build_print_from_spec` understands the following fields:
//
//   {
//     "sheet":       "Sheet1",          // REQUIRED: which sheet to
//                                       // paginate (0-based index is
//                                       // resolved from the name).
//     "print_area":  "A1:H80",          // optional; an A1 range, or a
//                                       // comma-separated multi-area
//                                       // range. Absent => the
//                                       // pagination engine falls back
//                                       // to the sheet's used range.
//     "print_titles": {                 // optional repeat rows/cols
//       "rows": "1:1",                  // 1-based Excel row span
//       "cols": "A:A"                   // column-letter span
//     },
//     "page_setup": {                   // optional; absent fields keep
//                                       // the OOXML/Excel defaults.
//       "orientation": "portrait",      // portrait | landscape | default
//       "paper":       9,               // OOXML paperSize code
//       "scale":       100,             // percentage
//       "fit_to_width":  1,             // pages; non-zero => fit-to-page
//       "fit_to_height": 0              // pages; non-zero => fit-to-page
//     },
//     "manual_breaks": {                // optional manual page breaks
//       "rows": [40],                   // 1-based Excel row numbers; a
//                                       // break is placed BEFORE the row
//       "cols": ["D"]                   // column letters; a break is
//                                       // placed BEFORE the column
//     }
//   }
//
// Row / column units: `print_titles.rows` and `manual_breaks.rows` use
// 1-based Excel row numbers; `print_titles.cols` and `manual_breaks.cols`
// use column letters. A `manual_breaks` entry of row 40 / column "D"
// places a break before that track (Excel's "insert page break"
// semantics), so the builder stores the 0-based index `40 - 1 == 39` /
// `column("D") == 3` as the `ManualBreak::id`.
//
// When `page_setup.fit_to_width` or `page_setup.fit_to_height` is
// non-zero the builder sets `PageSetup::fit_to_page = true` so the
// pagination engine derives a shrink factor; an explicit `scale` is then
// ignored, mirroring Excel's mutually-exclusive scale-vs-fit toggle.
//
// Case-level `column_widths` / `row_heights` (shared with the pivot
// path) feed `Sheet::mutable_layout()`: a `"A:D"` width key maps to one
// `ColumnLayout{first,last,width}`, a `"3"` row-height key to one
// `RowLayout{row,height}`.
//
// The `expect.print` block the workbook-oracle verifier diffs has shape:
//
//   { "print": {
//       "print_area": "A1:H80",   // resolved area as an A1 string
//       "h_breaks":   [40, 80],   // 0-based row indices each h-break
//                                 // precedes
//       "v_breaks":   [4],        // 0-based col indices each v-break
//                                 // precedes
//       "pages":      6           // total physical page count
//   } }

#ifndef FORMULON_TESTS_ORACLE_WORKBOOK_BUILDER_H_
#define FORMULON_TESTS_ORACLE_WORKBOOK_BUILDER_H_

#include <cstdint>
#include <memory>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "tests/oracle/json_reader.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace tests {
namespace oracle {

/// The in-memory objects produced from a declarative workbook + pivot
/// spec. `workbook` is heap-owned so that any `Value::text` views the
/// cache / table reference into the workbook's text storage stay valid
/// for the lifetime of the `BuiltPivot`.
struct BuiltPivot {
  std::unique_ptr<Workbook> workbook;
  pivot::PivotCache cache;
  pivot::PivotTable table;
};

/// Builds a `Workbook` + `PivotCache` + `PivotTable` from a declarative
/// workbook case `spec`.
///
/// `spec` must be a JSON object carrying a `sheets` block (sheet-name ->
/// {A1 -> value}) and a `pivot` block (shape documented at the top of
/// this header). The returned `BuiltPivot` is ready to feed into
/// `pivot::evaluate` and `pivot::layout`.
///
/// Errors (all `FormulonErrorCode::kInvalidArgument`):
///   * `spec` is not an object, or has no `pivot` block.
///   * the `pivot` block is missing `source` / `anchor` / `data_fields`.
///   * a `source` / `anchor` address fails to parse.
///   * a declared row / col / data field name is not a source header.
///   * an `agg` string is not one of the recognised aggregation names.
Expected<BuiltPivot, Error> build_pivot_from_spec(const JsonValue& spec);

/// The in-memory objects produced from a declarative workbook + print
/// spec. `workbook` is heap-owned; `sheet_index` is the 0-based index of
/// the sheet the `print` block named, ready to feed into
/// `print::paginate`.
struct BuiltPrint {
  std::unique_ptr<Workbook> workbook;
  std::uint32_t sheet_index = 0;
};

/// Builds a `Workbook` from a declarative workbook case `spec` and
/// applies its `print` block, returning the workbook plus the 0-based
/// index of the sheet to paginate.
///
/// `spec` must be a JSON object carrying a `sheets` block (sheet-name ->
/// {A1 -> value}) and a `print` block (shape documented at the top of
/// this header). Optional case-level `column_widths` / `row_heights`
/// maps are applied into `Sheet::mutable_layout()`. The `print` block's
/// `print_area` / `print_titles` are installed as sheet-scoped
/// `_xlnm.Print_Area` / `_xlnm.Print_Titles` defined names; its
/// `page_setup` / `manual_breaks` are written into the sheet's
/// `SheetPrintSettings`. The returned `BuiltPrint` is ready to feed into
/// `print::paginate`.
///
/// Errors (all `FormulonErrorCode::kInvalidArgument`):
///   * `spec` is not an object, or has no `print` block.
///   * the `print` block is missing string `sheet`, or names an unknown
///     sheet.
///   * a `print_area` / `print_titles` / `manual_breaks` token fails to
///     parse.
///   * a `page_setup` field has the wrong JSON type.
Expected<BuiltPrint, Error> build_print_from_spec(const JsonValue& spec);

}  // namespace oracle
}  // namespace tests
}  // namespace formulon

#endif  // FORMULON_TESTS_ORACLE_WORKBOOK_BUILDER_H_
