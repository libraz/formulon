// Copyright 2026 libraz. Licensed under the MIT License.
//
// Workbook-level pivot anchor resolution. Given a sheet identity + a
// cell address, find the PivotTable (if any) whose layout bounds
// contain that cell. Pure structural lookup -- no evaluation, no I/O.
//
// Used primarily by GETPIVOTDATA to identify which pivot table the
// caller's anchor argument addresses.
//
// Performance: linear scan over `Workbook::sheets()` and each sheet's
// `pivot_tables()`. Workbooks typically have fewer than 10 pivots; if
// hot, lift to a per-sheet R-tree later.

#ifndef FORMULON_PIVOT_PIVOT_INDEX_H_
#define FORMULON_PIVOT_PIVOT_INDEX_H_

#include <cstdint>
#include <string_view>

namespace formulon {
class Workbook;
class Sheet;
namespace pivot {

class PivotTable;

/// Returns the pivot table whose layout bounds contain `(row, col)` on
/// the sheet identified by `sheet_name`. Sheet name comparison is
/// case-insensitive (matches `EvalContext::resolve_ref` semantics).
/// Returns `nullptr` when no pivot covers that cell, the sheet is
/// unknown, or `wb` has no sheets.
const PivotTable* find_pivot_at_anchor(const Workbook& wb, std::string_view sheet_name, std::uint32_t row,
                                       std::uint32_t col) noexcept;

}  // namespace pivot
}  // namespace formulon

#endif  // FORMULON_PIVOT_PIVOT_INDEX_H_
