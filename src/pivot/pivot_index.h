// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
class PivotCache;

/// Returns the pivot table whose layout bounds contain `(row, col)` on
/// the sheet identified by `sheet_name`. Sheet name comparison is
/// case-insensitive (matches `EvalContext::resolve_ref` semantics).
/// Returns `nullptr` when no pivot covers that cell, the sheet is
/// unknown, or `wb` has no sheets.
const PivotTable* find_pivot_at_anchor(const Workbook& wb, std::string_view sheet_name, std::uint32_t row,
                                       std::uint32_t col) noexcept;

/// Fills in the names a pivot-table definition could not resolve at
/// read time because the bound cache had not been loaded yet.
///
/// The OOXML pivot-table part identifies its source columns positionally
/// (a `<pivotField>` matches the cache field at the same index) and its
/// items by a cache index (`<item x="N">`), never by name. This helper,
/// run once both the table and its `PivotCache` are in memory, fills:
///   * each `PivotField::source_name` that is still empty, from the
///     positionally-matching cache field's name; and
///   * each `PivotItem::name` that is still empty, from the cache field's
///     `shared_items[cache_index]` rendered the same way pivot labels are.
///
/// Names already set (e.g. through the C API, or a `<pivotField name=...>`
/// captured as `custom_name`) are left untouched. Fields the cache does
/// not cover (index out of range) are skipped. Without this, GETPIVOTDATA
/// on a loaded workbook cannot match a field by its source-column name.
void resolve_pivot_names(PivotTable& table, const PivotCache& cache);

/// Runs `resolve_pivot_names` for every pivot table in the workbook,
/// binding each to its cache via `Workbook::find_pivot_cache`. Intended
/// to be called once by the OOXML reader after both the pivot tables and
/// their caches are in memory, keeping the reader-side hook to a single
/// line. Tables whose cache id is unknown are left unresolved (no crash),
/// matching the reader's tolerant contract.
void resolve_all_pivot_names(Workbook& wb);

}  // namespace pivot
}  // namespace formulon

#endif  // FORMULON_PIVOT_PIVOT_INDEX_H_
