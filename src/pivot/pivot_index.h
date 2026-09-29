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

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
class Workbook;
class Sheet;
class Value;
namespace pivot {

class PivotTable;
class PivotCache;
struct PivotItem;

/// Returns the pivot table whose layout bounds contain `(row, col)` on
/// the sheet identified by `sheet_name`. Sheet name comparison uses
/// locale-independent Unicode simple case folding (matching
/// `EvalContext::resolve_ref`; this is not full Excel-equivalence).
/// Returns `nullptr` when no pivot covers that cell, the sheet is
/// unknown, or `wb` has no sheets.
const PivotTable* find_pivot_at_anchor(const Workbook& wb, std::string_view sheet_name, std::uint32_t row,
                                       std::uint32_t col) noexcept;

/// Resolves the cache value item `item` of field `field_index` binds to:
/// `shared_items[item.cache_index]` of the positionally-matching cache
/// field. Returns `nullptr` when the cache does not cover the field or the
/// index falls outside `shared_items`: such an item has no trustworthy
/// binding.
const Value* pivot_item_cache_value(const PivotCache& cache, std::size_t field_index, const PivotItem& item);

/// The label item `item` of field `field_index` is drawn and matched with:
/// its own `name` when set, else the display string of the cache value it
/// binds to, else empty. An item built by cache index carries no name
/// until a load resolves it, so evaluation reads labels through this
/// instead of `PivotItem::name`.
std::string pivot_item_label(const PivotCache& cache, std::size_t field_index, const PivotItem& item);

/// Fills in the names a pivot-table definition could not resolve at
/// read time because the bound cache had not been loaded yet.
///
/// The OOXML pivot-table part identifies its source columns positionally
/// (a `<pivotField>` matches the cache field at the same index) and its
/// items by a cache index (`<item x="N">`), never by name. This helper,
/// run once both the table and its `PivotCache` are in memory, fills:
///   * each `PivotField::source_name` that is still empty, from the
///     positionally-matching cache field's name; and
///   * each `PivotItem::name` that is still empty, from
///     `pivot_item_label`.
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
