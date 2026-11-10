//
// Writer pair for the OOXML pivot-cache parts. Symmetric counterpart of
// `src/io/pivot_cache_reader.{h,cpp}`:
//   * `xl/pivotCache/pivotCacheDefinition*.xml` <- field schema + shared items
//   * `xl/pivotCache/pivotCacheRecords*.xml`    <- per-row data
//
// Round-trip target: feeding the bytes produced here back into the reader
// must reproduce a `pivot::PivotCache` byte-equivalent on the relevant
// fields (`fields()` names + `shared_items` payloads, `records()` cell
// values). The writer does NOT emit the workbook-level
// `<pivotCache cacheId="..." r:id="..."/>` entry; that lives in the
// workbook part and is the writer-integration's responsibility (a later
// PR in the pivot-writer umbrella).
//
// Design references:
//   * src/io/pivot_cache_reader.h (sister reader; canonical grammar)

#ifndef FORMULON_IO_PIVOT_CACHE_WRITER_H_
#define FORMULON_IO_PIVOT_CACHE_WRITER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "value.h"

namespace formulon::io {

/// Write-time shape of one cache field.
struct PivotCacheFieldWritePlan {
  /// True when the writer emits the field as a discrete one with shared items
  /// built from its record values; false keeps the model's own encoding.
  bool derived = false;
  /// Distinct record values in first-appearance order (the shared items).
  std::vector<Value> items;
  /// `<sharedItems>` content hints: those of `items` for a derived field; for
  /// any other field with neither shared items nor captured hints, the hints
  /// to emit in place of the range placeholder when `present`.
  pivot::SharedItemsHints hints;
  /// Per item of `items`, whether it is a date serial; the mark of the first
  /// record cell holding that value.
  std::vector<bool> item_is_date;
  /// Write the field's numeric record values as `<d>` dates.
  bool dates = false;
  /// Per record, the index into `items` of its value.
  std::vector<std::uint32_t> record_index;
  /// Indices into `items` in ascending `value_less` order, blank last.
  std::vector<std::uint32_t> asc_order;
};

/// Per-field write plan of one cache, parallel to `cache.fields()`.
///
/// A cache built through the C API carries no shared items: text columns sit
/// inline in the records and the definition declares them numeric, which Excel
/// cannot open. The plan derives the discrete form at write time so the model
/// stays as authored.
struct PivotCacheWritePlan {
  std::vector<PivotCacheFieldWritePlan> fields;
  /// Emit `refreshOnLoad="1"`, replacing any captured value.
  bool refresh_on_load = false;
  /// The workbook's date system, for every `<d>` and `minDate` / `maxDate`.
  bool date1904 = false;
};

/// Builds the plan for `cache`; `tables` are the pivot tables drawing on it.
///
/// A field is derived only when it is a database field with no shared items,
/// no captured hints and no grouping, and either some table puts it on the
/// Row / Col / Page axis or authors items for it, or a record value is text,
/// boolean, error or blank. Any other field stays range-typed and takes the
/// hints derived from its record values and date marks (see
/// `PivotCacheFieldWritePlan::hints`); with no records it keeps the
/// `containsNumber="1"` placeholder.
///
/// The base of a date group with neither shared items nor captured hints gets
/// date hints (`minDate` at the earliest record date, `maxDate` at the day
/// after the latest, both at midnight), and its records are written as `<d>`.
///
/// `refresh_on_load` is set for a cache built in memory (no captured root
/// attributes); the saver adds the caches whose edited tables get no items.
///
/// `date1904` is the workbook's date system.
///
/// The caller must hold an `eval::ScopedEvalProfile`: values are deduplicated
/// and ordered with `pivot::value_less`.
PivotCacheWritePlan build_pivot_cache_write_plan(const pivot::PivotCache& cache,
                                                 const std::vector<const pivot::PivotTable*>& tables,
                                                 bool date1904 = false);

/// Emits a complete `xl/pivotCache/pivotCacheDefinition*.xml` document.
///
/// The result starts with the standard XML declaration, declares the
/// SpreadsheetML and relationships namespaces on the root, and lists one
/// `<cacheField>` per `cache.fields()` entry. For each field:
///   * `<sharedItems>` enumerates the field's `shared_items` when that
///     vector is non-empty. The encoding uses one child per item:
///     `<s>`, `<n>`, `<b>`, `<m/>`, `<e>` -- the same five elements the
///     reader recognises.
///   * An empty `shared_items` vector is treated as a range-typed field and
///     emits its captured hints, else the plan's derived hints, else a
///     placeholder `<sharedItems containsNumber="1"/>`. The reader picks this
///     up as "no per-item children" and leaves `shared_items` empty on the
///     round-trip; the records part then carries inline values for that
///     column.
///   * A shared field without captured hints gets the hints Excel writes for
///     its item kinds, dates (`shared_item_is_date`) included.
///
/// `cache.cache_id()` is intentionally NOT emitted: the cacheId attribute
/// lives on the workbook-level `<pivotCache>` entry, not in the part
/// itself. The function still takes `const PivotCache&` rather than just
/// `fields()` so the cache-source `<cacheSource type="worksheet"/>` can
/// later be enriched with a real source range as the workbook gains
/// source-tracking metadata.
///
/// `r:id="rId1"` on the root is a stable placeholder for the records
/// part; the cache-definition's own rels file (emitted by the writer-
/// integration PR) will define `rId1` as the records-part target. A cache
/// with `save_data() == false` has no records part, so the attribute is
/// omitted.
///
/// A calculated field (`formula`) is written as a childless `<cacheField>`;
/// `<calculatedItems>` follows the `trailing_children_xml()` run.
///
/// With a non-null `plan`, a derived field is written as a discrete field:
/// `<sharedItems>` lists the plan's items with its hints, and
/// `refreshOnLoad` follows `plan->refresh_on_load`. A null plan writes the
/// model as it stands, with `refreshOnLoad="1"` only for a cache built in
/// memory.
std::string write_pivot_cache_definition(const pivot::PivotCache& cache, const PivotCacheWritePlan* plan = nullptr);

/// Emits a complete `xl/pivotCache/pivotCacheRecords*.xml` document.
///
/// Each `pivot::PivotCacheRecord` becomes one `<r>` element. Per-cell
/// encoding walks parallel to `cache.fields()`:
///   * For a field whose `shared_items` is non-empty AND whose record
///     cell is `Value::number(N)`: emit `<x v="N"/>` (cast to integer,
///     no decimal). This matches how the reader resolves the index into
///     `shared_items`.
///   * Otherwise emit the inline-typed encoding (`<s>`, `<n>`, `<b>`,
///     `<m/>`, `<e>`) of the cell value.
///
/// The writer always emits exactly `cache.fields().size()` children per
/// `<r>`, padding with `<m/>` when the record holds fewer cells than the
/// field count. Excess cells beyond the field count are silently ignored
/// (the reader truncates symmetrically).
///
/// With a non-null `plan`, a derived field's cell is always the shared-item
/// index `<x v="plan.record_index[r]"/>`, whatever the model holds.
std::string write_pivot_cache_records(const pivot::PivotCache& cache, const PivotCacheWritePlan* plan = nullptr);

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_CACHE_WRITER_H_
