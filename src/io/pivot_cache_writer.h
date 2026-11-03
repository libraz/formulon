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
// Date grouping (`<fieldGroup>` / `<rangePr>`) and calculated cache fields
// remain out of scope at this layer, mirroring the reader's coverage.
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
  /// `<sharedItems>` content hints to emit alongside `items`.
  pivot::SharedItemsHints hints;
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
};

/// Builds the plan for `cache`; `tables` are the pivot tables drawing on it.
///
/// A field is derived only when it is a database field with no shared items,
/// no captured hints and no grouping, and either some table puts it on the
/// Row / Col / Page axis or authors items for it, or a record value is text,
/// boolean, error or blank. A purely numeric, blank-free field off every axis
/// keeps the range-typed placeholder.
///
/// The caller must hold an `eval::ScopedEvalProfile`: values are deduplicated
/// and ordered with `pivot::value_less`.
PivotCacheWritePlan build_pivot_cache_write_plan(const pivot::PivotCache& cache,
                                                 const std::vector<const pivot::PivotTable*>& tables);

/// Emits a complete `xl/pivotCache/pivotCacheDefinition*.xml` document.
///
/// The result starts with the standard XML declaration, declares the
/// SpreadsheetML and relationships namespaces on the root, and lists one
/// `<cacheField>` per `cache.fields()` entry. For each field:
///   * `<sharedItems>` enumerates the field's `shared_items` when that
///     vector is non-empty. The encoding uses one child per item:
///     `<s>`, `<n>`, `<b>`, `<m/>`, `<e>` -- the same five elements the
///     reader recognises.
///   * An empty `shared_items` vector is treated as a range-typed
///     (numeric) field and emits a placeholder `<sharedItems
///     containsNumber="1"/>`. The reader picks this up as "no per-item
///     children" and leaves `shared_items` empty on the round-trip; the
///     records part then carries inline values for that column.
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
/// integration PR) will define `rId1` as the records-part target.
///
/// With a non-null `plan`, a derived field is written as a discrete field:
/// `<sharedItems>` lists the plan's items with its hints. A null plan writes
/// the model as it stands.
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
