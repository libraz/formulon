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

#include <string>

#include "pivot/pivot_cache.h"

namespace formulon::io {

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
std::string write_pivot_cache_definition(const pivot::PivotCache& cache);

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
std::string write_pivot_cache_records(const pivot::PivotCache& cache);

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_CACHE_WRITER_H_
