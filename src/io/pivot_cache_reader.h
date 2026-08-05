//
// Reader pair for the OOXML pivot-cache parts:
//   * xl/pivotCache/pivotCacheDefinition*.xml -> field schema + shared items
//   * xl/pivotCache/pivotCacheRecords*.xml    -> per-row data
//
// Produces a populated `formulon::pivot::PivotCache`. The workbook-side
// rels resolution (mapping cacheId integers to part paths) lives in the
// PivotTable definition reader (subsequent PR); this layer is pure XML ->
// struct.
//
// Date grouping (`<fieldGroup>` / `<rangePr>`) and calculated cache fields
// are deliberately out of scope at this layer: a follow-up PR will add
// real semantics. Today such fields land with an empty `shared_items`
// list and their records carry the raw `<x>` indices unchanged, which is
// enough to round-trip the workbook without losing data.
//
// Design references:
//   * src/io/tables_reader.h (closest precedent)

#ifndef FORMULON_IO_PIVOT_CACHE_READER_H_
#define FORMULON_IO_PIVOT_CACHE_READER_H_

#include <cstdint>
#include <vector>

#include "pivot/pivot_cache.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// Parses one `xl/pivotCache/pivotCacheDefinition*.xml` part.
///
/// Behaviour:
///   * The root must be `<pivotCacheDefinition>`. Anything else surfaces
///     `kIoContentTypeInvalid`.
///   * Each `<cacheField>` becomes a `pivot::PivotCacheField`. The
///     field's `name` attribute is captured verbatim; missing/empty
///     `name` does NOT fail (Excel writes anonymous fields for
///     synthesized data fields and we round-trip them).
///   * `<sharedItems>` children are walked in document order. Recognised
///     children:
///       - `<s v="...">` -> `Value::text(...)`
///       - `<n v="...">` -> `Value::number(strtod)`; non-numeric is
///         `kIoSheetCorrupt`.
///       - `<b v="0|1">` -> `Value::boolean(...)`
///       - `<m/>`        -> `Value::blank()`
///       - `<e v="...">` -> `Value::error(...)` (Excel error display name).
///       - any other element is silently skipped (forward compatibility
///         with Excel additions; the writer will preserve unknowns via
///         a separate passthrough path in a later PR).
///   * Range-typed sharedItems (no per-item children, just attributes
///     such as `containsNumber="1" minValue="..." maxValue="..."`) leave
///     `shared_items` empty. Records for such fields carry inline values
///     instead of `<x>` indices (see `read_pivot_cache_records`).
///   * `<cacheSource type="external"/>` -> `kIoContentTypeInvalid` (we
///     do not support external sources today).
///   * `cacheSource type` defaults to `"worksheet"` when absent (matching
///     the OOXML default).
///
/// Errors:
///   * `kIoXmlParse` — pugixml could not parse the bytes.
///   * `kIoContentTypeInvalid` — root not `<pivotCacheDefinition>`, or
///     external cache source.
///   * `kIoSheetCorrupt` — `<n v="...">` failed to parse as a double.
Expected<pivot::PivotCache, Error> read_pivot_cache_definition(const std::vector<std::uint8_t>& definition_bytes);

/// Populates the records of `cache` from one
/// `xl/pivotCache/pivotCacheRecords*.xml` part. The `cache` argument
/// must already be populated (call `read_pivot_cache_definition` first
/// and pass its result here).
///
/// Behaviour:
///   * The root must be `<pivotCacheRecords>`. Otherwise
///     `kIoContentTypeInvalid`.
///   * Each `<r>` becomes one `pivot::PivotCacheRecord`. Inside `<r>`,
///     children appear in field order (one child per `cacheField`).
///     Recognised children:
///       - `<x v="N">` -> `Value::number(N)`, the raw index into
///         `cache.fields()[i].shared_items` (out-of-bounds N ->
///         `kIoSheetCorrupt`). The index is stored as-is; callers resolve
///         it through `pivot::cell_value`.
///       - `<n v="...">` -> `Value::number(strtod)`; non-numeric is
///         `kIoSheetCorrupt`.
///       - `<s v="...">` -> `Value::text(...)`
///       - `<b v="0|1">` -> `Value::boolean(...)`
///       - `<m/>`        -> `Value::blank()`
///       - `<e v="...">` -> `Value::error(...)`
///   * The number of children per `<r>` does NOT have to equal the
///     number of cache fields. Excel sometimes elides trailing blank
///     columns; missing trailing values become `Value::blank()`.
///   * Excess children (more than the cache field count) are ignored
///     with no error (forward compatibility with Excel additions).
///
/// Errors:
///   * `kIoXmlParse`, `kIoContentTypeInvalid`, `kIoSheetCorrupt`.
Expected<void, Error> read_pivot_cache_records(const std::vector<std::uint8_t>& records_bytes,
                                               pivot::PivotCache& cache);

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_CACHE_READER_H_
