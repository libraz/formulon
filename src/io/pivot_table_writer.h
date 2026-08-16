//
// Writer for one OOXML `xl/pivotTables/pivotTable*.xml` part. Symmetric
// counterpart of `src/io/pivot_table_reader.{h,cpp}`. Round-trip target:
// feeding the bytes produced here back into
// `read_pivot_table_definition` must reproduce a `pivot::PivotTable`
// byte-equivalent on the fields that part of the spec (name, cacheId,
// anchor, axis assignments, item visibility, row/col/data field order
// and aggregation).
//
// `<pageFields>`, `<formats>`, `<conditionalFormats>`, `<chartFormats>`,
// `<pivotTableStyleInfo>`, and `<extLst>` are intentionally NOT emitted;
// these will land via a separate passthrough mechanism in a follow-up
// PR. The reader silently skips them today, so the round-trip is closed
// for the subset both ends agree on.
//
// Design references:
//   * src/io/pivot_table_reader.h (sister reader; canonical grammar)
//   * src/io/pivot_cache_writer.h (style precedent)

#ifndef FORMULON_IO_PIVOT_TABLE_WRITER_H_
#define FORMULON_IO_PIVOT_TABLE_WRITER_H_

#include <cstdint>
#include <optional>
#include <string>

#include "pivot/pivot_table.h"

namespace formulon::io {

/// Extent, in cells, of the grid a pivot actually renders.
///
/// Supplied by the caller (which owns the evaluator and the layout
/// projection; this writer deliberately depends on the pivot *model*
/// only) so `<location ref>` can describe the report Excel will draw
/// rather than whatever span the model happens to carry.
struct PivotRenderedSpan {
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
};

/// Emits a complete `xl/pivotTables/pivotTable*.xml` document.
///
/// Output structure (only the elements the reader currently consumes):
///   * `<pivotTableDefinition xmlns="..." name="..." cacheId="N">`
///   * `<location ref="A1:B2"/>` — anchor in A1 range form
///   * `<pivotFields count="K">` — one `<pivotField>` per `table.fields()`
///       - `axis="axisRow|axisCol|axisPage"` for non-Value axes
///       - `dataField="1"` for Value-axis fields (matches what real
///         Excel files emit; the reader folds both `axisValues` and
///         "no axis attribute + dataField=1" to `PivotAxis::Value`).
///       - `name="..."` from `field.custom_name` (omitted when empty).
///       - `subtotalTop="1"` only when `field.subtotal_top` is true.
///       - `<items count="M">` only when `field.items` is non-empty.
///         Each item emits `<item x="I"/>` where `I` is the document-
///         order index, plus `h="1"` when `!item.visible`.
///   * `<rowFields count="K">` / `<colFields count="K">` — emitted only
///     when the corresponding order vector is non-empty.
///   * `<dataFields count="K">` — emitted only when non-empty. Each
///     `<dataField>` carries `name`, `fld`, `subtotal` (always emitted
///     for clarity, matching real Excel files), and an optional
///     `numFmtId` (passthrough of `data_field.number_format`).
///
/// `rendered_span`, when supplied, is the extent the caller projected for
/// this pivot. It bears on `<location ref>` in one direction only, as a
/// lower bound, and only for a span the reader did not decode
/// (`PivotTable::has_authored_span()`):
///
///   * A span read back from a file is Excel's own and is re-emitted
///     untouched, so a read -> write cycle cannot rewrite authored bytes on
///     the strength of our own projection.
///   * Otherwise the emitted range is the larger of the model's span and
///     the projection. Under-sizing is the failure that matters — Excel
///     terminates when it refreshes a pivot whose `ref` is smaller than the
///     grid it draws — and a model assembled in memory keeps whatever
///     placeholder span it was created with however many fields are added
///     afterwards. Over-sizing is harmless, so a caller that deliberately
///     reserved a wider range keeps it.
///
/// Passing it unconditionally is therefore safe — this function owns the
/// choice.
///
/// Returns an owning `std::string`; there is no error channel because
/// the inputs are pure-data (no I/O, no allocation failure surfaced).
std::string write_pivot_table_definition(const pivot::PivotTable& table,
                                         std::optional<PivotRenderedSpan> rendered_span = std::nullopt);

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_TABLE_WRITER_H_
