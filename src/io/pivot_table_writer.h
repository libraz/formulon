// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
//   * backup/plans/15-pivot-and-advanced.md §15.1.1 / §15.1.4
//   * src/io/pivot_table_reader.h (sister reader; canonical grammar)
//   * src/io/pivot_cache_writer.h (style precedent)

#ifndef FORMULON_IO_PIVOT_TABLE_WRITER_H_
#define FORMULON_IO_PIVOT_TABLE_WRITER_H_

#include <string>

#include "pivot/pivot_table.h"

namespace formulon::io {

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
/// Returns an owning `std::string`; there is no error channel because
/// the inputs are pure-data (no I/O, no allocation failure surfaced).
std::string write_pivot_table_definition(const pivot::PivotTable& table);

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_TABLE_WRITER_H_
