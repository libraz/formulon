// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// OOXML (.xlsx) package writer. Round-trips the workbook, including
// passive metadata (defined names, table parts) and parts the reader
// did not model — both `<Override>`-listed parts and Default-typed
// binary/media parts (vbaProject.bin, images, drawings, VML) captured
// as unknown-part passthrough. Cells are emitted with inline strings
// (`t="inlineStr"`); SST emission is intentionally not done here — the
// inline form round-trips cleanly already.

#ifndef FORMULON_IO_OOXML_WRITER_H_
#define FORMULON_IO_OOXML_WRITER_H_

#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Serialises `wb` into an in-memory OOXML byte stream.
///
/// The workbook content type emitted in `[Content_Types].xml` follows
/// `wb.kind()`: plain `.xlsx`, `.xlsm` (macro-enabled), `.xltx`
/// (template), or `.xltm` (macro-enabled template) all share the same
/// part schemas; only the workbook part's content type changes. The
/// engine never executes VBA — `xl/vbaProject.bin` (when present) is
/// carried verbatim through `wb.passthrough_parts()` and re-emitted by
/// the same passthrough path used for any other Override-listed part.
///
/// Always-emitted parts:
///   * `[Content_Types].xml`
///   * `_rels/.rels`
///   * `xl/workbook.xml`        — includes a `<definedNames>` block when
///                                 `wb.defined_names()` is non-empty.
///   * `xl/_rels/workbook.xml.rels`
///   * `xl/worksheets/sheet<N>.xml` (one per sheet, 1-based) — carries
///                                 a `<tableParts>` block when the sheet
///                                 owns at least one table.
///   * `xl/styles.xml`
///
/// Conditionally emitted parts:
///   * `xl/tables/table<N>.xml` — one per `wb.tables()` entry. The
///     numeric suffix follows the `TableMetadata::id`; tables with
///     `id == 0` get a per-write fallback id (logged via
///     `StructuredLog`).
///   * `xl/worksheets/_rels/sheet<N>.xml.rels` — for sheets that own
///     tables, pivot tables, hyperlinks, comments, printer settings, or
///     a DrawingML part; lists the corresponding relationships. A sheet
///     that references a drawing re-emits both the `<drawing r:id>`
///     element and the matching drawing relationship.
///   * Passthrough parts from `wb.passthrough_parts()` — written
///     verbatim, with their `<Override>` registration replicated when
///     `content_type` is non-empty. Default-typed passthrough parts
///     (empty `content_type` — vbaProject.bin, images, drawings, VML)
///     are written without an Override; their `<Default Extension>`
///     registration is round-tripped from `wb.default_content_types()`.
///
/// Sheet IDs and relationship IDs are assigned sequentially per sheet,
/// with the styles relationship following the last worksheet. Non-ASCII
/// sheet names are emitted as UTF-8 bytes with the five XML-critical
/// characters (`& < > " '`) escaped.
///
/// Collisions: if a passthrough part path matches a generated part path
/// (e.g. someone preserved `xl/styles.xml` verbatim), the generated
/// version wins; the passthrough copy is dropped and a warning is
/// logged via `StructuredLog`. This keeps the output well-formed even
/// when stale metadata is fed in.
///
/// Returns `FormulonErrorCode::kIoWriteFailed` on any miniz failure;
/// the error context identifies the offending part.
Expected<std::vector<std::uint8_t>, Error> write_ooxml(const Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WRITER_H_
