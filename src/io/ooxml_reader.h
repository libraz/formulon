// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// OOXML (.xlsx) package reader. The current slice extracts the workbook
// structure (sheet names + order) and per-sheet cell contents — every
// `<c>` in `xl/worksheets/sheet*.xml` is decoded into the workbook via
// `Workbook::set_cell_value` / `set_cell_formula`. Shared strings are
// resolved against the SST in-pipeline; styles, defined names, and
// table metadata are parsed for round-trip preservation (defined names
// land on `Workbook::defined_names()`, tables on `Workbook::tables()`,
// the writer slice emits both back unchanged). Every Override-listed
// part the reader did not consume is captured raw — path, content
// type, and bytes — into `OoxmlReadResult::unknown_parts` so the
// writer can re-emit them verbatim ("unknown-part passthrough"). The
// same payload is also stashed on the workbook via
// `Workbook::passthrough_parts()` so callers that only retain the
// workbook (e.g. `read_ooxml(...).workbook` on the right-hand side of
// an assignment) still get the round-trip guarantee.
//
// Default-typed binary parts (images, OLE objects) are NOT captured by
// this slice; only Override-listed parts round-trip. This is acceptable
// for v1 round-trip and matches the design's "minimal corpus" target.
//
// Macro-enabled (`.xlsm` / `.xltm`) and template (`.xltx`) packages are
// recognised by the workbook content type and detected purely as a
// `WorkbookKind` discriminator on the result; the reader treats them
// exactly like a plain `.xlsx` for cell content (the schemas are
// identical), and any `xl/vbaProject.bin` payload is captured via the
// existing passthrough mechanism so the writer can re-emit it verbatim.
// The engine NEVER executes VBA — preservation is byte-level only.
// Unrecognised workbook content types surface a structured-log warning
// (`ooxml.reader.unknown_workbook_content_type`) and fall back to
// `WorkbookKind::kXlsx` rather than failing the read; this is
// intentional Excel-compatibility-first behaviour.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.2 (package structure)
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/26-implementation-plan.md (Phase 2 sequencing,
//     Phase 4.3 xlsm + xltx detection)

#ifndef FORMULON_IO_OOXML_READER_H_
#define FORMULON_IO_OOXML_READER_H_

#include <cstdint>
#include <vector>

#include "io/passthrough_part.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Result of `read_ooxml`: a constructed (but not recalc'd) Workbook plus
/// a list of OOXML parts the reader did not consume. Unknown parts are
/// preserved with their raw bytes so the writer can round-trip them
/// unchanged.
///
/// `pending_sst_count` reports the number of `t="s"` cell references the
/// reader resolved against the shared-strings table during this read.
/// (Earlier slices used the same field to surface "still pending"
/// counts; with the SST reader wired in, every reference is resolved
/// in-pipeline and the field is purely an audit counter.) Cells that
/// previously held a `Text("")` placeholder now carry a `Value::text`
/// view into the corresponding SST entry, owned by the workbook's
/// `text_storage()` deque.
///
/// Text payloads (inline strings from `<is><t>...</t></is>`, SST
/// entries, phonetic runs) are interned into the workbook itself
/// (`Workbook::intern_text` / `mutable_text_storage`) so that
/// `Value::text` views remain valid for the lifetime of the workbook
/// even after the `OoxmlReadResult` is destroyed and the workbook is
/// moved out. Earlier slices stored the deque on this struct, which
/// caused a use-after-free when callers wrote
/// `Workbook wb = std::move(read_ooxml(...).value().workbook);` — the
/// result (and its deque) went out of scope at the end of the
/// statement while text views were still aliasing it.
///
/// `unknown_parts` carries the same passthrough payload that is also
/// copied onto the workbook via `set_passthrough_parts`. Both views are
/// populated; callers that only retain the workbook still get a
/// round-trip-clean writer pass.
struct OoxmlReadResult {
  Workbook workbook;
  std::vector<PassthroughPart> unknown_parts;
  std::uint32_t pending_sst_count = 0;
};

/// Reads an OOXML (.xlsx) package from in-memory bytes.
///
/// The current slice consumes:
///   * `[Content_Types].xml`         — only validates that workbook +
///     worksheet content types are present (does not yet build a
///     content-type registry)
///   * `_rels/.rels`                 — root relationship lookup
///   * `xl/workbook.xml`             — sheet enumeration in document order
///     plus `<definedNames>` metadata (preserved on the workbook for
///     round-trip)
///   * `xl/_rels/workbook.xml.rels`  — sheet, sharedStrings and styles
///     relationship target lookup
///   * `xl/worksheets/sheet*.xml`    — cell contents (literals + formulas)
///     decoded via `cell_parser` and dispatched through the public
///     Workbook API; the recalc engine is registered for every formula
///     cell. Cells with `t="s"` are resolved against the loaded SST.
///   * `xl/worksheets/_rels/sheet*.xml.rels` (when present) — walked
///     for table relationships; non-table relationships are deferred to
///     a later bundle.
///   * `xl/sharedStrings.xml`        — flat shared-string list; entries
///     are interned into `Workbook::text_storage()` and aliased by the
///     resolved cells.
///   * `xl/styles.xml`               — parsed for validation only (this
///     slice does not yet build a runtime style model).
///   * `xl/tables/table*.xml`        — table metadata (id, name, ref,
///     header/totals row, columns); preserved on the workbook for
///     round-trip but not yet wired into evaluation.
///
/// Returns `FormulonErrorCode::kIoZipCorrupt` for archive-level failures,
/// `kIoXmlParse` for malformed XML, and `kIoRelationshipBroken` /
/// `kIoContentTypeInvalid` when required relationships or content types
/// are missing. `kIoSheetCorrupt` surfaces for an empty sheet list, a
/// malformed `<c>` element, or an unresolved shared-formula reference.
Expected<OoxmlReadResult, Error> read_ooxml(ByteSpan bytes);

namespace internal {

/// Test hook for the OOXML rels-target path normaliser. Resolves `target`
/// against `base_dir`, collapsing `.` / `..` segments, and refuses any
/// input that escapes the package root (multi-level `..`, leading `/`).
/// Returned errors carry `kIoZipSlip`. Callers in production code use the
/// in-TU helper directly; this entry exists solely so the unit tests can
/// exercise the safety predicate in isolation.
Expected<std::string, Error> ResolveRelativePathForTesting(std::string_view base_dir, std::string_view target);

}  // namespace internal

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_READER_H_
