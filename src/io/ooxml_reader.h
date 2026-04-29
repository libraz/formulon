// Copyright 2026 libraz. Licensed under the MIT License.
//
// OOXML (.xlsx) package reader. The current slice extracts the workbook
// structure (sheet names + order) and per-sheet cell contents — every
// `<c>` in `xl/worksheets/sheet*.xml` is decoded into the workbook via
// `Workbook::set_cell_value` / `set_cell_formula`. Shared strings are
// not yet resolved (cells that carry `t="s"` write a `Text("")`
// placeholder and surface the SST index via
// `OoxmlReadResult::pending_sst_count`); styles, defined names, and
// tables are parsed by follow-up bundles (2.3 - 2.5). Every part not
// consumed by this slice is recorded in `OoxmlReadResult::unknown_parts`
// so callers can detect "we have a part but did not load it" cases.
// Bundle 2.5 will switch the policy from "everything we did not parse"
// to "everything we did not recognise".
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.2 (package structure)
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/26-implementation-plan.md (Phase 2 sequencing)

#ifndef FORMULON_IO_OOXML_READER_H_
#define FORMULON_IO_OOXML_READER_H_

#include <cstdint>
#include <deque>
#include <string>
#include <vector>

#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Result of `read_ooxml`: a constructed (but not recalc'd) Workbook plus
/// a list of OOXML parts the reader did not consume yet. Unknown parts
/// are preserved so a future round-trip slice (Bundle 2.5) can write
/// them back unchanged.
///
/// `pending_sst_count` reports the number of `t="s"` cell references the
/// reader resolved against the shared-strings table during this read.
/// (Earlier slices used the same field to surface "still pending"
/// counts; with the SST reader wired in, every reference is resolved
/// in-pipeline and the field is purely an audit counter.) Cells that
/// previously held a `Text("")` placeholder now carry a `Value::text`
/// view into the corresponding SST entry, which is itself owned by
/// `text_storage`.
///
/// `text_storage` holds the inline-string payloads decoded from
/// `<is><t>...</t></is>` cells. `Value::text` is non-owning, so the
/// engine-wide rule is "the workbook's text views are valid for as long
/// as their backing storage is alive". `text_storage` is that backing
/// storage for cells loaded via the OOXML reader: callers must keep the
/// `OoxmlReadResult` alive at least as long as they read text values
/// from `workbook`. (The recalc engine has the same per-pass arena
/// constraint internally — see the note in
/// `eval/recalc_engine.cpp` Phase 4b.) Once Bundle 2.3 introduces a
/// workbook-owned shared-string pool the storage will move there.
struct OoxmlReadResult {
  Workbook workbook;
  std::vector<std::string> unknown_parts;
  std::uint32_t pending_sst_count = 0;
  // `std::deque` is intentional: we need pointer/iterator stability so
  // cell `Value::text` views into earlier entries do not invalidate
  // when later strings are appended. `std::vector` would reallocate.
  // Using a list-of-strings instead of a single concatenated buffer
  // keeps lookup-free O(1) appends and avoids rebasing string_views on
  // every push_back.
  // (Public C++17 has no `pmr::monotonic_buffer_resource` portability
  // story we can lean on across all targets, so this stays simple.)
  // Iterator stability is the only guarantee we rely on; ordering and
  // duplicate-merging are explicitly NOT promised.
  // NOTE: header field is intentionally a plain forwarded type to keep
  // ABI surface predictable; if dependence on `<deque>` becomes a size
  // concern (per backup/plans/18-wasm-size-optimization.md) we can
  // PIMPL it later.
  std::deque<std::string> text_storage;
};

/// Reads an OOXML (.xlsx) package from in-memory bytes.
///
/// The current slice consumes:
///   * `[Content_Types].xml`         — only validates that workbook +
///     worksheet content types are present (does not yet build a
///     content-type registry)
///   * `_rels/.rels`                 — root relationship lookup
///   * `xl/workbook.xml`             — sheet enumeration in document order
///   * `xl/_rels/workbook.xml.rels`  — sheet, sharedStrings and styles
///     relationship target lookup
///   * `xl/worksheets/sheet*.xml`    — cell contents (literals + formulas)
///     decoded via `cell_parser` and dispatched through the public
///     Workbook API; the recalc engine is registered for every formula
///     cell. Cells with `t="s"` are resolved against the loaded SST.
///   * `xl/sharedStrings.xml`        — flat shared-string list; entries
///     are owned by `text_storage` and aliased by the resolved cells.
///   * `xl/styles.xml`               — parsed for validation only (this
///     slice does not yet build a runtime style model).
///
/// Returns `FormulonErrorCode::kIoZipCorrupt` for archive-level failures,
/// `kIoXmlParse` for malformed XML, and `kIoRelationshipBroken` /
/// `kIoContentTypeInvalid` when required relationships or content types
/// are missing. `kIoSheetCorrupt` surfaces for an empty sheet list, a
/// malformed `<c>` element, or an unresolved shared-formula reference.
Expected<OoxmlReadResult, Error> read_ooxml(ByteSpan bytes);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_READER_H_
