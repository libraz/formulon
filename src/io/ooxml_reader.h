// Copyright 2026 libraz. Licensed under the MIT License.
//
// OOXML (.xlsx) package reader. The current slice (Bundle 2.1) extracts
// the workbook structure: sheet names and order. Cells, shared strings,
// styles and tables are parsed by follow-up bundles (2.2 - 2.5). Every
// part not consumed by this slice is recorded in
// `OoxmlReadResult::unknown_parts` so callers can detect "we have a part
// but did not load it" cases. Bundle 2.5 will switch the policy from
// "everything we did not parse" to "everything we did not recognise".
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.2 (package structure)
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/26-implementation-plan.md (Phase 2 sequencing)

#ifndef FORMULON_IO_OOXML_READER_H_
#define FORMULON_IO_OOXML_READER_H_

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
struct OoxmlReadResult {
  Workbook workbook;
  std::vector<std::string> unknown_parts;
};

/// Reads an OOXML (.xlsx) package from in-memory bytes.
///
/// The current slice consumes:
///   * `[Content_Types].xml`         — only validates that workbook +
///     worksheet content types are present (does not yet build a
///     content-type registry)
///   * `_rels/.rels`                 — root relationship lookup
///   * `xl/workbook.xml`             — sheet enumeration in document order
///   * `xl/_rels/workbook.xml.rels`  — sheet relationship target lookup
///
/// Cell data is **not** loaded by this slice: every `Sheet` returned is
/// empty. The reader still resolves each sheet's relationship target so
/// later bundles can layer the per-sheet `<sheetData>` parser on top.
///
/// Returns `FormulonErrorCode::kIoZipCorrupt` for archive-level failures,
/// `kIoXmlParse` for malformed XML, and `kIoRelationshipBroken` /
/// `kIoContentTypeInvalid` when required relationships or content types
/// are missing. An empty sheet list is rejected with `kIoSheetCorrupt`
/// (Excel does the same).
Expected<OoxmlReadResult, Error> read_ooxml(ByteSpan bytes);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_READER_H_
