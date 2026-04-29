// Copyright 2026 libraz. Licensed under the MIT License.
//
// Minimal `xl/styles.xml` reader. This slice only validates the part is
// well-formed and counts the top-level `<cellXfs>` / `<numFmts>` blocks
// so the OOXML reader can mark the styles part consumed and stop
// surfacing it as `unknown_parts`. A full style-runtime model
// (numFmt/font/fill/border lookups for the formatter) lands in a later
// bundle when the print/format pipeline begins.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/26-implementation-plan.md (Phase 2.3)

#ifndef FORMULON_IO_STYLES_READER_H_
#define FORMULON_IO_STYLES_READER_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Lightweight summary of a parsed `xl/styles.xml`.
///
/// Phase 2.3 placeholder; populated to a richer shape (per-xf
/// numFmt/font/fill resolution, custom number-format strings, etc.) once
/// the formatter pipeline starts consuming it. For now the reader simply
/// records that the part parsed cleanly and exposes two coarse counts so
/// follow-up bundles can be written against a stable struct.
struct StylesTable {
  /// Number of `<xf>` children inside `<cellXfs>` (Excel's "format
  /// records" per cell). Zero when the element is absent.
  std::size_t cell_xfs_count = 0;
  /// Number of `<numFmt>` children inside `<numFmts>` (custom
  /// number-format definitions). Zero when the element is absent — the
  /// built-in formats Excel ships with are not enumerated here.
  std::size_t num_fmts_count = 0;
};

/// Parses an OOXML styles part.
///
/// Behaviour:
///   * Empty `<styleSheet/>` (no children) is valid and yields a
///     `StylesTable` with both counts set to zero.
///   * Children other than `<cellXfs>` and `<numFmts>` (e.g. `<fonts>`,
///     `<fills>`, `<borders>`, `<cellStyleXfs>`, `<dxfs>`,
///     `<tableStyles>`, `<extLst>`) are accepted but ignored at this
///     layer.
///
/// Errors:
///   * `kIoXmlParse` — pugixml could not parse the document.
///   * `kIoContentTypeInvalid` — the document parses but its root is not
///     `<styleSheet>` (treated as the wrong content type rather than a
///     malformed XML).
Expected<StylesTable, Error> read_styles(const std::vector<std::uint8_t>& styles_bytes);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_STYLES_READER_H_
