// Copyright 2026 libraz. Licensed under the MIT License.
//
// Per-sheet `<sheetData>` reader. Walks the rows / cells of a parsed
// `sheet*.xml` document, decodes each `<c>` via `cell_parser`, and
// pushes the result into the workbook through the public
// `set_cell_value` / `set_cell_formula` API so the recalc engine sees
// the cell. Shared-formula bookkeeping (`<f t="shared" si="N">`) lives
// here; SST and styles handoff is out of scope until later bundles.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/26-implementation-plan.md (Phase 2.2)

#ifndef FORMULON_IO_SHEET_READER_H_
#define FORMULON_IO_SHEET_READER_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <tuple>
#include <vector>

#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Per-sheet reading context.
///
/// Shared formulas (`<f t="shared" si="N">`) reference each other by `si`
/// index within a single sheet, so the reader needs a scratch table to
/// resolve them. `pending_sst_cells` collects the addresses of cells
/// that carry a shared-string index (`t="s"`); the higher-level reader
/// resolves these against the workbook's shared-string pool once the
/// SST part has been loaded (Bundle 2.3).
struct SheetReadContext {
  /// Side table for SST-typed cells: list of (row, col, sst_index).
  /// Bundle 2.3 will iterate this list and replace the placeholder
  /// `Text("")` cells with the resolved string from the SST.
  std::vector<std::tuple<std::uint32_t, std::uint32_t, std::uint32_t>> pending_sst_cells;
};

/// Reads `<sheetData>` from a parsed `sheet*.xml` document and writes every
/// cell into `workbook.sheet(sheet_index)` via the public Workbook API
/// (`Workbook::set_cell_value` / `Workbook::set_cell_formula`), so the
/// recalc engine learns about each formula.
///
/// Behaviour:
///   * Literal cells (no `<f>`): routed through `Workbook::set_cell_value`.
///   * Formula cells: routed through `Workbook::set_cell_formula` with the
///     formula text (the leading `=` is added if not present, matching the
///     `Sheet::set_cell_formula` convention). The cached `<v>` is *not*
///     written back; the next `recalc()` populates it.
///   * Shared formulas (`<f t="shared" si="N" ref="A1:B5">...</f>`):
///       - the master occurrence (with formula body text) is recorded;
///       - later occurrences (`<f t="shared" si="N"/>`, no body) reuse the
///         master's formula text *verbatim*. No R1C1-style relative shift
///         is performed — Bundle 2.5 / Phase 4 will add proper adjustment.
///         For now, callers that hit a shared formula with relative refs
///         should expect the slave cells to evaluate to the master's
///         result; this is acceptable for the minimal-corpus tests this
///         slice covers, and is documented as a known limitation.
///   * Inline strings (`t="inlineStr"`): walked via `cell_parser`; rich-
///     text formatting is dropped (concatenated as plain text).
///   * SST cells (`t="s"`): the placeholder `Text("")` is written, and
///     `(row, col, sst_index)` is queued in `ctx.pending_sst_cells` for
///     a later resolution pass.
///
/// `sheet_doc` is the parsed pugixml document for this sheet; `sheet_index`
/// must be `< workbook.sheet_count()`. Returns `kIoSheetCorrupt` on any
/// malformed cell or on bookkeeping inconsistencies (a slave shared
/// formula that references an unknown `si` is a hard error).
///
/// `text_storage` is the workbook-lifetime backing-store for
/// inline-string `Value::text` payloads (a `std::deque<std::string>`
/// for pointer stability across appends). The reader appends each
/// decoded inline string here; cells store a `string_view` into the
/// resulting deque entry, so the deque must outlive the workbook's use
/// of the text values.
Expected<void, Error> read_sheet_data(const pugi::xml_document& sheet_doc, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage);

/// Sheet-XML byte size at which the OOXML reader switches from the
/// pugixml DOM path to the streaming SAX scanner. Below this threshold
/// the DOM fits comfortably in memory and the per-cell pugixml
/// overhead is amortised; above it the DOM grows linearly with cell
/// count and the SAX path is preferred.
///
/// The native default (256 KiB) is a conservative empirical pick: a
/// worksheet with ~10000 cells is typically ~150-200 KB, well below
/// the threshold; a 1M-cell sheet is multi-MB and obviously above it.
///
/// On WASM we set the threshold to `SIZE_MAX` so the SAX path is
/// statically unreachable; the linker then dead-code-eliminates the
/// streaming scanner and shrinks the binary by ~15-20 KiB. WASM
/// callers that need streaming reads can recompile with
/// `-DFORMULON_WASM_ENABLE_SAX=1` (see `cmake/FormulonWasm.cmake`).
#if defined(FORMULON_WASM) && !defined(FORMULON_WASM_ENABLE_SAX)
constexpr std::size_t kSaxThresholdBytes = static_cast<std::size_t>(-1);
#else
constexpr std::size_t kSaxThresholdBytes = 256U * 1024U;
#endif

/// Streaming variant of `read_sheet_data`. Walks the sheet XML
/// directly off `sheet_xml.data` via the SAX scanner and writes cells
/// into the workbook one at a time, without building a DOM. Behaviour
/// is identical to the DOM path (same cell coverage, same shared-
/// formula bookkeeping, same SST queueing); only the underlying parser
/// differs.
///
/// Used by the OOXML reader when the sheet's raw XML is at least
/// `kSaxThresholdBytes`. The DOM path remains the default below the
/// threshold so the WASM build does not pay the SAX setup cost on
/// every small sheet.
Expected<void, Error> read_sheet_data_sax(ByteSpan sheet_xml, std::size_t sheet_index, Workbook& workbook,
                                          SheetReadContext& ctx, std::deque<std::string>& text_storage);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SHEET_READER_H_
