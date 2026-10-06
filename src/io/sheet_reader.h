//
// Per-sheet `<sheetData>` reader. Walks the rows / cells of a parsed
// `sheet*.xml` document, decodes each `<c>` via `cell_parser`, and
// pushes the result into the workbook through the public
// `set_cell_value` / `set_cell_formula` API so the recalc engine sees
// the cell. Shared-formula bookkeeping (`<f t="shared" si="N">`) lives
// here; SST and styles handoff is out of scope until later bundles.

#ifndef FORMULON_IO_SHEET_READER_H_
#define FORMULON_IO_SHEET_READER_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <tuple>
#include <unordered_map>
#include <vector>

#include "io/array_anchor_budget.h"
#include "io/package_diagnostics.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
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
  /// Dynamic-array anchors (`<f t="array" ref=...>`) discovered while
  /// scanning cells. Neither `read_sheet_data` nor `read_sheet_data_sax`
  /// registers these itself -- the caller must resolve `pending_sst_cells`
  /// first and only then pass this list to `RegisterArraySpills`. A
  /// phantom cell's cached value is captured into the spill region at
  /// registration time, so registering before SST resolution would freeze
  /// an SST-typed phantom's unresolved `Text("")` placeholder into the
  /// region permanently.
  std::vector<ArrayAnchor> array_anchors;
  /// `(row, col, cm)` of every formula cell carrying `cm=`. The caller keeps
  /// as dynamic-array formulas those whose `cm` names the XLDAPR entry of
  /// `xl/metadata.xml` (`apply_loaded_dynamic_array_marks`).
  std::vector<std::tuple<std::uint32_t, std::uint32_t, std::uint32_t>> cell_metadata;
};

/// Registers each of `anchors` as a spill region on `sheet`, so the cached
/// spill targets Excel wrote as bare `<v>` cells (e.g. F7 / F8 of a
/// `=SEQUENCE(3)` anchored at F6) do not read back as independent literals
/// that collide (`#SPILL!`) with the anchor's re-spill on recalc. The
/// phantom values are captured into the region and the underlying
/// non-anchor cells are blanked so recalc can overwrite the footprint
/// freely. A one-cell region preserves dynamic-array metadata with no
/// phantom cells to capture.
///
/// Callers must resolve every SST-typed cell within `anchors`' footprints
/// (see `SheetReadContext::array_anchors`) before calling this -- both
/// `read_sheet_data` and `read_sheet_data_sax` leave that ordering to the
/// caller rather than enforcing it internally.
Expected<void, Error> RegisterArraySpills(Sheet& sheet, const std::vector<ArrayAnchor>& anchors);

/// Reads `<sheetData>` from a parsed `sheet*.xml` document and writes every
/// cell into `workbook.sheet(sheet_index)` via the public Workbook API
/// (`Workbook::set_cell_value` / `Workbook::set_cell_formula`), so the
/// recalc engine learns about each formula.
///
/// Behaviour:
///   * Literal cells (no `<f>`): routed through `Workbook::set_cell_value`.
///   * Formula cells: routed through `Workbook::set_cell_formula` with the
///     formula text (the leading `=` is added if not present, matching the
///     `Sheet::set_cell_formula` convention). A non-blank cached `<v>` is
///     retained as the cell's value until a caller explicitly calls
///     `recalc()` — a workbook using functions this engine does not
///     implement would otherwise lose Excel's answer to a read-only
///     inspection or a save/load round-trip. Callers must therefore not
///     read a freshly loaded formula cell as an engine-computed result.
///   * Shared formulas (`<f t="shared" si="N" ref="A1:B5">...</f>`):
///       - the master occurrence (with formula body text) is recorded;
///       - later occurrences (`<f t="shared" si="N"/>`, no body) reuse the
///         master's formula text shifted by the slave's row/column offset,
///         matching Excel's relative-reference semantics for drag-filled
///         shared formulas. If the formula dialect cannot be parsed yet, the
///         reader falls back to the master's text so the workbook remains
///         loadable.
///   * Inline strings (`t="inlineStr"`): walked via `cell_parser`; rich-
///     text formatting is dropped (concatenated as plain text).
///   * SST cells (`t="s"`): the placeholder `Text("")` is written, and
///     `(row, col, sst_index)` is queued in `ctx.pending_sst_cells` for
///     a later resolution pass.
///   * `<f t="dataTable">` (What-If data table): the element carries no
///     body, so the cell falls back to its cached `<v>` like any other
///     formula-less cell; `diagnostics->skipped_feature_count` (when
///     `diagnostics` is non-null) counts the drop so it is never silent,
///     matching every other unmodelled-feature drop in this reader.
///
/// `sheet_doc` is the parsed pugixml document for this sheet; `sheet_index`
/// must be `< workbook.sheet_count()`. Returns `kIoSheetCorrupt` on any
/// malformed cell or on bookkeeping inconsistencies (a slave shared
/// formula that references an unknown `si` is a hard error). Returns
/// `kIoFileTooLarge` once the cumulative `Cell` storage the walk would
/// materialise -- including the gap-filled slots a styled-but-sparse row
/// forces between its written columns -- exceeds `max_cell_bytes`
/// (default `kMaxSheetLoadCellBytes`; tests inject a smaller value, the
/// same pattern `zip_reader.h`'s `open()` uses for its own budget).
///
/// `text_storage` is the workbook-lifetime backing-store for
/// inline-string `Value::text` payloads (a `std::deque<std::string>`
/// for pointer stability across appends). The reader appends each
/// decoded inline string here; cells store a `string_view` into the
/// resulting deque entry, so the deque must outlive the workbook's use
/// of the text values. `diagnostics` may be null (tests, and callers
/// that do not track read diagnostics).
Expected<void, Error> read_sheet_data(const pugi::xml_document& sheet_doc, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                      ReadDiagnostics* diagnostics = nullptr,
                                      std::uint64_t max_cell_bytes = kMaxSheetLoadCellBytes);

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
/// "Identical" is a load-bearing claim, not a summary: because the
/// dispatch above turns on sheet size alone, any divergence would make
/// the same bytes decode differently for reasons the file's author cannot
/// see. So for any sheet XML byte sequence — well-formed or not — the two
/// entry points produce the same workbook mutations and the same
/// success / failure. Both paths first observe pugixml's `parse_default`
/// XML conversion contract (predefined entities, decimal/lowercase-`x`
/// character references, PCDATA CR/CRLF normalization, and attribute
/// TAB/LF/CR-to-space normalization) before applying the shared lexical
/// helpers: `parse_a1_ref` for `r=` on `<c>`, `parse_xsd_nonneg_int`
/// (`io/xsd_int.h`) for `s=` / `si=` / `r=` on `<row>`,
/// `parse_xsd_bool` for `hidden=`, and `decode_cell_payload` for the
/// whole `<v>` / `<is>` body. The DOM path performs the XML conversion
/// while pugixml builds the tree; the SAX path preserves zero-copy views
/// when no conversion is needed and uses distinct scratch backing for
/// simultaneously exposed semantic attributes. The disposition of a
/// rejected value is per-attribute but shared across the paths: a bad
/// `si=` fails the sheet, a bad `s=` degrades to the schema default 0, a
/// bad `<row r=>` drops that row's layout override, and `customHeight=` is
/// surfaced but does not itself create a row-layout override.
///
/// Used by the OOXML reader when the sheet's raw XML is at least
/// `kSaxThresholdBytes`. The DOM path remains the default below the
/// threshold so the WASM build does not pay the SAX setup cost on
/// every small sheet.
///
/// `diagnostics` and `max_cell_bytes` mirror `read_sheet_data`'s
/// parameters of the same name, including the `<f t="dataTable">` skip
/// count and the cell-materialisation budget -- required for the two
/// paths' "identical" contract above to hold on diagnostics and on
/// rejection, not just on the workbook mutations of a load that succeeds.
Expected<void, Error> read_sheet_data_sax(ByteSpan sheet_xml, std::size_t sheet_index, Workbook& workbook,
                                          SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                          ReadDiagnostics* diagnostics = nullptr,
                                          std::uint64_t max_cell_bytes = kMaxSheetLoadCellBytes);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SHEET_READER_H_
