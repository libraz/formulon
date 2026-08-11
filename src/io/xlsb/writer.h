//
// MS-XLSB (.xlsb) package writer. The write-side counterpart to
// `io/xlsb/reader.h`: walks a `Workbook` and produces a complete OPC
// package (ZIP container of `[Content_Types].xml`, `_rels/.rels`,
// `xl/workbook.bin`, `xl/worksheets/sheet*.bin`, optional
// `xl/sharedStrings.bin`, plus any passthrough parts the workbook
// preserved from a previous read).
//
// Round-trip with the reader is the primary correctness contract:
// `read_xlsb(write_xlsb(wb))` reproduces every cell value `wb` carried
// in for the in-scope literal kinds (number, boolean, text, error,
// blank). Formula cells round-trip through a full AST→Ptg encoding
// pipeline (`io/xlsb/cell_writer.cpp`'s `EncodeCellFormula`, backed by
// `io/xlsb/ptg_writer.h`'s `encode_ptgs`): `cell.formula_text` is
// re-parsed and lowered directly to the `rgce` Ptg byte stream spliced
// into a `BrtFmla*` record. A formula that cannot be parsed or lowered
// is written as its cached literal value; callers receive an explicit
// downgrade count instead of losing the rest of the package.
//
// Styles, row/column layout, merged cells, pane state, and defined names are
// emitted from the model.
//
// Conditional formats, data validation, hyperlinks, auto-filter, print
// settings and tables are not lowered to records. A sheet read from an
// `.xlsb` keeps them as `Sheet::xlsb_tail()` and this writer re-emits those
// bytes verbatim (together with the sheet's own rels, so the ids they carry
// still resolve). A workbook built in memory or read from `.xlsx` has no such
// bytes; its modelled equivalents are counted in `deferred_feature_count` and
// logged under `xlsb.writer.deferred` rather than being dropped silently.

#ifndef FORMULON_IO_XLSB_WRITER_H_
#define FORMULON_IO_XLSB_WRITER_H_

#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

struct XlsbWriteResult {
  std::vector<std::uint8_t> bytes;
  std::uint32_t downgraded_formula_count = 0;
  std::uint32_t deferred_feature_count = 0;
};

/// Serialises a workbook and reports formulas emitted as cached literals
/// because their AST could not be lowered to XLSB Ptg tokens, plus modelled
/// features the current XLSB writer could not represent.
Expected<XlsbWriteResult, Error> write_xlsb_with_result(const Workbook& workbook);

/// Serialises `workbook` into an in-memory `.xlsb` byte stream.
///
/// Always-emitted parts:
///   * `[Content_Types].xml`           — Default extensions for `rels`,
///                                       `xml`, `bin`; per-part Override
///                                       entries for the workbook,
///                                       worksheets, and SST when
///                                       present.
///   * `_rels/.rels`                   — package-level rels pointing at
///                                       `xl/workbook.bin`.
///   * `xl/_rels/workbook.bin.rels`    — workbook-level rels (worksheets
///                                       + optional SST).
///   * `xl/workbook.bin`               — `BrtBeginBook | BrtBeginBundleShs
///                                       | BrtBundleSh* | BrtEndBundleShs
///                                       | BrtEndBook`.
///   * `xl/worksheets/sheet<N>.bin`    — one per sheet, 1-based.
///
/// Conditionally emitted parts:
///   * `xl/sharedStrings.bin`          — only when at least one cell
///                                       interns a text value.
///   * Passthrough parts                — every entry in
///                                       `workbook.passthrough_parts()`;
///                                       Default-typed entries reuse the
///                                       matching registrations from
///                                       `workbook.default_content_types()`;
///                                       entries that do not collide with a
///                                       generated path. Collisions are
///                                       resolved in favour of the
///                                       generated copy with a
///                                       `xlsb.writer.passthrough_collision`
///                                       warning.
///
/// Returns `FormulonErrorCode::kIoWriteFailed` on any miniz failure;
/// the error context identifies the offending part. Returns
/// `kInvalidArgument` when `workbook.sheet_count() == 0`.
Expected<std::vector<std::uint8_t>, Error> write_xlsb(const Workbook& workbook);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_WRITER_H_
