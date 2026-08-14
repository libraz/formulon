//
// MS-XLSB per-sheet record-stream emitter. Wraps the cells of a
// `Sheet` in the standard `BrtBeginSheet | BrtBeginSheetData |
// BrtRowHdr* | (BrtCell* | BrtFmla*)* | BrtEndSheetData | BrtEndSheet`
// sequence and returns the bytes ready to be packaged as
// `xl/worksheets/sheet<N>.bin` by the top-level writer.
//
// Cells are emitted in `(row, col)` ascending order, grouped by row
// so each `BrtRowHdr` is followed by its row's cells before the next
// `BrtRowHdr`. Column/row layout, frozen panes and merged-cell
// rectangles are emitted from the model.
//
// Conditional-format rules, data validation, auto-filter, print setup and
// page breaks are not modelled per-record. Hyperlinks are model-owned and
// emitted as BrtHLink records. The remaining unsupported records from a
// sheet that came from an `.xlsb` survive as `Sheet::xlsb_tail()`, whose
// framed bytes are appended around the merged-cell and hyperlink blocks; a
// sheet from any other source carries no retained tail and the writer reports
// unsupported features through
// `XlsbWriteResult::diagnostics.deferred_feature_count`.
//
// Design references:
//   * [MS-XLSB] §2.4.x (BrtBeginSheet / BrtRowHdr / cell records)

#ifndef FORMULON_IO_XLSB_SHEET_WRITER_H_
#define FORMULON_IO_XLSB_SHEET_WRITER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/sst_writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Returns one relationship id per model hyperlink. Empty entries represent
/// internal hyperlinks (which carry their target in `location`). Existing
/// external ids are reused when they do not collide with unrelated retained
/// relationships; fresh ids avoid every retained id.
std::vector<std::string> hyperlink_relationship_ids(const Sheet& sheet);

/// Serialises `sheet` as the body of an `xl/worksheets/sheet<N>.bin`
/// part. `sst` is updated as text cells are interned (the same builder
/// is shared across every sheet of the workbook so all text cells
/// flow into one shared-string table). `sheet_names` is the ordered
/// workbook sheet-name list used to resolve a qualified reference's
/// `ixti` when encoding formula Ptg streams.
///
/// Formulas that cannot be lowered are emitted as cached literals and counted
/// through `downgraded_formula_count`.
///
/// `dynamic_array_ifmd` is the 1-based cell-metadata index a spill anchor's
/// `BrtCellMeta` record must carry to name the dynamic-array entry of the
/// `xl/metadata.bin` part the package ships. Zero means the shipping part has
/// no identifiable dynamic-array entry — or that no part ships at all — and
/// every anchor is then written as a plain `BrtArrFmla` with no `BrtCellMeta`
/// record, because a dangling index makes Excel repair the file.
Expected<std::vector<std::uint8_t>, Error> emit_sheet(const Sheet& sheet, SstBuilder& sst,
                                                      const std::vector<std::string>& sheet_names,
                                                      const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                                      std::uint32_t* downgraded_formula_count = nullptr,
                                                      std::uint32_t dynamic_array_ifmd = 0);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_WRITER_H_
