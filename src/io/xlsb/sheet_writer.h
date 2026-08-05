//
// MS-XLSB per-sheet record-stream emitter. Wraps the cells of a
// `Sheet` in the standard `BrtBeginSheet | BrtBeginSheetData |
// BrtRowHdr* | (BrtCell* | BrtFmla*)* | BrtEndSheetData | BrtEndSheet`
// sequence and returns the bytes ready to be packaged as
// `xl/worksheets/sheet<N>.bin` by the top-level writer.
//
// Cells are emitted in `(row, col)` ascending order, grouped by row
// so each `BrtRowHdr` is followed by its row's cells before the next
// `BrtRowHdr`. Column/row layout and merged-cell rectangles are also
// emitted. Conditional-format rules, data validation, hyperlinks,
// page breaks, frozen panes, and other sheet-level metadata remain
// outside this stream emitter's current scope.
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

/// Serialises `sheet` as the body of an `xl/worksheets/sheet<N>.bin`
/// part. `sst` is updated as text cells are interned (the same builder
/// is shared across every sheet of the workbook so all text cells
/// flow into one shared-string table). `sheet_names` is the ordered
/// workbook sheet-name list used to resolve a qualified reference's
/// `ixti` when encoding formula Ptg streams.
///
/// Formulas that cannot be lowered are emitted as cached literals and counted
/// through `downgraded_formula_count`.
Expected<std::vector<std::uint8_t>, Error> emit_sheet(const Sheet& sheet, SstBuilder& sst,
                                                      const std::vector<std::string>& sheet_names,
                                                      const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                                      std::uint32_t* downgraded_formula_count = nullptr,
                                                      bool emit_dynamic_metadata = false);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_WRITER_H_
