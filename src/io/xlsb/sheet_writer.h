// Copyright 2026 libraz. Licensed under the MIT License.
//
// MS-XLSB per-sheet record-stream emitter. Wraps the cells of a
// `Sheet` in the standard `BrtBeginSheet | BrtBeginSheetData |
// BrtRowHdr* | (BrtCell* | BrtFmla*)* | BrtEndSheetData | BrtEndSheet`
// sequence and returns the bytes ready to be packaged as
// `xl/worksheets/sheet<N>.bin` by the top-level writer.
//
// Cells are emitted in `(row, col)` ascending order, grouped by row
// so each `BrtRowHdr` is followed by its row's cells before the next
// `BrtRowHdr`. Defined names, conditional-format rules, page
// breaks, frozen panes, and other sheet-level metadata are out of
// scope here; if any are present on the workbook, the top-level
// writer logs `xlsb.writer.deferred` and skips them.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.6 (XLSB record stream layout)
//   * [MS-XLSB] §2.4.x (BrtBeginSheet / BrtRowHdr / cell records)

#ifndef FORMULON_IO_XLSB_SHEET_WRITER_H_
#define FORMULON_IO_XLSB_SHEET_WRITER_H_

#include <cstdint>
#include <vector>

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
/// flow into one shared-string table).
///
/// Returns no errors today; the `Expected` shape is preserved for
/// forward compatibility with future per-sheet metadata that might
/// fail at emit time (e.g. validation rules with malformed refs).
Expected<std::vector<std::uint8_t>, Error> emit_sheet(const Sheet& sheet, SstBuilder& sst);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_WRITER_H_
