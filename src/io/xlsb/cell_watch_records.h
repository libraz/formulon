//
// Cell watches as xlsb records: BrtBeginCellWatches, one BrtCellWatch per watched cell (u32 row, u32 col, both
// 0-based) and BrtEndCellWatches. The model keeps them as the retained `<cellWatches>` worksheet child, so both
// formats read and write the one element.

#ifndef FORMULON_IO_XLSB_CELL_WATCH_RECORDS_H_
#define FORMULON_IO_XLSB_CELL_WATCH_RECORDS_H_

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/zip_reader.h"
#include "sheet_passthrough.h"

namespace formulon::io::xlsb {

/// The A1 reference of the cell a BrtCellWatch payload names; nullopt when the payload is not two u32 fields or the
/// cell lies outside the grid.
std::optional<std::string> decode_cell_watch(ByteSpan payload);

/// Replaces the retained `<cellWatches>` child with one `<cellWatch>` per reference, in order; an empty list removes
/// it.
void write_cell_watches(WorksheetRawExtensions& children, const std::vector<std::string>& refs);

/// Inserts the cell-watch block for the retained `<cellWatches>` child into the post-hyperlink tail, after the print
/// records and BrtBigName that precede it; nothing when the sheet has no watch. A watch whose reference does not parse
/// as a cell is left out.
void splice_cell_watches(std::vector<std::uint8_t>& after_hyperlinks, const WorksheetRawExtensions& children);

}  // namespace formulon::io::xlsb

#endif  // FORMULON_IO_XLSB_CELL_WATCH_RECORDS_H_
