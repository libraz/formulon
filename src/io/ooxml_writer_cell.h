//
// Internal header for the OOXML cell/row/sheetData builder. Lives inside
// src/io/; not part of the public API. The unit-test friend is the only
// reason these helpers are exposed at all -- BuildSheetDataXml is the
// single function the zip-orchestration TU (ooxml_writer.cpp) calls.
//
// Cells are emitted with first-class dynamic-array spill awareness: spill
// anchors carry t="array" on <f>; phantom cells (covered by another
// anchor's region) are suppressed entirely.

#ifndef FORMULON_IO_OOXML_WRITER_CELL_H_
#define FORMULON_IO_OOXML_WRITER_CELL_H_

#include <cstdint>
#include <string>

namespace formulon {
class Sheet;
struct Cell;
namespace io {
class SharedStrings;

/// True when a save writes a `<c>` element for `cell`.
///
/// A cell is omitted only when it holds nothing at all: no formula, a
/// blank value, and the default style. A blank cell carrying a style
/// index still ships, as a bare `<c r="..." s="N"/>`, because Excel
/// preserves the formatting of an empty cell.
///
/// Shared so `<dimension>` is derived from the same test that decides
/// what `<sheetData>` contains. The two answering differently is what
/// makes a package's declared used range contradict its own cells.
bool CellIsEmitted(const Cell& cell);

/// Returns the <sheetData>...</sheetData> markup for a single sheet. The
/// caller wraps it in <worksheet>. Pure function: no I/O, no allocation
/// outside the returned string.
std::string BuildSheetDataXml(const Sheet& sheet, const SharedStrings* shared_strings = nullptr);

/// Encodes a 0-based (row, col) into the Excel A1 address (1-based, e.g.
/// "A1", "AA1", "XFD1048576"). Exposed for unit testing; not consumed
/// outside ooxml_writer_cell.cpp.
std::string EncodeA1(std::uint32_t row, std::uint32_t col);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WRITER_CELL_H_
