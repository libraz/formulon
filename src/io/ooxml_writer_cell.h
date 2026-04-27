// Copyright 2026 libraz. Licensed under the MIT License.
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
namespace io {

/// Returns the <sheetData>...</sheetData> markup for a single sheet. The
/// caller wraps it in <worksheet>. Pure function: no I/O, no allocation
/// outside the returned string.
std::string BuildSheetDataXml(const Sheet& sheet);

/// Encodes a 0-based (row, col) into the Excel A1 address (1-based, e.g.
/// "A1", "AA1", "XFD1048576"). Exposed for unit testing; not consumed
/// outside ooxml_writer_cell.cpp.
std::string EncodeA1(std::uint32_t row, std::uint32_t col);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WRITER_CELL_H_
