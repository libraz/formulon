//
// Internal header for the OOXML cell/row/sheetData builder. Lives inside
// src/io/; not part of the public API. The unit-test friend is the only
// reason these helpers are exposed at all -- AppendSheetDataXml is the
// single function the zip-orchestration TU (ooxml_writer.cpp) calls.
//
// Cells are emitted with first-class dynamic-array spill awareness: spill
// anchors carry t="array" on <f>; phantom cells (covered by another
// anchor's region) are suppressed entirely.

#ifndef FORMULON_IO_OOXML_WRITER_CELL_H_
#define FORMULON_IO_OOXML_WRITER_CELL_H_

#include <cstdint>
#include <string>

#include "io/rich_value_errors.h"
#include "io/xlsb/ptg_writer.h"

namespace formulon {
class Sheet;
struct Cell;
namespace parser {
struct ExternalBookIndexer;
}  // namespace parser
namespace io {
class SharedStrings;

/// True when a save writes a `<c>` element for `cell`.
///
/// A cell is omitted only when it holds nothing at all: no formula, a
/// blank value, and no explicit style assignment. A blank cell carrying
/// a style index (including explicit `s="0"`) still ships, as a bare
/// `<c r="..." s="N"/>`, because Excel preserves the formatting of an
/// empty cell.
///
/// Shared so `<dimension>` is derived from the same test that decides
/// what `<sheetData>` contains. The two answering differently is what
/// makes a package's declared used range contradict its own cells.
bool CellIsEmitted(const Cell& cell);

/// Appends the <sheetData>...</sheetData> markup for a single sheet to `out`.
/// The caller wraps it in <worksheet>. Pure function: no I/O.
///
/// `dynamic_array_cm_index` is the 1-based `<cellMetadata>/<bk>` index a
/// spill anchor's `<c cm="N">` names; 0 (the default) omits `cm=`
/// entirely, which is correct whenever the saved package carries no
/// XLDAPR-typed `xl/metadata.xml` for it to reference. `name_shapes`
/// feeds a legacy formula's `xlsb::legacy_intersections`. `indexer` numbers
/// the external books formulas name; null spells them by name.
/// `rich_error_vm` gives the `vm=` of each cell saved with a rich error
/// (null when the save writes none); it follows `cm=`.
void AppendSheetDataXml(std::string& out, const Sheet& sheet, const SharedStrings* shared_strings = nullptr,
                        std::uint32_t dynamic_array_cm_index = 0U,
                        const xlsb::NameShapes& name_shapes = xlsb::NameShapes(),
                        const parser::ExternalBookIndexer* indexer = nullptr,
                        const RichErrorVmMap* rich_error_vm = nullptr);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WRITER_CELL_H_
