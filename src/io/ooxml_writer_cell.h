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
#include <optional>
#include <string>
#include <unordered_set>

#include "cell.h"
#include "io/rich_value_errors.h"
#include "io/xlsb/ptg_writer.h"
#include "row_cells.h"

namespace formulon {
class CellStyleSavePlan;
class Sheet;
struct Cell;
namespace parser {
struct ExternalBookIndexer;
}  // namespace parser
namespace io {
class SharedStrings;

/// True when a save writes a `<c>` element for `cell`, the stored cell at
/// `(row, col)` of `sheet`.
///
/// A cell with a formula, a value, a non-zero style or phonetic state ships.
/// A blank one holding only an explicit default style ships, as a bare
/// `<c r="..."/>`, only where that default overrides a row or column format
/// (`explicit_default_observable`); elsewhere it is indistinguishable from
/// an absent cell and Excel drops it too. Style runs follow the same rule
/// through `row_emit_plan`.
///
/// Shared so `<dimension>` is derived from the same test that decides
/// what `<sheetData>` contains. The two answering differently is what
/// makes a package's declared used range contradict its own cells.
///
/// With `style_plan`, a cell outside a merged range (`merged` false) whose xf
/// the plan collapses into xf 0 is judged as an explicit xf 0.
bool CellIsEmitted(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell,
                   const CellStyleSavePlan* style_plan = nullptr, bool merged = false);

/// An inclusive cell rectangle.
struct SheetDimension {
  std::uint32_t first_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
};

/// The used range a save declares (`<dimension ref>`, BrtWsDim), the one
/// source both writers use: the span of the rows the save writes (rows with
/// a written cell, run or spill phantom, rows written as a row format, rows
/// with a layout override) by the span of the written cell columns together
/// with every column of non-zero style. Width-only columns do not count. An
/// axis with nothing on it is its first index; nullopt when neither has
/// anything. With `style_plan`, an xf the plan collapses into xf 0 counts as
/// xf 0, so a row or column carrying only such a style is not counted.
std::optional<SheetDimension> sheet_dimension(const Sheet& sheet, const CellStyleSavePlan* style_plan = nullptr);

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
/// (null when the save writes none); it follows `cm=`. `style_plan` renumbers
/// every written cell, row and phantom xf (`CellStyleSavePlan::cell_xf_index`);
/// whether a cell or row is written is still decided on the model xf. Null
/// writes model indices. `recalc_marks` (this sheet's
/// `Workbook::saved_recalc_marks`, null for none) marks each formula it holds `ca="1"`.
void AppendSheetDataXml(std::string& out, const Sheet& sheet, const SharedStrings* shared_strings = nullptr,
                        std::uint32_t dynamic_array_cm_index = 0U,
                        const xlsb::NameShapes& name_shapes = xlsb::NameShapes(),
                        const parser::ExternalBookIndexer* indexer = nullptr,
                        const RichErrorVmMap* rich_error_vm = nullptr, const CellStyleSavePlan* style_plan = nullptr,
                        const std::unordered_set<CellAddress, CellAddressHash>* recalc_marks = nullptr);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WRITER_CELL_H_
