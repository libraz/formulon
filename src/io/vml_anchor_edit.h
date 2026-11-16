//
// Row/column edits on the legacy VML shapes a sheet keeps beside its notes:
// form controls, OLE objects and the other `x:ClientData` shapes.
//
// Excel moves each shape's `x:Anchor` (column, offset, row, offset of the
// top-left and bottom-right cells, offsets in pixels) by the rule it applies
// to the matching drawing and worksheet anchors (`move_cell_anchor`), and
// re-derives the shape's `margin-left`, `margin-top`, `width` and `height` in
// its `style` from the moved anchor. In VML the flags read inverted: an
// `x:MoveWithCells` element marks a shape that does not move with cells, an
// `x:SizeWithCells` element one that does not size with them.

#ifndef FORMULON_IO_VML_ANCHOR_EDIT_H_
#define FORMULON_IO_VML_ANCHOR_EDIT_H_

#include <cstdint>
#include <functional>
#include <string>

#include "sheet_metadata.h"

namespace formulon::io {

/// Distance in points from the sheet's top (row axis) or left edge to the
/// start of `line`, on the sheet as it stands after the edit.
using LineOffsetPt = std::function<double(bool row_axis, std::uint32_t line)>;

/// Moves the anchor of every non-note shape in `shapes` (the retained VML
/// shape text, see `vml_non_note_shapes`) for a whole row/column `edit`, and
/// re-derives the moved shapes' `style` geometry through `offset_pt`. Returns
/// whether any shape changed.
bool remap_vml_anchors(std::string& shapes, const StructuralEdit& edit, const LineOffsetPt& offset_pt);

}  // namespace formulon::io

#endif  // FORMULON_IO_VML_ANCHOR_EDIT_H_
