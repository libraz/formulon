//
// Coordinate rewriting inside the retained `<worksheet>` children for row /
// column structural edits.
//
// A child the model does not interpret keeps its source XML, but several of
// them name cells and Excel moves those with the edit, each kind by its own
// rule (tests/fixtures/excel/retained_children_edit/):
//
//   protectedRange sqref, sortState / sortCondition ref, and the sourceRef
//   of a range webPublishItem      -- ordinary ranges (`shift_sqref_ranges`);
//                                     a sort state goes with its last
//                                     condition;
//   ignoredError sqref             -- cut at the edit (`cut_sqref_ranges`);
//                                     emptying one also drops every entry
//                                     after it;
//   scenario inputCells r          -- a cell a delete removes, taking the
//                                     scenario with its last cell; the
//                                     current / shown index then clamps to
//                                     the scenarios left;
//   cellWatch r                    -- a cell a delete leaves where it was.
//
//   controlPr / objectPr anchor    -- the cell anchor of a form control or
//                                     OLE object (`move_cell_anchor`),
//                                     whole row/column edits only.
//
// Under a band edit (Insert / Delete Cells) only a range inside the band moves.
//
// `dataConsolidate`, `customSheetViews` and the `<scenarios sqref>` keep
// their text, as they do in Excel.

#ifndef FORMULON_IO_WORKSHEET_CHILD_REFS_H_
#define FORMULON_IO_WORKSHEET_CHILD_REFS_H_

#include <cstdint>
#include <string>
#include <vector>

#include "sheet.h"
#include "sheet_passthrough.h"

namespace formulon::io {

/// One axis of an anchor marker: a 0-based row or column and the offset into it.
struct AnchorLine {
  std::uint32_t line = 0;
  std::int64_t offset = 0;
};

/// Moves `marker` for a row/column edit along its axis: at or after an insert
/// it moves by the count (held on the sheet), past a deleted span it moves
/// back, and inside one it lands on the deletion start with offset 0.
AnchorLine shift_anchor_line(AnchorLine marker, const StructuralEdit& edit);

/// Moves a cell-anchored object's markers along the edit's axis as Excel does:
/// one that moves and sizes with cells shifts each marker; one that only moves
/// shifts `from` and moves `to` by as many lines, keeping its line span.
void move_cell_anchor(AnchorLine& from, AnchorLine& to, bool size_with_cells, const StructuralEdit& edit);

/// Whether a whole-line delete removes a move-and-size object anchored from
/// `from` to `to`: every line it covers is deleted, a `to` at offset 0 on the
/// line after the span included. Excel removes such an object rather than
/// shrinking it; one that only moves is never removed.
bool cell_anchor_deleted(const AnchorLine& from, const AnchorLine& to, const StructuralEdit& edit);

/// The `shapeId`s of the form controls and OLE objects in `children` whose
/// anchor moves and sizes with cells and that `edit` deletes by
/// `cell_anchor_deleted`.
std::vector<std::uint32_t> deleted_cell_objects(const WorksheetRawExtensions& children, const StructuralEdit& edit);

/// Rewrites the cell coordinates of every child in `children` for `edit` by
/// the rules above. A child the edit leaves without content is erased, as
/// Excel drops it; a child that does not parse is left alone.
void remap_worksheet_children(WorksheetRawExtensions& children, const StructuralEdit& edit);

/// Removes the form controls and OLE objects whose `shapeId` is in
/// `shape_ids` from the retained `<controls>` / `<oleObjects>`, with the
/// `mc:AlternateContent` around each, then any container the removal empties.
/// Returns the relationship ids the removed objects named.
std::vector<std::string> remove_worksheet_objects(WorksheetRawExtensions& children,
                                                  const std::vector<std::uint32_t>& shape_ids);

}  // namespace formulon::io

#endif  // FORMULON_IO_WORKSHEET_CHILD_REFS_H_
