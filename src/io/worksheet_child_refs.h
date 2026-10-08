//
// Coordinate rewriting inside the retained `<worksheet>` children for row /
// column structural edits.
//
// A child the model does not interpret keeps its source XML, but several of
// them name cells and Excel moves those with the edit, each kind by its own
// rule (measured on Mac Excel 365; tests/fixtures/excel/retained_children_edit/):
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
// `dataConsolidate`, `customSheetViews` and the `<scenarios sqref>` keep
// their text, as they do in Excel.

#ifndef FORMULON_IO_WORKSHEET_CHILD_REFS_H_
#define FORMULON_IO_WORKSHEET_CHILD_REFS_H_

#include "sheet.h"
#include "sheet_passthrough.h"

namespace formulon::io {

/// Rewrites the cell coordinates of every child in `children` for `edit` by
/// the rules above. A child the edit leaves without content is erased, as
/// Excel drops it; a child that does not parse is left alone.
void remap_worksheet_children(WorksheetRawExtensions& children, const StructuralEdit& edit);

}  // namespace formulon::io

#endif  // FORMULON_IO_WORKSHEET_CHILD_REFS_H_
