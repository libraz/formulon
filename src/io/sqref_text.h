//
// The A1 text form of an sqref (`A1:B2 D4`) shared by the retained-XML
// coordinate rewriters.

#ifndef FORMULON_IO_SQREF_TEXT_H_
#define FORMULON_IO_SQREF_TEXT_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "sheet.h"

namespace formulon::io {

/// Appends the space-separated A1 ranges of `text` to `out`. False, with
/// `out` partly filled, when a token does not parse or `text` names none.
bool parse_sqref_text(std::string_view text, std::vector<MergeRange>& out);

/// Appends the ranges of a `<selection sqref>`: A1 forms plus whole-row
/// (`3:4`) and whole-column (`A:B`) tokens widened to the full sheet. A token
/// that does not parse, is inverted or leaves the grid is skipped.
void parse_selection_sqref_text(std::string_view text, std::vector<MergeRange>& out);

/// The A1 text of `ranges`, a single cell written without `:`.
std::string format_sqref_text(const std::vector<MergeRange>& ranges);

/// One pane's selection as Excel saves it: the active cell lies in a range.
struct ResolvedSelection {
  std::string_view pane;  ///< `SheetView::pane_key` form.
  std::uint32_t active_row = 0;
  std::uint32_t active_col = 0;
  std::vector<MergeRange> ranges;    ///< Never empty.
  std::uint32_t active_cell_id = 0;  ///< Index of the range holding the active cell.
};

/// Every pane `view` has, in Excel's save order (topLeft, topRight,
/// bottomLeft, bottomRight), with its first stored selection. A pane with
/// none selects one cell, as Excel does: A1 in the top-left pane, the scroll
/// origin (`top_left_cell`) in the top-right and bottom-left panes, and
/// `pane_origin` in the bottom-right one. An active cell outside every range
/// moves to the first range's top-left.
std::vector<ResolvedSelection> resolve_view_selections(const SheetView& view);

/// The row and column of the lower-right pane's top-left cell: the stored
/// `pane_top_left_cell`, else the first unfrozen cell (A1 for a split or a
/// view without panes).
void pane_origin(const SheetView& view, std::uint32_t* row, std::uint32_t* col);

}  // namespace formulon::io

#endif  // FORMULON_IO_SQREF_TEXT_H_
