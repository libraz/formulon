// Style operations over a selection: row-style spans, the attribute-patch
// remap Excel applies to a range / row / column / whole-sheet selection, and
// an opaque snapshot that restores a selection's style state for undo.
//
// A rect spanning every row is a column selection, one spanning every column
// a row selection, both the whole sheet. The remap keeps Excel's rule that
// every cell of the selection ends at `map(its effective xf before)`: axis
// records are rewritten where the selection covers the axis, and a coordinate
// gets its own format only where the surviving axis precedence (row over
// column) would otherwise resolve to something else. Blank coordinates take
// it as style runs, never as stored cells. Only style state is touched;
// values, formulas, phonetic runs, comments and spill state are not.
//
// @size-budget: 4 KB

#ifndef FORMULON_STYLE_REGION_H_
#define FORMULON_STYLE_REGION_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "row_cells.h"
#include "sheet_layout.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Sheet;

/// Inclusive cell rectangle with `first_* <= last_*`, inside the grid.
struct StyleRect {
  std::uint32_t first_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
};

/// One cell of a sheet.
struct StyleCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
};

/// Sets (`styled`) or removes the column style of `[first, last]`, splitting
/// spans at the boundaries and keeping width, hidden and outline state; spans
/// left with no state are dropped.
void set_column_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, bool styled,
                           std::uint32_t xf_index);

/// Gives every row in `[first, last]` the row style `xf_index`, creating row
/// entries as needed and keeping height, hidden and outline state.
void set_row_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, std::uint32_t xf_index,
                        bool default_hidden);

/// Removes the row style from every row in `[first, last]`; an entry left
/// with no other state is dropped, except an existing visible row under
/// `zeroHeight`, whose presence itself is a visibility override.
void clear_row_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, bool default_hidden);

/// Every xf some cell of `rect` outside `excluded` resolves to, ascending: the
/// xfs of its owning cells and style runs, then the row, column or default (0)
/// style of any cell left over. Excluded cells outside `rect` and repeats are
/// ignored.
std::vector<std::uint32_t> range_style_bases(const Sheet& sheet, const StyleRect& rect,
                                             std::vector<StyleCell> excluded);

/// One coordinate and its own xf.
struct StyleCellXf {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t xf = 0;
};

/// Number of coordinates of `rect` with a non-zero own xf: stored cells and
/// style-run columns. Costs the rows and runs, not the area.
std::size_t count_cell_xf_indices(const Sheet& sheet, const StyleRect& rect);

/// Writes the first `capacity` coordinates `count_cell_xf_indices` counts into
/// `out`, row-major, and returns how many it wrote.
std::size_t copy_cell_xf_indices(const Sheet& sheet, const StyleRect& rect, StyleCellXf* out, std::size_t capacity);

/// Every stored cell of `rect` whose Show Phonetic flag is set, row-major.
std::vector<StyleCell> show_phonetic_cells(const Sheet& sheet, const StyleRect& rect);

/// The one guard on a range style application: a rect that is neither whole
/// columns nor whole rows and covers more than `kMaxStyledRangeCells` cells
/// fails with `kSecurityResourceLimit`. File readers never call it.
Expected<void, Error> check_styled_range_area(const StyleRect& rect);

/// Rewrites the selection so each cell resolves to `map(its xf before)`:
/// `check_styled_range_area`, then the stored owners and the planned writes
/// are taken from the current state, the axis records are rewritten, and each
/// affected row has its existing runs mapped once, its planned pieces written
/// and its owners mapped. An xf-0 blank survives only where the rewritten axis
/// format is non-zero. Nothing changes on failure. The caller holds the
/// workbook's compound mutation lock.
Expected<void, Error> remap_range_style(Sheet& sheet, const StyleRect& rect, const XfMap& map);

/// `remap_range_style` with every xf mapped to `xf_index`.
Expected<void, Error> set_range_xf_index(Sheet& sheet, const StyleRect& rect, std::uint32_t xf_index);

/// Captures the style state `remap_range_style` and the axis / cell style
/// setters can change inside `rect` as opaque, versioned bytes: the axis
/// records, then per row the formatted cells and the style runs, clipped.
std::vector<std::uint8_t> snapshot_range_style(const Sheet& sheet, const StyleRect& rect);

/// Puts a `snapshot_range_style` capture back, of this or the earlier
/// per-cell version. Runs and stored cells in the captured rect that were not
/// captured lose their style; a captured explicit default comes back only
/// where an axis format still makes it observable. Fails with
/// `kInvalidArgument`, changing nothing, on malformed bytes (rows or entries
/// out of order, overlapping or outside the rect, an unknown tag, flag or
/// version) or an xf at or past `xf_count`.
Expected<void, Error> restore_range_style(Sheet& sheet, const std::uint8_t* bytes, std::size_t len,
                                          std::size_t xf_count);

}  // namespace formulon

#endif  // FORMULON_STYLE_REGION_H_
