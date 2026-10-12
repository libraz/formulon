// Style operations over a selection: row-style spans, the attribute-patch
// remap Excel applies to a range / row / column / whole-sheet selection, and
// an opaque snapshot that restores a selection's style state for undo.
//
// A rect spanning every row is a column selection, one spanning every column
// a row selection, both the whole sheet. The remap keeps Excel's rule that
// every cell of the selection ends at `map(its effective xf before)`: axis
// records are rewritten where the selection covers the axis, and a cell is
// materialised only where the surviving axis precedence (row over column)
// would otherwise resolve to something else. Only style state is touched;
// values, formulas, phonetic runs, comments and spill state are not.
//
// @size-budget: 4 KB

#ifndef FORMULON_STYLE_REGION_H_
#define FORMULON_STYLE_REGION_H_

#include <cstddef>
#include <cstdint>
#include <vector>

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

/// One `from -> to` pair of a remap; a remap's pairs are sorted by `from`
/// with no repeats.
struct XfRemap {
  std::uint32_t from = 0;
  std::uint32_t to = 0;
};

/// Sets (`styled`) or removes the column style of `[first, last]`, splitting
/// spans at the boundaries and keeping width, hidden and outline state; spans
/// left with no state are dropped.
void set_column_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, bool styled,
                           std::uint32_t xf_index);

/// Gives every row in `[first, last]` the row style `xf_index`, creating row
/// entries as needed and keeping height, hidden and outline state.
void set_row_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, std::uint32_t xf_index);

/// Removes the row style from every row in `[first, last]`; an entry left
/// with no other state is dropped.
void clear_row_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last);

/// Every xf a cell of `rect` may resolve to, ascending: xf 0, the styles of
/// the rows and columns the rect covers, and the xfs of its owning cells.
std::vector<std::uint32_t> range_style_bases(const Sheet& sheet, const StyleRect& rect);

/// Rewrites the selection so each cell resolves to `map(its xf before)`;
/// an xf not in `map` maps to itself. Fails with `kPreconditionFailed`, and
/// changes nothing, when more than `cell_limit` cells would be written.
Expected<void, Error> remap_range_style(Sheet& sheet, const StyleRect& rect, const std::vector<XfRemap>& map,
                                        std::uint64_t cell_limit);

/// Captures the style state `remap_range_style` and the axis / cell style
/// setters can change inside `rect` as opaque, versioned bytes.
std::vector<std::uint8_t> snapshot_range_style(const Sheet& sheet, const StyleRect& rect);

/// Puts a `snapshot_range_style` capture back. Stored cells in the captured
/// rect that were not captured lose their style. Fails with
/// `kInvalidArgument`, changing nothing, on malformed bytes or an xf at or
/// past `xf_count`.
Expected<void, Error> restore_range_style(Sheet& sheet, const std::uint8_t* bytes, std::size_t len,
                                          std::size_t xf_count);

}  // namespace formulon

#endif  // FORMULON_STYLE_REGION_H_
