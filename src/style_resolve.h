//
// Effective cell style: the formatting a cell shows after row, column and cell XF precedence.
//
// The style is read from the selected `cellXfs` record directly (font, fill,
// border and number-format indices as stored); named-style inheritance is not
// recomputed. Conditional formatting is not applied.

#ifndef FORMULON_STYLE_RESOLVE_H_
#define FORMULON_STYLE_RESOLVE_H_

#include <array>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "color_resolve.h"

namespace formulon {

class CellStyleSavePlan;

class Workbook;
class Sheet;
struct SheetLayout;
struct Cell;
class RowCells;

/// Which level supplied the XF.
enum class StyleSource : std::uint8_t {
  kCell = 0,
  kRow = 1,
  kColumn = 2,
  kDefault = 3,
};

/// The lightweight result shared by style inspection and display rendering.
/// `xf_index` is the selected raw cellXfs index and is not range-checked
/// here; a consumer that reads the styles table (e.g. `effective_style`)
/// clamps an index beyond it to xf 0 while preserving `source`.
struct EffectiveXf {
  std::uint32_t xf_index = 0;
  StyleSource source = StyleSource::kDefault;
};

struct EffectiveStyle {
  std::uint32_t xf_index = 0;
  StyleSource source = StyleSource::kDefault;
  std::uint32_t font_index = 0;
  std::uint32_t fill_index = 0;
  std::uint32_t border_index = 0;
  ResolvedColor font_color;
  std::uint8_t fill_pattern = 0;
  ResolvedColor fill_foreground;
  ResolvedColor fill_background;
  /// Left, right, top, bottom, diagonal.
  std::array<ResolvedColor, 5> border_colors;
  /// Code the cell renders with; a built-in id takes the workbook locale's form.
  std::string num_fmt_code;
  bool locked = true;
  bool hidden = false;
};

/// Selects an XF using cell, row, column, default precedence. A coordinate
/// with its own format (`Sheet::cell_style_at`: any stored cell, or a blank
/// style run) supplies it; any other coordinate inherits. Row style is
/// effective only when the row carries `customFormat`. The returned source
/// uses the `StyleSource` values, and an absent override returns xf 0.
EffectiveXf select_effective_xf(const Sheet& sheet, std::uint32_t row, std::uint32_t col);

/// The row-then-column part of `select_effective_xf`: the XF an empty
/// coordinate shows, read from the layout alone.
EffectiveXf select_axis_xf(const SheetLayout& layout, std::uint32_t row, std::uint32_t col);

/// One piece of a row's axis format: columns `[first, end)` show `xf`.
struct AxisSegment {
  std::uint32_t first = 0;
  std::uint64_t end = 0;
  std::uint32_t xf = 0;
};

/// The axis formats over columns `[first, end)` of `row`: the row style as one
/// piece when the row carries one, else the styled column spans clipped to the
/// band, ascending. Columns no piece covers show xf 0. `layout.columns` must be
/// sorted by `first`.
std::vector<AxisSegment> axis_segments(const SheetLayout& layout, std::uint32_t row, std::uint32_t first,
                                       std::uint64_t end);

/// The columns of `[first, end)` in `row` where an explicit default (xf 0)
/// differs from inheriting: the `axis_segments` pieces whose xf is non-zero.
/// A row style shadows the columns even when it is xf 0.
std::vector<AxisSegment> explicit_default_observable(const SheetLayout& layout, std::uint32_t row, std::uint32_t first,
                                                     std::uint64_t end);

/// Columns `[first, end)` of one row that a save writes as blank cells of `xf`.
struct EmittedRun {
  std::uint32_t first = 0;
  std::uint64_t end = 0;
  std::uint32_t xf = 0;
};

/// How a save writes one stored row's style runs.
struct RowEmitPlan {
  /// The row is written as a row format of xf 0 (customFormat, no cells)
  /// instead of its runs.
  bool as_row_format = false;
  /// The run pieces written as blank cells, ascending and disjoint.
  std::vector<EmittedRun> runs;
};

/// The emit plan of `cells`, the stored row `row` of `sheet`, judged on the
/// current state; `row_has_phantoms` says whether a spill phantom sits in the
/// row. The one source both writers and the dimension use.
///
/// A row with no row format whose content is only xf-0 runs, each under a
/// column of non-zero style, together covering every such column of the sheet
/// and at least three of them, becomes a row format. Otherwise each run piece
/// is written unless its xf equals the xf it would inherit (the row format,
/// else the column format, else 0), split per `axis_segments` piece.
///
/// With `style_plan`, a run piece outside `merged_spans` (the row's merged
/// column spans, `{first, last}`) whose xf the plan collapses into xf 0 is
/// judged as an xf-0 run.
RowEmitPlan row_emit_plan(const Sheet& sheet, std::uint32_t row, const RowCells& cells, bool row_has_phantoms,
                          const CellStyleSavePlan* style_plan = nullptr,
                          const std::vector<std::pair<std::uint32_t, std::uint32_t>>* merged_spans = nullptr);

/// Whether `cell` supplies its own XF rather than inheriting one: an explicit
/// assignment, a non-zero index, a formula or a non-blank value.
bool cell_owns_style(const Cell& cell) noexcept;

/// XF precedence: an owning cell, then its row (when it carries a style),
/// then its column, then xf 0. An xf index beyond `cellXfs` falls back to 0.
EffectiveStyle effective_style(const Workbook& wb, const Sheet& sheet, std::uint32_t row, std::uint32_t col);

}  // namespace formulon

#endif  // FORMULON_STYLE_RESOLVE_H_
