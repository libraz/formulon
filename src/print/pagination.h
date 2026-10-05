//
// Page-break (pagination) engine.
//
// Given a worksheet, computes where Excel would place automatic page
// breaks: it sizes every column and row of the print area in points,
// fits them into the printable body area derived from the page setup,
// and records the row/column index each break precedes. Manual breaks
// from `<rowBreaks>` / `<colBreaks>` are honoured on top of the automatic
// flow.
//
// Exact 1-bit parity with Excel's pagination is best-effort: the
// character-width to pixel rounding depends on the rendering font's
// metrics. The workbook's Normal-style font/size selects a measured
// calibration when it is one of the sampled Windows combinations, and
// falls back to Calibri-11 constants otherwise. Structural correctness —
// page count, break ordering, and which track each break sits before —
// is the firm goal.

#ifndef FORMULON_PRINT_PAGINATION_H_
#define FORMULON_PRINT_PAGINATION_H_

#include <cstdint>
#include <vector>

#include "print/page_setup.h"
#include "print/print_area.h"
#include "print/sheet_geometry.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace print {

/// The physical paper a sheet prints on, orientation-adjusted.
struct PaperInfo {
  double width_pt = 0.0;   ///< Page width after the orientation swap, in points.
  double height_pt = 0.0;  ///< Page height after the orientation swap, in points.
  bool landscape = false;
  /// False when `paperSize` is not a recognised code and A4 was substituted.
  bool known = true;
};

/// Page margins in points.
struct MarginsPt {
  double left = 0.0;
  double right = 0.0;
  double top = 0.0;
  double bottom = 0.0;
  double header = 0.0;  ///< Page edge to header band.
  double footer = 0.0;  ///< Page edge to footer band.
};

/// The repeat-rows / repeat-columns in effect (0-based, inclusive).
struct PrintTitleSpan {
  bool has_rows = false;
  std::uint32_t first_row = 0;
  std::uint32_t last_row = 0;
  bool has_cols = false;
  std::uint32_t first_col = 0;
  std::uint32_t last_col = 0;
};

/// One physical page: the cell block it carries and where it sits.
///
/// `width_pt` / `height_pt` are the scaled extent of the page's own cell
/// block (repeated titles excluded). `origin_*_pt` is the block's top-left
/// corner on the paper; a page that reprints title rows / columns starts
/// below / right of them.
struct PageLayout {
  std::uint32_t area_index = 0;  ///< Index into `PaginationResult::print_area` (0 for the used-range fallback).
  std::uint32_t first_row = 0;
  std::uint32_t last_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_col = 0;
  double origin_x_pt = 0.0;
  double origin_y_pt = 0.0;
  double width_pt = 0.0;
  double height_pt = 0.0;
};

/// The result of paginating one worksheet.
struct PaginationResult {
  /// The sheet's declared `_xlnm.Print_Area`, as one or more rectangles.
  /// Empty when the sheet declares no print area — this field mirrors
  /// Excel's own `PageSetup.PrintArea`, which is likewise empty for a
  /// sheet that has never had one set, and is *not* backfilled with the
  /// used range. Pagination itself still falls back to the used range in
  /// that case, so `page_count` can be non-zero while this is empty.
  std::vector<CellRange> print_area;
  /// 0-based row index each horizontal (page-down) break precedes, in
  /// ascending order. Computed for the bounding box of `print_area`.
  std::vector<std::uint32_t> h_breaks;
  /// 0-based column index each vertical (page-right) break precedes, in
  /// ascending order.
  std::vector<std::uint32_t> v_breaks;
  /// Total physical page count: column-pages multiplied by row-pages.
  std::uint32_t page_count = 0;
  /// Paper after the orientation swap.
  PaperInfo paper;
  /// Page margins in points.
  MarginsPt margins;
  /// Cell-body rectangle before print-title reservation, in points from the
  /// page's top-left corner: it starts at the left / top margin.
  RectPt printable;
  /// Effective scale as a factor (1.0 is 100%): the `scale` percentage, or
  /// the fit-to-page factor.
  double scale = 1.0;
  PageOrder page_order = PageOrder::kDownThenOver;
  PrintTitleSpan print_titles;
  /// Every physical page, in print order (`page_order`; print areas in
  /// declaration order). `pages.size() == page_count`.
  std::vector<PageLayout> pages;
  /// Parallel to `h_breaks` / `v_breaks`: true when that break is a manual
  /// one rather than computed by the page walk.
  std::vector<bool> h_break_manual;
  std::vector<bool> v_break_manual;
};

/// Paginates sheet `sheet_index` (0-based) of `wb`.
///
/// Resolves the print area, sizes its columns/rows in points, derives the
/// printable body area from the page setup (applying `fitToPage` /
/// `fitToWidth` / `fitToHeight` or the uniform `scale`), then walks the
/// tracks accumulating points until the body limit is reached, recording
/// a break before each overflowing track. Manual breaks always force a
/// break before their track.
///
/// Returns `kInvalidArgument` when `sheet_index` is out of range, or a
/// `kPrintInvalidArea` propagated from print-area resolution. An empty
/// sheet with no print area yields a result with `page_count == 0`.
///
/// Returns `kPrintPageCountOverflow` when the page grid a sheet declares
/// exceeds `kMaxPaginationPages`, which a file listing a manual break
/// before nearly every row and column can reach. `page_count` is either
/// the true total for the whole print area or that error; it is never a
/// truncated count.
Expected<PaginationResult, Error> paginate(const Workbook& wb, std::uint32_t sheet_index);

}  // namespace print
}  // namespace formulon

#endif  // FORMULON_PRINT_PAGINATION_H_
