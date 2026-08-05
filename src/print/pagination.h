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
// metrics, which are approximated here with Calibri 11 constants.
// Structural correctness — page count, break ordering, and which track
// each break sits before — is the firm goal.

#ifndef FORMULON_PRINT_PAGINATION_H_
#define FORMULON_PRINT_PAGINATION_H_

#include <cstdint>
#include <vector>

#include "print/print_area.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace print {

/// The result of paginating one worksheet.
struct PaginationResult {
  /// The resolved print area (one or more rectangles). When the sheet has
  /// no `_xlnm.Print_Area` this is the sheet's used range; when the sheet
  /// is empty it is left empty.
  std::vector<CellRange> print_area;
  /// 0-based row index each horizontal (page-down) break precedes, in
  /// ascending order. Computed for the bounding box of `print_area`.
  std::vector<std::uint32_t> h_breaks;
  /// 0-based column index each vertical (page-right) break precedes, in
  /// ascending order.
  std::vector<std::uint32_t> v_breaks;
  /// Total physical page count: column-pages multiplied by row-pages.
  std::uint32_t page_count = 0;
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
Expected<PaginationResult, Error> paginate(const Workbook& wb, std::uint32_t sheet_index);

}  // namespace print
}  // namespace formulon

#endif  // FORMULON_PRINT_PAGINATION_H_
