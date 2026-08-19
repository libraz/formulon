//
// Print_Area / Print_Titles resolution.
//
// Excel stores a sheet's print area and repeat-rows/columns as the
// built-in sheet-scoped defined names `_xlnm.Print_Area` and
// `_xlnm.Print_Titles`. This module locates those names on a workbook,
// parses their A1-style range formulas, and exposes the result as plain
// 0-based inclusive rectangles for the pagination engine to consume.

#ifndef FORMULON_PRINT_PRINT_AREA_H_
#define FORMULON_PRINT_PRINT_AREA_H_

#include <cstdint>
#include <optional>
#include <utility>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace print {

/// A rectangular cell range, 0-based and inclusive on both corners.
///
/// `first_row <= last_row` and `first_col <= last_col` hold for every
/// range returned by this module; the parser normalises swapped corners.
struct CellRange {
  std::uint32_t first_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
};

/// The repeat-rows / repeat-columns encoded by `_xlnm.Print_Titles`.
///
/// `repeat_rows` is a 0-based inclusive `[first_row, last_row]` span;
/// `repeat_cols` is a 0-based inclusive `[first_col, last_col]` span.
/// Either may be absent when the defined name only encodes one axis.
struct PrintTitles {
  std::optional<std::pair<std::uint32_t, std::uint32_t>> repeat_rows;
  std::optional<std::pair<std::uint32_t, std::uint32_t>> repeat_cols;
};

/// Parses one print-area token into a normalised rectangle.
///
/// `token` is a single area with no sheet qualifier: a full `A1:H80`
/// range, a degenerate single cell `A1`, or a whole-column (`A:D`) /
/// whole-row (`1:50`) span. `$` anchors are tolerated. Endpoints are
/// normalised (swapped corners are ordered) and clamped to Excel's grid, so
/// an over-large reference cannot drive a runaway track vector downstream.
///
/// Exposed so the authoring API validates a caller-supplied area against the
/// exact grammar the resolver will apply to it later; a token this accepts
/// is one `resolve_print_area` can read back.
///
/// Returns false when the token is not one of those shapes.
bool parse_area_token(std::string_view token, CellRange* out_range);

/// Resolves the print area for sheet `sheet_index` (0-based).
///
/// Scans the workbook's defined names for `_xlnm.Print_Area` scoped to
/// the sheet and parses its (possibly multi-area, comma-separated)
/// formula into one rectangle per area. Sheet qualifiers (`Sheet1!`) and
/// `$` anchors are stripped.
///
/// Returns an empty vector when the defined name is absent (an absent
/// print area is not an error). Returns `kPrintInvalidArea` when the
/// formula is present but malformed.
Expected<std::vector<CellRange>, Error> resolve_print_area(const Workbook& wb, std::uint32_t sheet_index);

/// Resolves the repeat-rows / repeat-columns for sheet `sheet_index`.
///
/// Scans for `_xlnm.Print_Titles` scoped to the sheet. The formula
/// encodes a whole-row span (`Sheet1!$1:$1`), a whole-column span
/// (`Sheet1!$A:$A`), or both, comma-separated. Returns an empty
/// `PrintTitles` when the defined name is absent; `kPrintInvalidArea`
/// when present but malformed.
Expected<PrintTitles, Error> resolve_print_titles(const Workbook& wb, std::uint32_t sheet_index);

}  // namespace print
}  // namespace formulon

#endif  // FORMULON_PRINT_PRINT_AREA_H_
