//
// Sheet geometry: column widths, row heights and cell rectangles in points.
//
// Walks a sheet's column and row overrides to place cells. Column widths are
// the stored character unit and go through the width model in
// `column_width.h`; row heights are already points.
//
// Cell rectangles have their origin at the top-left of cell A1; hidden rows
// and columns contribute zero extent.

#ifndef FORMULON_PRINT_SHEET_GEOMETRY_H_
#define FORMULON_PRINT_SHEET_GEOMETRY_H_

#include <cstdint>
#include <string>
#include <vector>

#include "column_width.h"

namespace formulon {

class Sheet;

namespace print {

/// A rectangle in points, origin at the top-left of the sheet.
struct RectPt {
  double x = 0.0;
  double y = 0.0;
  double width = 0.0;
  double height = 0.0;
};

/// The sheet's default row height in points (`defaultRowHeight`, else the
/// 102/7 pt).
double default_row_height_pt(const Sheet& sheet);

/// Effective stored width of column `col`; zero when hidden.
double effective_column_width_chars(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model);

/// Effective height of row `row` in points; zero when hidden.
double effective_row_height_pt(const Sheet& sheet, std::uint32_t row);

/// Effective stored widths of columns `first..last` (inclusive). Equivalent
/// to calling `effective_column_width_chars` per column, in one pass over the
/// layout.
std::vector<double> column_widths_chars(const Sheet& sheet, std::uint32_t first, std::uint32_t last,
                                        const ColumnWidthModel& model);

/// Effective heights of rows `first..last` (inclusive) in points.
std::vector<double> row_heights_pt(const Sheet& sheet, std::uint32_t first, std::uint32_t last);

/// Effective width of column `col` in points under `model`.
double effective_column_width_pt(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model);

/// Distance from the sheet's left edge to the left edge of column `col`.
double column_offset_pt(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model);

/// Distance from the sheet's top edge to the top edge of row `row`.
double row_offset_pt(const Sheet& sheet, std::uint32_t row);

/// The rectangle of cell (`row`, `col`), 0-based, in points.
RectPt cell_rect_pt(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const ColumnWidthModel& model);

}  // namespace print
}  // namespace formulon

#endif  // FORMULON_PRINT_SHEET_GEOMETRY_H_
