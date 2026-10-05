//
// Sheet geometry: column widths, row heights and cell rectangles in points.
//
// Excel stores a column width in character units relative to the Normal
// font's maximum digit width (MDW). This module owns the Normal-font
// resolution, the measured MDW calibration, and the character-unit to point
// conversion, in two modes:
//
//   * `kDisplay` — the Windows screen figure (`Range.Width` at 96 DPI;
//     Calibri 11 resolves to 5.25 pt per character plus 3.75 pt padding).
//   * `kPrint` — the DPI-stable figure the pagination break math uses
//     (Calibri 11: 39/7 pt per character plus 27/7 pt padding).
//
// Only Windows calibration exists; Mac Excel follows a different regime and
// is deliberately not modelled. A Normal font outside the calibrated set
// falls back to the Calibri 11 figure and reports `calibrated == false`.
//
// Cell rectangles have their origin at the top-left of cell A1; hidden rows
// and columns contribute zero extent.

#ifndef FORMULON_PRINT_SHEET_GEOMETRY_H_
#define FORMULON_PRINT_SHEET_GEOMETRY_H_

#include <cstdint>
#include <string>
#include <vector>

namespace formulon {

class Sheet;
struct FontRecord;
struct StylesTable;

namespace print {

/// Which Windows column-width figure a conversion uses.
enum class GeometryMode : std::int32_t {
  kDisplay = 0,  ///< Screen figure (96 DPI `Range.Width`).
  kPrint = 1,    ///< Print-layout figure used by pagination.
};

/// The character-unit to point conversion for one Normal font and mode.
struct ColumnWidthModel {
  double points_per_char = 0.0;  ///< Points per character unit (the MDW).
  double padding_pt = 0.0;       ///< Flat per-visible-column padding, in points.
  std::string normal_font_name;  ///< Normal-style font family the model was resolved for.
  double normal_font_size = 0.0;
  /// False when `(normal_font_name, normal_font_size)` is not a measured
  /// calibration point and the Calibri 11 figure was used instead.
  bool calibrated = true;
  const char* platform = "win";  ///< The only calibrated platform.
};

/// A rectangle in points, origin at the top-left of the sheet.
struct RectPt {
  double x = 0.0;
  double y = 0.0;
  double width = 0.0;
  double height = 0.0;
};

/// Resolves the workbook's Normal-style font: the `"Normal"` cell style's
/// font, else `fonts[0]`. Returns `nullptr` only when `styles.fonts` is empty.
const FontRecord* resolve_normal_font(const StylesTable& styles);

/// Builds the conversion model for `styles`' Normal font in `mode`. An
/// empty font table yields the Calibri 11 model with `calibrated == false`.
ColumnWidthModel resolve_column_width_model(const StylesTable& styles, GeometryMode mode);

/// Converts a width in character units to points. Zero characters (a hidden
/// or zero-width column) is exactly zero points: no padding applies.
double column_chars_to_points(double chars, const ColumnWidthModel& model);

/// Inverse of `column_chars_to_points`; zero or negative points map to zero.
double column_points_to_chars(double points, const ColumnWidthModel& model);

/// The sheet's default column width in character units
/// (`defaultColWidth`, else Excel's 8.43).
double default_column_width_chars(const Sheet& sheet);

/// The sheet's default row height in points (`defaultRowHeight`, else the
/// measured 102/7 pt).
double default_row_height_pt(const Sheet& sheet);

/// Effective width of column `col` in character units; zero when hidden.
double effective_column_width_chars(const Sheet& sheet, std::uint32_t col);

/// Effective height of row `row` in points; zero when hidden.
double effective_row_height_pt(const Sheet& sheet, std::uint32_t row);

/// Effective widths of columns `first..last` (inclusive) in character units.
/// Equivalent to calling `effective_column_width_chars` per column, in one
/// pass over the layout.
std::vector<double> column_widths_chars(const Sheet& sheet, std::uint32_t first, std::uint32_t last);

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
