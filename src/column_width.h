//
// Column-width model: the stored OOXML column width and the points and pixels
// it lays out at.
//
// Excel stores a column width in character units relative to the Normal
// font's maximum digit width (MDW), padding included (OOXML `<col width>` and
// `defaultColWidth`; Calibri 11 stores the 8.43-character column as
// 9.140625). Every width this module takes or returns is that stored unit,
// except `column_display_chars`, which gives the figure Excel shows. It owns
// the Normal-font resolution, the MDW calibration, and the conversions, in
// two modes:
//
//   * `kDisplay` — the screen figure of the profile's host: on Windows
//     `Range.Width` at 96 DPI, the stored width snapping to whole pixels of
//     the MDW (ECMA-376 §18.3.1.13); on Mac the whole-point layout.
//   * `kPrint` — the DPI-stable Windows figure the pagination break math uses
//     on every host (Calibri 11: 39/7 pt per displayed character plus 27/7 pt
//     padding).
//
// Dimensions come from `font_metrics.h` per Excel host. Windows converts
// through 96-DPI pixels; Mac lays columns out in whole points, one pixel being
// one point. A Normal font outside the host's table falls back to that host's
// Calibri 11 figure and reports `calibrated == false`.

#ifndef FORMULON_COLUMN_WIDTH_H_
#define FORMULON_COLUMN_WIDTH_H_

#include <cstdint>
#include <string>

#include "excel_profile.h"

namespace formulon {

class Sheet;
class Workbook;
struct FontRecord;
struct StylesTable;
struct Theme;

/// Which column-width figure a conversion uses.
enum class GeometryMode : std::int32_t {
  kDisplay = 0,  ///< Screen figure of the profile's host.
  kPrint = 1,    ///< Windows print-layout figure used by pagination, on every host.
};

/// The character-unit to point conversion for one Normal font and mode.
struct ColumnWidthModel {
  double points_per_char = 0.0;  ///< Points per displayed character (the MDW).
  double padding_pt = 0.0;       ///< Flat per-visible-column padding, in points.
  int max_digit_width_px = 7;    ///< Screen MDW in pixels (Mac: points, a pixel being a point).
  int padding_px = 5;            ///< Screen column padding in pixels (Mac: points).
  /// Font family the dimensions were looked up for: the stored name, or on
  /// Mac the theme's face when the Normal font follows the theme.
  std::string normal_font_name;
  double normal_font_size = 0.0;
  /// False when `(normal_font_name, normal_font_size)` is not a tabulated
  /// calibration point of `host` and its Calibri 11 figure was used instead.
  bool calibrated = true;
  ExcelHost host = ExcelHost::kWin365;  ///< Host whose table the model uses; `kPrint` is always Windows.
  const char* platform = "win";         ///< `"win"` or `"mac"`, spelling of `host`.
};

/// Resolves the workbook's Normal-style font: the `"Normal"` cell style's
/// font, else `fonts[0]`. Returns `nullptr` only when `styles.fonts` is empty.
const FontRecord* resolve_normal_font(const StylesTable& styles);

/// Builds the conversion model for `styles`' Normal font in `mode` under
/// `profile`. An empty font table yields the host's Calibri 11 model with
/// `calibrated == false`. `theme` supplies the face of a Mac Normal font that
/// follows the theme.
ColumnWidthModel resolve_column_width_model(const StylesTable& styles, const Theme& theme, ExcelProfile profile,
                                            GeometryMode mode);

/// As above for a workbook's own styles and theme; the theme part is parsed
/// only when a Mac Normal font follows it.
ColumnWidthModel resolve_column_width_model(const Workbook& wb, ExcelProfile profile, GeometryMode mode);

/// Converts a stored width to points. Zero (a hidden or zero-width column) is
/// exactly zero points.
double column_chars_to_points(double chars, const ColumnWidthModel& model);

/// Inverse of `column_chars_to_points`: the stored width Excel writes for a
/// column of `points` (whole screen pixels, whole points on Mac); zero or
/// negative maps to zero.
double column_points_to_chars(double points, const ColumnWidthModel& model);

/// The width Excel displays for a stored width (`Range.ColumnWidth`, the
/// Column Width box): the whole screen pixels it snaps to, less the padding,
/// in MDW units to two decimals. Columns narrower than one character scale by
/// MDW plus padding instead (Calibri 11: 9.140625 shows as 8.43).
double column_display_chars(double chars, const ColumnWidthModel& model);

/// The sheet's default stored column width: `defaultColWidth`, else the
/// width Excel derives from `baseColWidth` and the Normal font (Windows,
/// Calibri 11, base 8: 9.140625, shown as 8.43). On Mac a sheet that stores
/// no default row height (no `<sheetFormatPr>`) derives it from base 10.
double default_column_width_chars(const Sheet& sheet, const ColumnWidthModel& model);

}  // namespace formulon

#endif  // FORMULON_COLUMN_WIDTH_H_
