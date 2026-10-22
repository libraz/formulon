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
//   * `kDisplay` — the Windows screen figure (`Range.Width` at 96 DPI): the
//     stored width snaps to whole pixels of the MDW (ECMA-376 §18.3.1.13).
//   * `kPrint` — the DPI-stable figure the pagination break math uses
//     (Calibri 11: 39/7 pt per displayed character plus 27/7 pt padding).
//
// Only Windows calibration exists; Mac Excel follows a different regime and
// is deliberately not modelled. A Normal font outside the calibrated set
// falls back to the Calibri 11 figure and reports `calibrated == false`.

#ifndef FORMULON_COLUMN_WIDTH_H_
#define FORMULON_COLUMN_WIDTH_H_

#include <cstdint>
#include <string>

namespace formulon {

class Sheet;
struct FontRecord;
struct StylesTable;

/// Which Windows column-width figure a conversion uses.
enum class GeometryMode : std::int32_t {
  kDisplay = 0,  ///< Screen figure (96 DPI `Range.Width`).
  kPrint = 1,    ///< Print-layout figure used by pagination.
};

/// The character-unit to point conversion for one Normal font and mode.
struct ColumnWidthModel {
  double points_per_char = 0.0;  ///< Points per displayed character (the MDW).
  double padding_pt = 0.0;       ///< Flat per-visible-column padding, in points.
  int max_digit_width_px = 7;    ///< Windows screen MDW, 96-DPI pixels.
  int padding_px = 5;            ///< Windows screen column padding, 96-DPI pixels.
  std::string normal_font_name;  ///< Normal-style font family the model was resolved for.
  double normal_font_size = 0.0;
  /// False when `(normal_font_name, normal_font_size)` is not a tabulated
  /// calibration point and the Calibri 11 figure was used instead.
  bool calibrated = true;
  const char* platform = "win";  ///< The only calibrated platform.
};

/// Resolves the workbook's Normal-style font: the `"Normal"` cell style's
/// font, else `fonts[0]`. Returns `nullptr` only when `styles.fonts` is empty.
const FontRecord* resolve_normal_font(const StylesTable& styles);

/// Builds the conversion model for `styles`' Normal font in `mode`. An
/// empty font table yields the Calibri 11 model with `calibrated == false`.
ColumnWidthModel resolve_column_width_model(const StylesTable& styles, GeometryMode mode);

/// Converts a stored width to points. Zero (a hidden or zero-width column) is
/// exactly zero points.
double column_chars_to_points(double chars, const ColumnWidthModel& model);

/// Inverse of `column_chars_to_points`: the stored width Excel writes for a
/// column of `points` (whole screen pixels); zero or negative maps to zero.
double column_points_to_chars(double points, const ColumnWidthModel& model);

/// The width Excel displays for a stored width (`Range.ColumnWidth`, the
/// Column Width box): the whole screen pixels it snaps to, less the padding,
/// in MDW units to two decimals. Columns narrower than one character scale by
/// MDW plus padding instead (Calibri 11: 9.140625 shows as 8.43).
double column_display_chars(double chars, const ColumnWidthModel& model);

/// The sheet's default stored column width: `defaultColWidth`, else the
/// width Excel derives from `baseColWidth` and the Normal font (Calibri 11,
/// base 8: 9.140625, shown as 8.43).
double default_column_width_chars(const Sheet& sheet, const ColumnWidthModel& model);

}  // namespace formulon

#endif  // FORMULON_COLUMN_WIDTH_H_
