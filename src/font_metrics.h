//
// Normal-font metrics per Excel host: the maximum digit width (MDW), the
// column padding and the automatic row height of the workbook's Normal font.
//
// Mac values are measured whole points (1 px = 1 pt). Windows values are the
// 96-DPI `Range.Width` calibration; the Windows row height is not measured
// per font, so it is always 102/7 and never marked calibrated.

#ifndef FORMULON_FONT_METRICS_H_
#define FORMULON_FONT_METRICS_H_

#include <cstdint>
#include <string>

#include "excel_profile.h"

namespace formulon {

struct StylesTable;
struct Theme;

/// `FontRecord::scheme` value of a font linked to the theme's minor font.
inline constexpr std::uint8_t kFontSchemeMinor = 2;

/// The dimensions of one Normal font on one host.
struct NormalFontMetrics {
  ExcelHost host = ExcelHost::kWin365;
  std::string family;          ///< Font name the dimensions were looked up for.
  double size_pt = 0.0;        ///< Stored font size; 0 when the style table has no font.
  double mdw_pt = 0.0;         ///< Points per character unit (Mac: whole points; Windows: 96-DPI display figure).
  double padding_pt = 0.0;     ///< Flat per-column padding, in points.
  double row_height_pt = 0.0;  ///< Automatic row height of the font.
  /// False when `(family, size_pt)` is not in the host's width table and the
  /// host's Calibri 11 figures were used.
  bool width_calibrated = false;
  /// True only for a measured automatic row height; false on every Windows row.
  bool row_height_calibrated = false;
};

/// Metrics of the Normal font of `styles` on `profile.host`. On a Mac host a
/// Normal font linked to the theme's minor font (`scheme="minor"`) is looked
/// up under the face `theme` names for the profile's locale script; without
/// one, under the locale's new-workbook font. A name outside the host's table
/// (after the fixed Mac English aliases) yields the host's Calibri 11 figures
/// with both calibration flags false.
NormalFontMetrics normal_font_metrics(const StylesTable& styles, const Theme& theme, ExcelProfile profile);

}  // namespace formulon

#endif  // FORMULON_FONT_METRICS_H_
