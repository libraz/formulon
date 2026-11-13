#include "column_width.h"

#include <algorithm>
#include <cmath>
#include <string>

#include "font_metrics.h"
#include "new_workbook_defaults.h"
#include "sheet.h"
#include "styles.h"
#include "theme.h"
#include "workbook.h"

namespace formulon {
namespace {

// --- Excel column-width geometry constants. ---
//
// Excel stores a column's width in "character" units relative to the
// default font's maximum digit width (MDW). The character-to-pixel
// conversion below is Excel's documented formula; the constants are
// named so the arithmetic carries no bare literals.
//
// Two distinct quantities share this stored-width input and must not be
// confused:
//
//   * Print-layout width, consumed only by `paginate()`'s own break math
//     (h_breaks / v_breaks / page count). Comparing a Windows Excel 365
//     ja-JP build 16.0.20228/16.0.20326 capture taken at 100% display
//     scaling (96 DPI) against one of the *same build* taken at 175%
//     (168 DPI) shows `Range.Width` itself is DPI-dependent (Calibri 11
//     resolves a 30-character column to 161.25 pt at 96 DPI and 171.0 pt
//     at 168 DPI) -- but the resulting page-break positions are
//     identical across both captures. Print output cannot depend on the
//     authoring screen's DPI, so the break math needs a DPI-stable
//     figure, which the 168-DPI capture's `Range.Width` approximates
//     and the 96-DPI one does not.
//   * Display width, `applied_geometry.column_widths_pt` in the workbook
//     oracle goldens: a diagnostic/reporting figure, not compared by any
//     verifier today (`workbook_oracle_test.cpp` never reads it). This
//     is the 96-DPI `Range.Width` figure, unadjusted.
//
// Only Calibri 11's print-layout width is known directly (from the
// print-track break positions, not `Range.Width`).
// Other fonts/sizes' print-layout widths are not known; they are
// inferred by scaling that font's *display* calibration by the ratio
// between Calibri 11's print and display figures, separately for the
// per-character rate and the padding term (`kPrintMdwRatio` /
// `kPrintPaddingRatio` below) -- an approximation.

/// Calibri 11's print-layout points-per-character-unit and per-column
/// padding: the pre-existing figure that both DPI captures' page breaks
/// agree with, matching the classic `39/7`, `27/7` sevenths that predate
/// this file's DPI investigation. This is also the print-layout fallback
/// for any (Normal font family, size) the Windows width table does not
/// cover.
constexpr double kPrintPointsPerColumnCharCalibri11 = 39.0 / 7.0;
constexpr double kPrintColumnPaddingPtCalibri11 = 27.0 / 7.0;

/// Calibri 11's *display* points-per-character-unit and per-column
/// padding (`Range.Width` at 96 DPI) -- the Windows width table entry for
/// `{"Calibri", 11}`, restated here so the print/display ratios below don't
/// depend on table lookup order.
constexpr double kDisplayPointsPerColumnCharCalibri11 = 5.25;
constexpr double kDisplayColumnPaddingPtCalibri11 = 3.75;

/// Print-layout-to-display ratios for the per-character rate and the
/// padding term, derived from Calibri 11 (the only font with a known
/// print-layout figure) and applied to every other font's
/// display calibration to approximate its print-layout one.
constexpr double kPrintMdwRatio = kPrintPointsPerColumnCharCalibri11 / kDisplayPointsPerColumnCharCalibri11;
constexpr double kPrintPaddingRatio = kPrintColumnPaddingPtCalibri11 / kDisplayColumnPaddingPtCalibri11;

/// Points per 96-DPI screen pixel.
constexpr double kPointsPerPixel = 72.0 / 96.0;

/// A stored width resolves in 1/256 of a character (ECMA-376 §18.3.1.13).
constexpr double kStoredWidthSteps = 256.0;

/// Windows rounds a default column derived from `baseColWidth` up to a
/// multiple of 8 pixels: Calibri 11 base 8 is 61 px, shown as 64 px (8.43).
constexpr int kDefaultColumnPixelStep = 8;
}  // namespace

const FontRecord* resolve_normal_font(const StylesTable& styles) {
  if (styles.fonts.empty()) {
    return nullptr;
  }
  for (const CellStyleRecord& cell_style : styles.cell_styles) {
    // By builtin id, not name: Excel localizes the Normal style's name.
    if (cell_style.builtin_id != 0U) {
      continue;
    }
    if (cell_style.xf_id < styles.cell_style_xfs.size()) {
      const std::uint32_t font_index = styles.cell_style_xfs[cell_style.xf_id].font_index;
      if (font_index < styles.fonts.size()) {
        return &styles.fonts[font_index];
      }
    }
    break;
  }
  return &styles.fonts[0];
}

ColumnWidthModel resolve_column_width_model(const StylesTable& styles, const Theme& theme, ExcelProfile profile,
                                            GeometryMode mode) {
  // Print is the Windows figure on every host.
  const ExcelProfile host_profile{mode == GeometryMode::kPrint ? ExcelHost::kWin365 : profile.host, profile.locale};
  const NormalFontMetrics metrics = normal_font_metrics(styles, theme, host_profile);
  ColumnWidthModel model;
  model.normal_font_name = metrics.family;
  model.normal_font_size = metrics.size_pt;
  model.calibrated = metrics.width_calibrated;
  model.host = metrics.host;
  if (metrics.host == ExcelHost::kMac365) {
    // A pixel is a point on Mac.
    model.points_per_char = metrics.mdw_pt;
    model.padding_pt = metrics.padding_pt;
    model.max_digit_width_px = static_cast<int>(std::lround(metrics.mdw_pt));
    model.padding_px = static_cast<int>(std::lround(metrics.padding_pt));
    model.platform = "mac";
    return model;
  }
  // Print scales the display calibration by Calibri 11's print/display ratios.
  const bool print = mode == GeometryMode::kPrint;
  model.points_per_char = print ? metrics.mdw_pt * kPrintMdwRatio : metrics.mdw_pt;
  model.padding_pt = print ? metrics.padding_pt * kPrintPaddingRatio : metrics.padding_pt;
  model.max_digit_width_px = static_cast<int>(std::lround(metrics.mdw_pt / kPointsPerPixel));
  model.padding_px = static_cast<int>(std::lround(metrics.padding_pt / kPointsPerPixel));
  return model;
}

ColumnWidthModel resolve_column_width_model(const Workbook& wb, ExcelProfile profile, GeometryMode mode) {
  const StylesTable& styles = wb.styles();
  if (profile.host == ExcelHost::kMac365 && mode == GeometryMode::kDisplay) {
    const FontRecord* font = resolve_normal_font(styles);
    if (font != nullptr && font->scheme == kFontSchemeMinor) {
      return resolve_column_width_model(styles, wb.load_theme().theme, profile, mode);
    }
  }
  return resolve_column_width_model(styles, default_theme(), profile, mode);
}

namespace {

/// Whole screen pixels a stored width snaps to (ECMA-376 §18.3.1.13).
double StoredWidthPixels(double chars, const ColumnWidthModel& model) {
  const double mdw = static_cast<double>(model.max_digit_width_px);
  return std::floor((kStoredWidthSteps * chars + std::floor(kStoredWidthSteps / 2.0 / mdw)) / kStoredWidthSteps * mdw);
}

}  // namespace

double column_chars_to_points(double chars, const ColumnWidthModel& model) {
  // A hidden or explicit zero-width column converts to exactly zero so it
  // never advances a page break or shifts the fit-to-page scale (Excel
  // excludes it from pagination extent); the flat padding term models the
  // allowance every *visible* column carries.
  if (chars == 0.0) {
    return 0.0;
  }
  const double mdw = static_cast<double>(model.max_digit_width_px);
  const double pixels = StoredWidthPixels(chars, model);
  // Scale by points per pixel first so whole screen pixels stay exact quarter points.
  const double points_per_pixel = model.points_per_char / mdw;
  return std::max(0.0, (pixels - static_cast<double>(model.padding_px)) * points_per_pixel + model.padding_pt);
}

double column_points_to_chars(double points, const ColumnWidthModel& model) {
  if (points <= 0.0 || model.points_per_char <= 0.0) {
    return 0.0;
  }
  const double mdw = static_cast<double>(model.max_digit_width_px);
  const double displayed_chars = (points - model.padding_pt) / model.points_per_char;
  const double pixels = std::round(displayed_chars * mdw + static_cast<double>(model.padding_px));
  if (pixels <= 0.0) {
    return 0.0;
  }
  return std::floor(pixels / mdw * kStoredWidthSteps) / kStoredWidthSteps;
}

double column_display_chars(double chars, const ColumnWidthModel& model) {
  const double pixels = StoredWidthPixels(chars, model);
  if (pixels <= 0.0) {
    return 0.0;
  }
  const double mdw = static_cast<double>(model.max_digit_width_px);
  const double padding = static_cast<double>(model.padding_px);
  const double displayed = pixels < mdw + padding ? pixels / (mdw + padding) : (pixels - padding) / mdw;
  return std::round(displayed * 100.0) / 100.0;
}

double default_column_width_chars(const Sheet& sheet, const ColumnWidthModel& model) {
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  if (defaults.has_default_col_width) {
    return defaults.default_col_width;
  }
  if (model.host == ExcelHost::kMac365) {
    // Whole points, no pixel step; a sheet without `<sheetFormatPr>` starts at base 10.
    const double base = defaults.has_default_row_height ? defaults.base_col_width : kNewWorkbookMacBaseColWidth;
    return base + std::floor(kStoredWidthSteps * model.padding_pt / model.points_per_char) / kStoredWidthSteps;
  }
  const double mdw = static_cast<double>(model.max_digit_width_px);
  const double step = static_cast<double>(kDefaultColumnPixelStep);
  const double pixels =
      std::ceil((defaults.base_col_width * mdw + static_cast<double>(model.padding_px)) / step) * step;
  return std::floor(pixels / mdw * kStoredWidthSteps) / kStoredWidthSteps;
}

}  // namespace formulon
