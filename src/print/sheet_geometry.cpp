#include "print/sheet_geometry.h"

#include <algorithm>
#include <cmath>
#include <unordered_set>
#include <utility>

#include "sheet.h"
#include "styles.h"

namespace formulon {
namespace print {
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
/// for any (Normal font family, size) `kColumnWidthCalibrations` below
/// does not cover.
constexpr double kPrintPointsPerColumnCharCalibri11 = 39.0 / 7.0;
constexpr double kPrintColumnPaddingPtCalibri11 = 27.0 / 7.0;

/// Calibri 11's *display* points-per-character-unit and per-column
/// padding (`Range.Width` at 96 DPI) -- the `kColumnWidthCalibrations`
/// entry for `{"Calibri", 11}`, restated here so the print/display
/// ratios below don't depend on table lookup order.
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

/// One (Normal font family, size) -> (MDW, padding) *display* calibration
/// point on Windows Excel 365 ja-JP (build 16.0.20228, 100%
/// display scaling / 96 DPI) via `Range.Width` at stored widths 30/100
/// chars. `Calibri`/11 reproduces `kDisplayPointsPerColumnCharCalibri11` /
/// `kDisplayColumnPaddingPtCalibri11` above exactly.
struct ColumnWidthCalibration {
  const char* family;
  int size;
  double mdw_pt;
  double pad_pt;
};

constexpr ColumnWidthCalibration kColumnWidthCalibrations[] = {
    {"Calibri", 8, 4.5, 3.75},          {"Calibri", 9, 4.5, 3.75},           {"Calibri", 10, 5.25, 3.75},
    {"Calibri", 11, 5.25, 3.75},        {"Calibri", 12, 6.0, 3.75},          {"Calibri", 14, 7.5, 5.25},
    {"Calibri", 16, 8.25, 5.25},        {"Calibri", 18, 9.0, 5.25},          {"ＭＳ Ｐゴシック", 8, 4.5, 3.75},
    {"ＭＳ Ｐゴシック", 9, 4.5, 3.75},  {"ＭＳ Ｐゴシック", 10, 5.25, 3.75}, {"ＭＳ Ｐゴシック", 11, 6.0, 3.75},
    {"ＭＳ Ｐゴシック", 12, 6.0, 3.75}, {"ＭＳ Ｐゴシック", 14, 7.5, 5.25},  {"ＭＳ Ｐゴシック", 16, 8.25, 5.25},
    {"ＭＳ Ｐゴシック", 18, 9.0, 5.25}, {"游ゴシック", 8, 4.5, 3.75},        {"游ゴシック", 9, 5.25, 3.75},
    {"游ゴシック", 10, 5.25, 3.75},     {"游ゴシック", 11, 6.0, 3.75},       {"游ゴシック", 12, 6.75, 5.25},
    {"游ゴシック", 14, 8.25, 5.25},     {"游ゴシック", 16, 9.0, 5.25},       {"游ゴシック", 18, 9.75, 6.75},
    {"Meiryo UI", 8, 5.25, 3.75},       {"Meiryo UI", 9, 5.25, 3.75},        {"Meiryo UI", 10, 6.0, 3.75},
    {"Meiryo UI", 11, 6.75, 5.25},      {"Meiryo UI", 12, 7.5, 5.25},        {"Meiryo UI", 14, 9.0, 5.25},
    {"Meiryo UI", 16, 9.75, 6.75},      {"Meiryo UI", 18, 11.25, 6.75},
};

/// The resolved points-per-character-unit and per-column padding for either
/// mode, plus whether the font was a tabulated calibration point.
struct ColumnWidthGeometry {
  double points_per_char = kPrintPointsPerColumnCharCalibri11;
  double padding_pt = kPrintColumnPaddingPtCalibri11;
  bool calibrated = false;
};

/// Looks up the *display* calibration for `(family, size)` in
/// `kColumnWidthCalibrations`. Falls back to the Calibri-11 display
/// figure (with `calibrated == false`) for any `(family, size)` pair the
/// table does not cover -- interpolating between sampled sizes would be
/// inventing data the capture does not support.
ColumnWidthGeometry ResolveColumnDisplayGeometry(const std::string& family, double size) {
  for (const ColumnWidthCalibration& row : kColumnWidthCalibrations) {
    if (static_cast<double>(row.size) == size && family == row.family) {
      return ColumnWidthGeometry{row.mdw_pt, row.pad_pt, true};
    }
  }
  return ColumnWidthGeometry{kDisplayPointsPerColumnCharCalibri11, kDisplayColumnPaddingPtCalibri11, false};
}

/// Height of one row override: hidden is zero, an explicit height wins,
/// anything else keeps the sheet default.
double OverrideRowHeight(const RowLayout& layout, double default_height) {
  if (layout.hidden) {
    return 0.0;
  }
  return (layout.has_height || layout.height != 0.0) ? layout.height : default_height;
}

/// Width of one column span in character units: hidden is zero, an explicit
/// width wins, anything else keeps the sheet default.
double SpanColumnWidth(const ColumnLayout& span, double default_width) {
  if (span.hidden) {
    return 0.0;
  }
  return HasExplicitColumnWidth(span) ? span.width : default_width;
}

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

ColumnWidthModel resolve_column_width_model(const StylesTable& styles, GeometryMode mode) {
  const FontRecord* font = resolve_normal_font(styles);
  ColumnWidthModel model;
  ColumnWidthGeometry display{kDisplayPointsPerColumnCharCalibri11, kDisplayColumnPaddingPtCalibri11, false};
  if (font != nullptr) {
    display = ResolveColumnDisplayGeometry(font->name, font->size);
    model.normal_font_name = font->name;
    model.normal_font_size = font->size;
  }
  // Print scales the display calibration by Calibri 11's print/display ratios.
  const ColumnWidthGeometry print{display.points_per_char * kPrintMdwRatio, display.padding_pt * kPrintPaddingRatio,
                                  display.calibrated};
  const ColumnWidthGeometry& geometry = mode == GeometryMode::kDisplay ? display : print;
  model.points_per_char = geometry.points_per_char;
  model.padding_pt = geometry.padding_pt;
  model.max_digit_width_px = static_cast<int>(std::lround(display.points_per_char / kPointsPerPixel));
  model.padding_px = static_cast<int>(std::lround(display.padding_pt / kPointsPerPixel));
  model.calibrated = geometry.calibrated;
  return model;
}

double column_chars_to_points(double chars, const ColumnWidthModel& model) {
  // A hidden or explicit zero-width column converts to exactly zero so it
  // never advances a page break or shifts the fit-to-page scale (Excel
  // excludes it from pagination extent); the flat padding term models the
  // allowance every *visible* column carries.
  if (chars == 0.0) {
    return 0.0;
  }
  const double mdw = static_cast<double>(model.max_digit_width_px);
  const double pixels =
      std::floor((kStoredWidthSteps * chars + std::floor(kStoredWidthSteps / 2.0 / mdw)) / kStoredWidthSteps * mdw);
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

double default_column_width_chars(const Sheet& sheet, const ColumnWidthModel& model) {
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  if (defaults.has_default_col_width) {
    return defaults.default_col_width;
  }
  const double mdw = static_cast<double>(model.max_digit_width_px);
  const double step = static_cast<double>(kDefaultColumnPixelStep);
  const double pixels =
      std::ceil((defaults.base_col_width * mdw + static_cast<double>(model.padding_px)) / step) * step;
  return std::floor(pixels / mdw * kStoredWidthSteps) / kStoredWidthSteps;
}

double default_row_height_pt(const Sheet& sheet) {
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  return defaults.has_default_row_height ? defaults.default_row_height : ooxml_defaults::kStandardRowHeightPt;
}

/// The nominal default remains available to callers that need the stored
/// height (for example, an explicitly visible row without its own `ht`). The
/// implicit extent is zero when `<sheetFormatPr zeroHeight="1">` hides rows by
/// default.
double implicit_row_height_pt(const Sheet& sheet) {
  return sheet.format_defaults().zero_height ? 0.0 : default_row_height_pt(sheet);
}

double effective_column_width_chars(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model) {
  const double default_width = default_column_width_chars(sheet, model);
  for (const ColumnLayout& span : sheet.layout().columns) {
    if (col >= span.first && col <= span.last) {
      // Hidden columns occupy no printed width, so they never advance the
      // page grid.
      return SpanColumnWidth(span, default_width);
    }
  }
  return default_width;
}

double effective_row_height_pt(const Sheet& sheet, std::uint32_t row) {
  const double default_height = default_row_height_pt(sheet);
  for (const RowLayout& override_row : sheet.layout().row_overrides) {
    if (override_row.row == row) {
      return OverrideRowHeight(override_row, default_height);
    }
  }
  return implicit_row_height_pt(sheet);
}

std::vector<double> column_widths_chars(const Sheet& sheet, std::uint32_t first, std::uint32_t last,
                                        const ColumnWidthModel& model) {
  const double default_width = default_column_width_chars(sheet, model);
  std::vector<double> widths(static_cast<std::size_t>(last - first) + 1U, default_width);
  std::vector<bool> assigned(widths.size(), false);
  for (const ColumnLayout& span : sheet.layout().columns) {
    const std::uint32_t span_first = std::max(span.first, first);
    const std::uint32_t span_last = std::min(span.last, last);
    for (std::uint32_t col = span_first; col <= span_last; ++col) {
      const auto index = static_cast<std::size_t>(col - first);
      if (assigned[index]) {
        continue;
      }
      assigned[index] = true;
      widths[index] = SpanColumnWidth(span, default_width);
    }
  }
  return widths;
}

std::vector<double> row_heights_pt(const Sheet& sheet, std::uint32_t first, std::uint32_t last) {
  const double default_height = default_row_height_pt(sheet);
  std::vector<double> heights(static_cast<std::size_t>(last - first) + 1U, implicit_row_height_pt(sheet));
  std::vector<bool> assigned(heights.size(), false);
  for (const RowLayout& layout : sheet.layout().row_overrides) {
    if (layout.row < first || layout.row > last) {
      continue;
    }
    const auto index = static_cast<std::size_t>(layout.row - first);
    if (assigned[index]) {
      continue;
    }
    assigned[index] = true;
    heights[index] = OverrideRowHeight(layout, default_height);
  }
  return heights;
}

double effective_column_width_pt(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model) {
  return column_chars_to_points(effective_column_width_chars(sheet, col, model), model);
}

double column_offset_pt(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model) {
  if (col == 0U) {
    return 0.0;
  }
  double offset = 0.0;
  for (const double chars : column_widths_chars(sheet, 0U, col - 1U, model)) {
    offset += column_chars_to_points(chars, model);
  }
  return offset;
}

double row_offset_pt(const Sheet& sheet, std::uint32_t row) {
  // Rows are numerous (up to 2^20), so adjust the default-height run by the
  // overrides instead of materialising every row.
  const double default_height = default_row_height_pt(sheet);
  const double implicit_height = implicit_row_height_pt(sheet);
  double offset = implicit_height * static_cast<double>(row);
  if (row == 0U) {
    return offset;
  }
  const auto& overrides = sheet.layout().row_overrides;
  const bool ordered =
      std::adjacent_find(overrides.begin(), overrides.end(), [](const RowLayout& lhs, const RowLayout& rhs) {
        return lhs.row >= rhs.row;
      }) == overrides.end();
  if (ordered) {
    for (const RowLayout& layout : overrides) {
      if (layout.row >= row) {
        break;
      }
      offset += OverrideRowHeight(layout, default_height) - implicit_height;
    }
    return offset;
  }
  // Preserve first-override precedence for directly constructed layouts
  // that bypass the sorted, unique row-edit helpers.
  std::unordered_set<std::uint32_t> seen;
  seen.reserve(overrides.size());
  for (const RowLayout& layout : overrides) {
    if (layout.row >= row) {
      continue;
    }
    if (!seen.insert(layout.row).second) {
      continue;
    }
    offset += OverrideRowHeight(layout, default_height) - implicit_height;
  }
  return offset;
}

RectPt cell_rect_pt(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const ColumnWidthModel& model) {
  RectPt rect;
  rect.x = column_offset_pt(sheet, col, model);
  rect.y = row_offset_pt(sheet, row);
  rect.width = effective_column_width_pt(sheet, col, model);
  rect.height = effective_row_height_pt(sheet, row);
  return rect;
}

}  // namespace print
}  // namespace formulon
