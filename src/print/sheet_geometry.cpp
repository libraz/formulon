#include "print/sheet_geometry.h"

#include <algorithm>
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
// Only Calibri 11's print-layout width is directly measured (from the
// print-track captures' own break positions, not a `Range.Width` sweep).
// Other fonts/sizes' print-layout widths are unmeasured; they are
// inferred by scaling that font's *display* calibration by the ratio
// between Calibri 11's print and display figures, separately for the
// per-character rate and the padding term (`kPrintMdwRatio` /
// `kPrintPaddingRatio` below) -- an approximation, not a second sweep.

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
/// padding term, derived from Calibri 11 (the only font with a directly
/// measured print-layout figure) and applied to every other font's
/// display calibration to approximate its print-layout one.
constexpr double kPrintMdwRatio = kPrintPointsPerColumnCharCalibri11 / kDisplayPointsPerColumnCharCalibri11;
constexpr double kPrintPaddingRatio = kPrintColumnPaddingPtCalibri11 / kDisplayColumnPaddingPtCalibri11;

/// Excel's standard default column width, in character units. Used when
/// neither a `<col>` override nor `<sheetFormatPr defaultColWidth>`
/// applies.
constexpr double kStandardColWidthChars = 8.43;

/// One (Normal font family, size) -> (MDW, padding) *display* calibration
/// point, measured on Windows Excel 365 ja-JP (build 16.0.20228, 100%
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
/// mode, plus whether the font was a measured calibration point.
struct ColumnWidthGeometry {
  double points_per_char = kPrintPointsPerColumnCharCalibri11;
  double padding_pt = kPrintColumnPaddingPtCalibri11;
  bool calibrated = false;
};

/// Looks up the measured *display* calibration for `(family, size)` in
/// `kColumnWidthCalibrations`. Falls back to the Calibri-11 display
/// figure (with `calibrated == false`) for any family or integer size the
/// table does not cover, and for any non-integer size -- interpolating
/// between the sampled sizes would be inventing data the capture does not
/// support.
ColumnWidthGeometry ResolveColumnDisplayGeometry(const std::string& family, double size) {
  const int size_int = static_cast<int>(size);
  if (static_cast<double>(size_int) == size) {
    for (const ColumnWidthCalibration& row : kColumnWidthCalibrations) {
      if (row.size == size_int && family == row.family) {
        return ColumnWidthGeometry{row.mdw_pt, row.pad_pt, true};
      }
    }
  }
  return ColumnWidthGeometry{kDisplayPointsPerColumnCharCalibri11, kDisplayColumnPaddingPtCalibri11, false};
}

/// The geometry pagination's break math must use: `(family, size)`'s
/// display calibration scaled by Calibri 11's print/display ratios. Reduces
/// to the measured `kPrintPointsPerColumnCharCalibri11` /
/// `kPrintColumnPaddingPtCalibri11` exactly for `{"Calibri", 11}` and for
/// any untabulated font/size.
ColumnWidthGeometry ResolveColumnPrintGeometry(const std::string& family, double size) {
  const ColumnWidthGeometry display = ResolveColumnDisplayGeometry(family, size);
  return ColumnWidthGeometry{display.points_per_char * kPrintMdwRatio, display.padding_pt * kPrintPaddingRatio,
                             display.calibrated};
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
    if (cell_style.name != "Normal") {
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
  ColumnWidthGeometry geometry;
  if (font != nullptr) {
    geometry = mode == GeometryMode::kDisplay ? ResolveColumnDisplayGeometry(font->name, font->size)
                                              : ResolveColumnPrintGeometry(font->name, font->size);
    model.normal_font_name = font->name;
    model.normal_font_size = font->size;
  } else {
    geometry = mode == GeometryMode::kDisplay
                   ? ColumnWidthGeometry{kDisplayPointsPerColumnCharCalibri11, kDisplayColumnPaddingPtCalibri11, false}
                   : ColumnWidthGeometry{};
  }
  model.points_per_char = geometry.points_per_char;
  model.padding_pt = geometry.padding_pt;
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
  return chars * model.points_per_char + model.padding_pt;
}

double column_points_to_chars(double points, const ColumnWidthModel& model) {
  if (points <= 0.0 || model.points_per_char <= 0.0) {
    return 0.0;
  }
  return std::max(0.0, (points - model.padding_pt) / model.points_per_char);
}

double default_column_width_chars(const Sheet& sheet) {
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  return defaults.has_default_col_width ? defaults.default_col_width : kStandardColWidthChars;
}

double default_row_height_pt(const Sheet& sheet) {
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  return defaults.has_default_row_height ? defaults.default_row_height : ooxml_defaults::kStandardRowHeightPt;
}

double effective_column_width_chars(const Sheet& sheet, std::uint32_t col) {
  const double default_width = default_column_width_chars(sheet);
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
  return default_height;
}

std::vector<double> column_widths_chars(const Sheet& sheet, std::uint32_t first, std::uint32_t last) {
  const double default_width = default_column_width_chars(sheet);
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
  std::vector<double> heights(static_cast<std::size_t>(last - first) + 1U, default_height);
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
  return column_chars_to_points(effective_column_width_chars(sheet, col), model);
}

double column_offset_pt(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model) {
  if (col == 0U) {
    return 0.0;
  }
  double offset = 0.0;
  for (const double chars : column_widths_chars(sheet, 0U, col - 1U)) {
    offset += column_chars_to_points(chars, model);
  }
  return offset;
}

double row_offset_pt(const Sheet& sheet, std::uint32_t row) {
  // Rows are numerous (up to 2^20), so adjust the default-height run by the
  // overrides instead of materialising every row.
  const double default_height = default_row_height_pt(sheet);
  double offset = default_height * static_cast<double>(row);
  std::vector<std::uint32_t> seen;
  for (const RowLayout& layout : sheet.layout().row_overrides) {
    if (layout.row >= row) {
      continue;
    }
    if (std::find(seen.begin(), seen.end(), layout.row) != seen.end()) {
      continue;
    }
    seen.push_back(layout.row);
    offset += OverrideRowHeight(layout, default_height) - default_height;
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
