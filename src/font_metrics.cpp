#include "font_metrics.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "column_width.h"
#include "new_workbook_defaults.h"
#include "sheet_layout.h"
#include "styles.h"
#include "theme.h"

// @size-budget: 5 KB

namespace formulon {

const char* theme_font_script(ExcelLocale locale) {
  switch (locale) {
    case ExcelLocale::kJaJP:
      return "Jpan";
    case ExcelLocale::kZhCN:
      return "Hans";
    case ExcelLocale::kZhTW:
      return "Hant";
    case ExcelLocale::kKoKR:
      return "Hang";
    case ExcelLocale::kThTH:
      return "Thai";
    default:
      return "";
  }
}

namespace {

struct FontMetricsRow {
  const char* family;
  int size;
  double mdw_pt;
  double row_height_pt;
};

// Mac Excel 16.115, whole points (1 px = 1 pt): MDW and `StandardHeight`, every row measured. Faces cover 8-12, 14,
// 16, 18 and 20 pt (Calibri and 游ゴシック also 13, 15, 22 and 24 pt). MDW depends on the face at a given size
// (11 pt: 新細明體 5, Calibri 6).
constexpr FontMetricsRow kMacRows[] = {
    {"Calibri", 8, 4.0, 11.0},
    {"Calibri", 9, 5.0, 12.0},
    {"Calibri", 10, 5.0, 14.0},
    {"Calibri", 11, 6.0, 15.0},
    {"Calibri", 12, 6.0, 16.0},
    {"Calibri", 13, 7.0, 17.0},
    {"Calibri", 14, 7.0, 19.0},
    {"Calibri", 15, 8.0, 20.0},
    {"Calibri", 16, 8.0, 21.0},
    {"Calibri", 18, 9.0, 24.0},
    {"Calibri", 20, 10.0, 26.0},
    {"Calibri", 22, 11.0, 29.0},
    {"Calibri", 24, 12.0, 31.0},
    {"Aptos", 8, 4.0, 11.0},
    {"Aptos", 9, 5.0, 13.0},
    {"Aptos", 10, 5.0, 14.0},
    {"Aptos", 11, 6.0, 15.0},
    {"Aptos", 12, 6.0, 16.0},
    {"Aptos", 14, 7.0, 19.0},
    {"Aptos", 16, 9.0, 22.0},
    {"Aptos", 18, 10.0, 24.0},
    {"Aptos", 20, 11.0, 27.0},
    {"Aptos Narrow", 8, 4.0, 11.0},
    {"Aptos Narrow", 9, 5.0, 13.0},
    {"Aptos Narrow", 10, 5.0, 14.0},
    {"Aptos Narrow", 11, 6.0, 15.0},
    {"Aptos Narrow", 12, 6.0, 16.0},
    {"Aptos Narrow", 14, 7.0, 19.0},
    {"Aptos Narrow", 16, 8.0, 22.0},
    {"Aptos Narrow", 18, 9.0, 24.0},
    {"Aptos Narrow", 20, 10.0, 27.0},
    {"游ゴシック", 8, 4.0, 13.0},
    {"游ゴシック", 9, 5.0, 16.0},
    {"游ゴシック", 10, 6.0, 17.0},
    {"游ゴシック", 11, 6.0, 18.0},
    {"游ゴシック", 12, 7.0, 20.0},
    {"游ゴシック", 13, 7.0, 23.0},
    {"游ゴシック", 14, 8.0, 24.0},
    {"游ゴシック", 15, 8.0, 25.0},
    {"游ゴシック", 16, 9.0, 27.0},
    {"游ゴシック", 18, 10.0, 31.0},
    {"游ゴシック", 20, 11.0, 33.0},
    {"游ゴシック", 22, 12.0, 38.0},
    {"游ゴシック", 24, 13.0, 40.0},
    {"Yu Gothic", 8, 4.0, 11.0},
    {"Yu Gothic", 9, 5.0, 13.0},
    {"Yu Gothic", 10, 6.0, 17.0},
    {"Yu Gothic", 11, 6.0, 17.0},
    {"Yu Gothic", 12, 7.0, 18.0},
    {"Yu Gothic", 14, 8.0, 22.0},
    {"Yu Gothic", 16, 9.0, 25.0},
    {"Yu Gothic", 18, 10.0, 29.0},
    {"Yu Gothic", 20, 11.0, 33.0},
    {"等线", 8, 4.0, 11.0},
    {"等线", 9, 5.0, 12.0},
    {"等线", 10, 5.0, 13.0},
    {"等线", 11, 6.0, 15.0},
    {"等线", 12, 6.0, 16.0},
    {"等线", 14, 7.0, 18.0},
    {"等线", 16, 8.0, 21.0},
    {"等线", 18, 9.0, 23.0},
    {"等线", 20, 11.0, 26.0},
    {"맑은 고딕", 8, 4.0, 13.0},
    {"맑은 고딕", 9, 5.0, 14.0},
    {"맑은 고딕", 10, 6.0, 15.0},
    {"맑은 고딕", 11, 6.0, 17.0},
    {"맑은 고딕", 12, 7.0, 18.0},
    {"맑은 고딕", 14, 8.0, 20.0},
    {"맑은 고딕", 16, 9.0, 23.0},
    {"맑은 고딕", 18, 10.0, 27.0},
    {"맑은 고딕", 20, 11.0, 30.0},
    {"新細明體", 8, 4.0, 11.0},
    {"新細明體", 9, 4.0, 12.0},
    {"新細明體", 10, 5.0, 13.0},
    {"新細明體", 11, 5.0, 14.0},
    {"新細明體", 12, 6.0, 15.0},
    {"新細明體", 14, 7.0, 20.0},
    {"新細明體", 16, 8.0, 22.0},
    {"新細明體", 18, 8.0, 25.0},
    {"新細明體", 20, 9.0, 27.0},
    {"Tahoma", 8, 4.0, 11.0},
    {"Tahoma", 9, 5.0, 12.0},
    {"Tahoma", 10, 5.0, 13.0},
    {"Tahoma", 11, 6.0, 14.0},
    {"Tahoma", 12, 7.0, 15.0},
    {"Tahoma", 14, 8.0, 18.0},
    {"Tahoma", 16, 9.0, 20.0},
    {"Tahoma", 18, 10.0, 23.0},
    {"Tahoma", 20, 11.0, 25.0},
    {"ＭＳ Ｐゴシック", 8, 4.0, 12.0},
    {"ＭＳ Ｐゴシック", 9, 5.0, 13.0},
    {"ＭＳ Ｐゴシック", 10, 5.0, 14.0},
    {"ＭＳ Ｐゴシック", 11, 6.0, 14.0},
    {"ＭＳ Ｐゴシック", 12, 6.0, 15.0},
    {"ＭＳ Ｐゴシック", 14, 7.0, 17.0},
    {"ＭＳ Ｐゴシック", 16, 8.0, 19.0},
    {"ＭＳ Ｐゴシック", 18, 9.0, 22.0},
    {"ＭＳ Ｐゴシック", 20, 10.0, 24.0},
    {"Meiryo", 8, 5.0, 13.0},
    {"Meiryo", 9, 6.0, 16.0},
    {"Meiryo", 10, 6.0, 17.0},
    {"Meiryo", 11, 7.0, 19.0},
    {"Meiryo", 12, 7.0, 20.0},
    {"Meiryo", 14, 9.0, 23.0},
    {"Meiryo", 16, 10.0, 26.0},
    {"Meiryo", 18, 11.0, 29.0},
    {"Meiryo", 20, 12.0, 32.0},
    {"Arial", 8, 4.0, 11.0},
    {"Arial", 9, 5.0, 12.0},
    {"Arial", 10, 6.0, 13.0},
    {"Arial", 11, 6.0, 14.0},
    {"Arial", 12, 7.0, 16.0},
    {"Arial", 14, 8.0, 18.0},
    {"Arial", 16, 9.0, 20.0},
    {"Arial", 18, 10.0, 23.0},
    {"Arial", 20, 11.0, 25.0},
    {"Cambria", 8, 4.0, 11.0},
    {"Cambria", 9, 5.0, 12.0},
    {"Cambria", 10, 6.0, 13.0},
    {"Cambria", 11, 6.0, 14.0},
    {"Cambria", 12, 7.0, 16.0},
    {"Cambria", 14, 8.0, 18.0},
    {"Cambria", 16, 9.0, 21.0},
    {"Cambria", 18, 10.0, 23.0},
    {"Cambria", 20, 11.0, 25.0},
    {"Times New Roman", 8, 4.0, 11.0},
    {"Times New Roman", 9, 5.0, 12.0},
    {"Times New Roman", 10, 5.0, 13.0},
    {"Times New Roman", 11, 6.0, 14.0},
    {"Times New Roman", 12, 6.0, 16.0},
    {"Times New Roman", 14, 7.0, 18.0},
    {"Times New Roman", 16, 8.0, 20.0},
    {"Times New Roman", 18, 9.0, 23.0},
    {"Times New Roman", 20, 10.0, 25.0},
    {"Verdana", 8, 5.0, 11.0},
    {"Verdana", 9, 6.0, 12.0},
    {"Verdana", 10, 6.0, 13.0},
    {"Verdana", 11, 7.0, 14.0},
    {"Verdana", 12, 8.0, 16.0},
    {"Verdana", 14, 9.0, 18.0},
    {"Verdana", 16, 10.0, 20.0},
    {"Verdana", 18, 11.0, 23.0},
    {"Verdana", 20, 13.0, 25.0},
    {"Helvetica", 8, 4.0, 11.0},
    {"Helvetica", 9, 5.0, 12.0},
    {"Helvetica", 10, 6.0, 13.0},
    {"Helvetica", 11, 6.0, 15.0},
    {"Helvetica", 12, 7.0, 16.0},
    {"Helvetica", 14, 8.0, 18.0},
    {"Helvetica", 16, 9.0, 21.0},
    {"Helvetica", 18, 10.0, 23.0},
    {"Helvetica", 20, 11.0, 26.0},
    {"Helvetica Neue", 8, 4.0, 11.0},
    {"Helvetica Neue", 9, 5.0, 12.0},
    {"Helvetica Neue", 10, 6.0, 13.0},
    {"Helvetica Neue", 11, 6.0, 14.0},
    {"Helvetica Neue", 12, 7.0, 16.0},
    {"Helvetica Neue", 14, 8.0, 18.0},
    {"Helvetica Neue", 16, 9.0, 20.0},
    {"Helvetica Neue", 18, 10.0, 23.0},
    {"Helvetica Neue", 20, 11.0, 25.0},
    {"Meiryo UI", 8, 5.0, 11.0},
    {"Meiryo UI", 9, 6.0, 14.0},
    {"Meiryo UI", 10, 6.0, 15.0},
    {"Meiryo UI", 11, 7.0, 16.0},
    {"Meiryo UI", 12, 7.0, 18.0},
    {"Meiryo UI", 14, 9.0, 20.0},
    {"Meiryo UI", 16, 10.0, 22.0},
    {"Meiryo UI", 18, 11.0, 25.0},
    {"Meiryo UI", 20, 12.0, 27.0},
    {"ＭＳ ゴシック", 8, 4.0, 12.0},
    {"ＭＳ ゴシック", 9, 5.0, 13.0},
    {"ＭＳ ゴシック", 10, 5.0, 14.0},
    {"ＭＳ ゴシック", 11, 6.0, 14.0},
    {"ＭＳ ゴシック", 12, 6.0, 15.0},
    {"ＭＳ ゴシック", 14, 7.0, 17.0},
    {"ＭＳ ゴシック", 16, 8.0, 19.0},
    {"ＭＳ ゴシック", 18, 9.0, 22.0},
    {"ＭＳ ゴシック", 20, 10.0, 24.0},
    {"ＭＳ 明朝", 8, 4.0, 12.0},
    {"ＭＳ 明朝", 9, 5.0, 13.0},
    {"ＭＳ 明朝", 10, 5.0, 14.0},
    {"ＭＳ 明朝", 11, 6.0, 14.0},
    {"ＭＳ 明朝", 12, 6.0, 15.0},
    {"ＭＳ 明朝", 14, 7.0, 17.0},
    {"ＭＳ 明朝", 16, 8.0, 19.0},
    {"ＭＳ 明朝", 18, 9.0, 22.0},
    {"ＭＳ 明朝", 20, 10.0, 24.0},
    {"游明朝", 8, 4.0, 13.0},
    {"游明朝", 9, 5.0, 16.0},
    {"游明朝", 10, 5.0, 17.0},
    {"游明朝", 11, 6.0, 18.0},
    {"游明朝", 12, 7.0, 20.0},
    {"游明朝", 14, 8.0, 24.0},
    {"游明朝", 16, 9.0, 27.0},
    {"游明朝", 18, 10.0, 31.0},
    {"游明朝", 20, 11.0, 33.0},
    {"Courier New", 8, 5.0, 11.0},
    {"Courier New", 9, 5.0, 13.0},
    {"Courier New", 10, 6.0, 14.0},
    {"Courier New", 11, 7.0, 15.0},
    {"Courier New", 12, 7.0, 17.0},
    {"Courier New", 14, 8.0, 19.0},
    {"Courier New", 16, 10.0, 22.0},
    {"Courier New", 18, 11.0, 24.0},
    {"Courier New", 20, 12.0, 27.0},
    {"Segoe UI", 8, 4.0, 11.0},
    {"Segoe UI", 9, 5.0, 13.0},
    {"Segoe UI", 10, 5.0, 14.0},
    {"Segoe UI", 11, 6.0, 16.0},
    {"Segoe UI", 12, 6.0, 16.0},
    {"Segoe UI", 14, 8.0, 21.0},
    {"Segoe UI", 16, 9.0, 23.0},
    {"Segoe UI", 18, 10.0, 26.0},
    {"Segoe UI", 20, 11.0, 28.0},
    {"Consolas", 8, 4.0, 11.0},
    {"Consolas", 9, 5.0, 12.0},
    {"Consolas", 10, 5.0, 14.0},
    {"Consolas", 11, 6.0, 15.0},
    {"Consolas", 12, 7.0, 16.0},
    {"Consolas", 14, 8.0, 19.0},
    {"Consolas", 16, 9.0, 23.0},
    {"Consolas", 18, 10.0, 26.0},
    {"Consolas", 20, 11.0, 28.0},
};

// Windows Excel 365 ja-JP (build 16.0.20228) at 100% display scaling (96 DPI), the scaling a win-* profile means:
// `Range.Width` at stored widths 30/100 chars and the new book's `StandardHeight`
// (tests/fixtures/excel/win/column_width_dpi100). Calibri 11 is the Windows fallback and matches the print-layout
// derivation in `column_width.cpp`.
constexpr FontMetricsRow kWinRows[] = {
    {"Calibri", 8, 4.5, 11.25},          {"Calibri", 9, 4.5, 12.0},           {"Calibri", 10, 5.25, 12.75},
    {"Calibri", 11, 5.25, 15.0},         {"Calibri", 12, 6.0, 15.75},         {"Calibri", 14, 7.5, 18.75},
    {"Calibri", 16, 8.25, 21.0},         {"Calibri", 18, 9.0, 23.25},         {"ＭＳ Ｐゴシック", 8, 4.5, 10.5},
    {"ＭＳ Ｐゴシック", 9, 4.5, 11.25},  {"ＭＳ Ｐゴシック", 10, 5.25, 12.0}, {"ＭＳ Ｐゴシック", 11, 6.0, 13.5},
    {"ＭＳ Ｐゴシック", 12, 6.0, 14.25}, {"ＭＳ Ｐゴシック", 14, 7.5, 17.25}, {"ＭＳ Ｐゴシック", 16, 8.25, 18.75},
    {"ＭＳ Ｐゴシック", 18, 9.0, 21.0},  {"游ゴシック", 8, 4.5, 12.75},       {"游ゴシック", 9, 5.25, 15.75},
    {"游ゴシック", 10, 5.25, 16.5},      {"游ゴシック", 11, 6.0, 18.75},      {"游ゴシック", 12, 6.75, 19.5},
    {"游ゴシック", 14, 8.25, 24.0},      {"游ゴシック", 16, 9.0, 25.5},       {"游ゴシック", 18, 9.75, 30.0},
    {"Meiryo UI", 8, 5.25, 12.0},        {"Meiryo UI", 9, 5.25, 12.0},        {"Meiryo UI", 10, 6.0, 14.25},
    {"Meiryo UI", 11, 6.75, 15.75},      {"Meiryo UI", 12, 7.5, 16.5},        {"Meiryo UI", 14, 9.0, 19.5},
    {"Meiryo UI", 16, 9.75, 21.0},       {"Meiryo UI", 18, 11.25, 24.0},
};

/// Points per screen pixel on each host: a pixel is a point on Mac; Windows is 96 DPI.
constexpr double kMacPointsPerPixel = 1.0;
constexpr double kWinPointsPerPixel = 72.0 / 96.0;

/// Column padding in screen pixels: a margin of ceil(MDW / 4) on each side plus the gridline pixel. Every measured
/// row of both hosts follows it (MDW 4 to 15 px).
int column_padding_px(int mdw_px) {
  return 2 * ((mdw_px + 3) / 4) + 1;
}

struct MacAlias {
  const char* alias;
  const char* family;
};

// Other spellings of a tabulated face, each measured under both names with identical width and row height at every
// tabulated size. A spelling whose measurements differ (Yu Gothic) has rows of its own.
constexpr MacAlias kMacAliases[] = {
    {"MS PGothic", "ＭＳ Ｐゴシック"}, {"DengXian", "等线"},   {"Malgun Gothic", "맑은 고딕"},
    {"PMingLiU", "新細明體"},          {"メイリオ", "Meiryo"},
};

template <std::size_t N>
const FontMetricsRow* find_row(const FontMetricsRow (&rows)[N], std::string_view family, double size) {
  for (const FontMetricsRow& row : rows) {
    if (static_cast<double>(row.size) == size && family == row.family) {
      return &row;
    }
  }
  return nullptr;
}

const FontMetricsRow* find_mac_row(std::string_view family, double size) {
  if (const FontMetricsRow* row = find_row(kMacRows, family, size)) {
    return row;
  }
  for (const MacAlias& alias : kMacAliases) {
    if (family == alias.alias) {
      return find_row(kMacRows, alias.family, size);
    }
  }
  return nullptr;
}

/// The face a Mac Normal font linked to the theme's minor font is drawn with.
std::string resolve_minor_face(const Theme& theme, ExcelLocale locale) {
  const std::string_view script = theme_font_script(locale);
  if (script.empty()) {
    if (!theme.fonts.minor_latin.empty()) {
      return theme.fonts.minor_latin;
    }
  } else {
    for (const ThemeScriptFont& entry : theme.minor_scripts) {
      if (entry.script == script && !entry.typeface.empty()) {
        return entry.typeface;
      }
    }
  }
  return new_workbook_defaults(locale).font_name;
}

}  // namespace

NormalFontMetrics normal_font_metrics(const StylesTable& styles, const Theme& theme, ExcelProfile profile) {
  const bool mac = profile.host == ExcelHost::kMac365;
  NormalFontMetrics out;
  out.host = profile.host;
  const FontRecord* font = resolve_normal_font(styles);
  const FontMetricsRow* row = nullptr;
  if (font != nullptr) {
    out.size_pt = font->size;
    out.family = mac && font->scheme == kFontSchemeMinor ? resolve_minor_face(theme, profile.locale) : font->name;
    row = mac ? find_mac_row(out.family, font->size) : find_row(kWinRows, out.family, font->size);
  }
  if (row == nullptr) {
    row = mac ? find_row(kMacRows, "Calibri", 11.0) : find_row(kWinRows, "Calibri", 11.0);
  } else {
    out.width_calibrated = true;
  }
  const double points_per_pixel = mac ? kMacPointsPerPixel : kWinPointsPerPixel;
  out.mdw_pt = row->mdw_pt;
  out.padding_pt = column_padding_px(static_cast<int>(std::lround(row->mdw_pt / points_per_pixel))) * points_per_pixel;
  out.row_height_pt = row->row_height_pt;
  out.row_height_calibrated = out.width_calibrated;
  return out;
}

double windows_row_height_pt(std::string_view family, double size) {
  const FontMetricsRow* row = find_row(kWinRows, family, size);
  return (row != nullptr ? row : find_row(kWinRows, "Calibri", 11.0))->row_height_pt;
}

}  // namespace formulon
