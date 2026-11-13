#include "font_metrics.h"

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
namespace {

struct FontMetricsRow {
  const char* family;
  int size;
  double mdw_pt;
  double pad_pt;
  double row_height_pt;
};

// Mac Excel 365, whole points (1 px = 1 pt). Every row is measured.
constexpr FontMetricsRow kMacRows[] = {
    {"Calibri", 11, 6.0, 5.0, 15.0},         {"Calibri", 14, 7.0, 5.0, 19.0},
    {"Calibri", 20, 10.0, 7.0, 26.0},        {"Aptos Narrow", 11, 6.0, 5.0, 15.0},
    {"Aptos Narrow", 12, 6.0, 5.0, 16.0},    {"游ゴシック", 11, 6.0, 5.0, 18.0},
    {"游ゴシック", 12, 7.0, 5.0, 20.0},      {"等线", 12, 6.0, 5.0, 16.0},
    {"맑은 고딕", 12, 7.0, 5.0, 18.0},       {"Tahoma", 12, 7.0, 5.0, 15.0},
    {"新細明體", 12, 6.0, 5.0, 15.0},        {"ＭＳ Ｐゴシック", 11, 6.0, 5.0, 14.0},
    {"Calibri", 12, 6.0, 5.0, 16.0},         {"Meiryo", 11, 7.0, 5.0, 19.0},
    {"Aptos", 11, 6.0, 5.0, 15.0},           {"Aptos", 12, 6.0, 5.0, 16.0},
    {"Arial", 10, 6.0, 5.0, 13.0},           {"Arial", 11, 6.0, 5.0, 14.0},
    {"Arial", 12, 7.0, 5.0, 16.0},           {"Calibri", 10, 5.0, 5.0, 14.0},
    {"Cambria", 11, 6.0, 5.0, 14.0},         {"Times New Roman", 12, 6.0, 5.0, 16.0},
    {"Verdana", 10, 6.0, 5.0, 13.0},         {"Helvetica", 12, 7.0, 5.0, 16.0},
    {"Helvetica Neue", 12, 7.0, 5.0, 16.0},  {"Meiryo", 10, 6.0, 5.0, 17.0},
    {"Meiryo", 12, 7.0, 5.0, 20.0},          {"Meiryo UI", 9, 6.0, 5.0, 14.0},
    {"Meiryo UI", 11, 7.0, 5.0, 16.0},       {"ＭＳ ゴシック", 11, 6.0, 5.0, 14.0},
    {"ＭＳ 明朝", 11, 6.0, 5.0, 14.0},       {"游明朝", 11, 6.0, 5.0, 18.0},
    {"メイリオ", 11, 7.0, 5.0, 19.0},        {"游ゴシック", 10, 6.0, 5.0, 17.0},
    {"游ゴシック", 14, 8.0, 5.0, 24.0},      {"Aptos Narrow", 10, 5.0, 5.0, 14.0},
    {"Aptos Narrow", 14, 7.0, 5.0, 19.0},    {"等线", 11, 6.0, 5.0, 15.0},
    {"맑은 고딕", 11, 6.0, 5.0, 17.0},       {"Tahoma", 10, 5.0, 5.0, 13.0},
    {"Tahoma", 11, 6.0, 5.0, 14.0},          {"Courier New", 10, 6.0, 5.0, 14.0},
    {"ＭＳ Ｐゴシック", 12, 6.0, 5.0, 15.0}, {"新細明體", 11, 5.0, 5.0, 14.0},
};

// Windows Excel 365 ja-JP (build 16.0.20228, 96 DPI) `Range.Width` at stored
// widths 30/100 chars. Row heights are not measured per font. Calibri 11 is
// the Windows fallback and matches the print-layout derivation in
// `column_width.cpp`.
constexpr FontMetricsRow kWinRows[] = {
    {"Calibri", 8, 4.5, 3.75, 0.0},           {"Calibri", 9, 4.5, 3.75, 0.0},
    {"Calibri", 10, 5.25, 3.75, 0.0},         {"Calibri", 11, 5.25, 3.75, 0.0},
    {"Calibri", 12, 6.0, 3.75, 0.0},          {"Calibri", 14, 7.5, 5.25, 0.0},
    {"Calibri", 16, 8.25, 5.25, 0.0},         {"Calibri", 18, 9.0, 5.25, 0.0},
    {"ＭＳ Ｐゴシック", 8, 4.5, 3.75, 0.0},   {"ＭＳ Ｐゴシック", 9, 4.5, 3.75, 0.0},
    {"ＭＳ Ｐゴシック", 10, 5.25, 3.75, 0.0}, {"ＭＳ Ｐゴシック", 11, 6.0, 3.75, 0.0},
    {"ＭＳ Ｐゴシック", 12, 6.0, 3.75, 0.0},  {"ＭＳ Ｐゴシック", 14, 7.5, 5.25, 0.0},
    {"ＭＳ Ｐゴシック", 16, 8.25, 5.25, 0.0}, {"ＭＳ Ｐゴシック", 18, 9.0, 5.25, 0.0},
    {"游ゴシック", 8, 4.5, 3.75, 0.0},        {"游ゴシック", 9, 5.25, 3.75, 0.0},
    {"游ゴシック", 10, 5.25, 3.75, 0.0},      {"游ゴシック", 11, 6.0, 3.75, 0.0},
    {"游ゴシック", 12, 6.75, 5.25, 0.0},      {"游ゴシック", 14, 8.25, 5.25, 0.0},
    {"游ゴシック", 16, 9.0, 5.25, 0.0},       {"游ゴシック", 18, 9.75, 6.75, 0.0},
    {"Meiryo UI", 8, 5.25, 3.75, 0.0},        {"Meiryo UI", 9, 5.25, 3.75, 0.0},
    {"Meiryo UI", 10, 6.0, 3.75, 0.0},        {"Meiryo UI", 11, 6.75, 5.25, 0.0},
    {"Meiryo UI", 12, 7.5, 5.25, 0.0},        {"Meiryo UI", 14, 9.0, 5.25, 0.0},
    {"Meiryo UI", 16, 9.75, 6.75, 0.0},       {"Meiryo UI", 18, 11.25, 6.75, 0.0},
};

struct MacAlias {
  const char* alias;
  const char* family;
};

// English names a Mac workbook may store for a tabulated face. Not measured
// under these names, so they are estimates.
constexpr MacAlias kMacAliases[] = {
    {"MS PGothic", "ＭＳ Ｐゴシック"},
    {"DengXian", "等线"},
    {"Malgun Gothic", "맑은 고딕"},
    {"PMingLiU", "新細明體"},
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

/// ISO 15924 script of the minor-font entry a locale reads; empty for `a:latin`.
const char* minor_script_for(ExcelLocale locale) {
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

/// The face a Mac Normal font linked to the theme's minor font is drawn with.
std::string resolve_minor_face(const Theme& theme, ExcelLocale locale) {
  const std::string_view script = minor_script_for(locale);
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
  out.mdw_pt = row->mdw_pt;
  out.padding_pt = row->pad_pt;
  if (mac) {
    out.row_height_pt = row->row_height_pt;
    out.row_height_calibrated = out.width_calibrated;
  } else {
    out.row_height_pt = ooxml_defaults::kStandardRowHeightPt;
  }
  return out;
}

}  // namespace formulon
