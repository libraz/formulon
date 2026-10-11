#include "new_workbook_defaults.h"

#include <cstddef>
#include <string>
#include <utility>

#include "sheet.h"
#include "sheet_passthrough.h"
#include "styles.h"

namespace formulon {
namespace {

constexpr NewSheetMargins kDefaultMargins = {
    "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>",
    PageMargins{}};
constexpr NewSheetMargins kDeMargins = {
    "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.78740157499999996\" bottom=\"0.78740157499999996\" "
    "header=\"0.3\" footer=\"0.3\"/>",
    PageMargins{ooxml_defaults::kPageMarginSideInches, ooxml_defaults::kPageMarginSideInches, 0.78740157499999996,
                0.78740157499999996, ooxml_defaults::kPageMarginHeaderFooterInches,
                ooxml_defaults::kPageMarginHeaderFooterInches}};
constexpr NewSheetMargins kPtBrMargins = {
    "<pageMargins left=\"0.511811024\" right=\"0.511811024\" top=\"0.78740157499999996\" "
    "bottom=\"0.78740157499999996\" header=\"0.31496062000000002\" footer=\"0.31496062000000002\"/>",
    PageMargins{0.511811024, 0.511811024, 0.78740157499999996, 0.78740157499999996, 0.31496062000000002,
                0.31496062000000002}};

// @size-budget: 2 KB
// Indexed by `ExcelLocale`; values from Mac Excel 365 16.113 new workbooks.
constexpr NewWorkbookDefaults kDefaults[] = {
    {"游ゴシック", 128, 6, false, "標準", 20, "Sheet1", &kDefaultMargins},      // ja_JP
    {"Aptos Narrow", 0, 0, false, "Normal", 16, "Sheet1", &kDefaultMargins},    // en_US
    {"Aptos Narrow", 0, 0, false, "Standard", 16, "Tabelle1", &kDeMargins},     // de_DE
    {"Aptos Narrow", 0, 0, false, "Normal", 16, "Feuil1", &kDefaultMargins},    // fr_FR
    {"等线", 134, 9, true, "常规", 16, "Sheet1", &kDefaultMargins},             // zh_CN
    {"맑은 고딕", 129, 8, true, "표준", 18, "Sheet1", &kDefaultMargins},        // ko_KR
    {"Tahoma", 222, 0, false, "ปกติ", 15, "Sheet1", &kDefaultMargins},           // th_TH
    {"Aptos Narrow", 0, 0, false, "Normal", 16, "Hoja1", &kDefaultMargins},     // es_ES
    {"Aptos Narrow", 0, 0, false, "Normal", 16, "Hoja1", &kDefaultMargins},     // es_MX
    {"Aptos Narrow", 0, 0, false, "Normal", 16, "Planilha1", &kPtBrMargins},    // pt_BR
    {"Aptos Narrow", 204, 0, false, "Обычный", 16, "Лист1", &kDefaultMargins},  // ru_RU
    {"新細明體", 136, 9, true, "一般", 15, "工作表1", &kDefaultMargins},        // zh_TW
    {"Aptos Narrow", 0, 0, false, "Normale", 16, "Foglio1", &kDefaultMargins},  // it_IT
    {"Aptos Narrow", 0, 0, false, "Standaard", 16, "Blad1", &kDefaultMargins},  // nl_NL
};
static_assert(sizeof(kDefaults) / sizeof(kDefaults[0]) == static_cast<std::size_t>(ExcelLocale::kNlNL) + 1U,
              "one row per ExcelLocale");

/// `<scheme val="minor"/>`, as `FontRecord::scheme` stores it.
constexpr std::uint8_t kSchemeMinor = 2U;
/// `CellXf::vertical_align` ordinal of `center`.
constexpr std::uint8_t kVerticalCenter = 1U;

void center_vertically(CellXf& xf) {
  xf.vertical_align = kVerticalCenter;
  xf.has_vertical_align = true;
  xf.has_alignment = true;
}

}  // namespace

const NewWorkbookDefaults& new_workbook_defaults(ExcelLocale locale) {
  return kDefaults[static_cast<std::size_t>(locale)];
}

void apply_new_workbook_defaults(ExcelProfile profile, StylesTable& styles, Sheet& sheet) {
  const NewWorkbookDefaults& row = new_workbook_defaults(profile.locale);
  const bool mac = profile.host == ExcelHost::kMac365;

  styles.fonts.reserve(2U);
  FontRecord& body = styles.fonts[0];
  body.name = row.font_name;
  body.has_charset = row.charset != 0U;
  body.charset = row.charset;
  body.scheme = kSchemeMinor;

  if (row.phonetic_size != 0U) {
    // The phonetic font is the body face without a colour, at its own size.
    styles.fonts.emplace_back(body).size = row.phonetic_size;
    center_vertically(styles.cell_style_xfs[0]);
    center_vertically(styles.cell_xfs[0]);
    sheet.mutable_raw_extensions().push_back(
        WorksheetRawChild{static_cast<std::uint32_t>(worksheet_child::slot_of("phoneticPr")),
                          row.phonetic_no_conversion ? "<phoneticPr fontId=\"1\" type=\"noConversion\"/>"
                                                     : "<phoneticPr fontId=\"1\"/>"});
  }
  body.size = mac ? kNewWorkbookMacFontSize : kNewWorkbookWinFontSize;
  body.color.kind = ColorSpec::Kind::kTheme;
  body.color.theme = 1U;
  styles.cell_styles[0].name = row.normal_style_name;

  if (mac) {
    SheetFormatDefaults& format = sheet.mutable_format_defaults();
    format.base_col_width = kNewWorkbookMacBaseColWidth;
    format.default_row_height = row.mac_row_height;
    format.has_default_row_height = true;
  }
  SheetPrintSettings& print = sheet.mutable_print_settings();
  print.page_margins_xml = row.margins->xml;
  print.page_margins = row.margins->values;
}

}  // namespace formulon
