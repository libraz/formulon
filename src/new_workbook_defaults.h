//
// What a new Excel workbook starts as in each UI locale: Normal font,
// phonetic font, Normal style name, first sheet name, row metrics and page
// margins. Values are those of the empty workbooks Mac Excel 365 saves per
// locale; Windows reuses them with its own body size and leaves the row
// metrics to the engine defaults.

#ifndef FORMULON_NEW_WORKBOOK_DEFAULTS_H_
#define FORMULON_NEW_WORKBOOK_DEFAULTS_H_

#include <cstdint>

#include "excel_profile.h"
#include "sheet_layout.h"

namespace formulon {

class Sheet;
struct StylesTable;

/// Body (font 0) size of a new workbook, in points, per host.
constexpr double kNewWorkbookMacFontSize = 12.0;
constexpr double kNewWorkbookWinFontSize = 11.0;

/// `sheetFormatPr baseColWidth` of a new Mac workbook, in characters.
constexpr double kNewWorkbookMacBaseColWidth = 10.0;

/// Page margins of a new sheet and the `<pageMargins>` element Excel writes
/// for them.
struct NewSheetMargins {
  const char* xml;
  PageMargins values;
};

/// One locale's row of the new-workbook table.
struct NewWorkbookDefaults {
  const char* font_name;  ///< Font 0 and phonetic font face.
  /// `<charset>` of font 0 and the phonetic font; 0 when Excel writes none.
  std::uint8_t charset;
  /// Phonetic font (font 1) size in points; 0 when the locale has none. A
  /// locale with a phonetic font also centres cells vertically.
  std::uint8_t phonetic_size;
  bool phonetic_no_conversion;  ///< `phoneticPr type="noConversion"`.
  const char* normal_style_name;
  std::uint8_t mac_row_height;  ///< Mac `defaultRowHeight`, points.
  const char* sheet_name;
  const NewSheetMargins* margins;
};

/// The table row for `locale`.
const NewWorkbookDefaults& new_workbook_defaults(ExcelLocale locale);

/// Applies `profile`'s row to the seeded style table and first sheet of a
/// new workbook: font 0 / font 1, the xf alignment, the Normal style name,
/// the sheet row metrics, page margins and `phoneticPr`.
void apply_new_workbook_defaults(ExcelProfile profile, StylesTable& styles, Sheet& sheet);

/// Applies the sheet half of `profile`'s row to a sheet added to a new
/// workbook later: the row metrics, page margins and `phoneticPr` its
/// first sheet received.
void apply_new_sheet_defaults(ExcelProfile profile, Sheet& sheet);

}  // namespace formulon

#endif  // FORMULON_NEW_WORKBOOK_DEFAULTS_H_
