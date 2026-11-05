// BrtWsProp codec: the sheet's `<sheetPr>` code name and tab colour to and
// from the worksheet-properties record.

#ifndef FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_
#define FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Sheet;
struct SheetFormatDefaults;

namespace io {
namespace xlsb {

/// Decodes `BrtWsProp` into the sheet's raw `<sheetPr>` fragment. Returns
/// `false` when the record carried flag or sync fields the model has no
/// room for.
Expected<bool, Error> decode_ws_prop(const XlsbRecord& rec, Sheet& sheet, std::size_t sheet_index);

/// Decodes `BrtWsFmtInfo` into the sheet's default column/row metrics.
Expected<void, Error> decode_ws_fmt_info(const XlsbRecord& rec, Sheet& sheet, std::size_t sheet_index);

/// Whether a `<sheetFormatPr>` base column width fits BrtWsFmtInfo.
inline bool IsRepresentableBaseColumnWidth(double value) {
  return std::isfinite(value) && value >= 0.0 && value <= 255.0 && std::floor(value) == value;
}

/// Whether a default column width (character units) fits 1/256-unit storage.
inline bool IsRepresentableDefaultColumnWidth(double value) {
  constexpr double kMaxDefaultColumnWidth = 65535.0 / 256.0;
  return std::isfinite(value) && value >= 0.0 && value <= kMaxDefaultColumnWidth;
}

/// Whether a default row height (points) fits twip storage.
inline bool IsRepresentableDefaultRowHeight(double value) {
  constexpr double kMaxDefaultRowHeight = 65535.0 / 20.0;
  if (!std::isfinite(value) || value < 0.0 || value > kMaxDefaultRowHeight) {
    return false;
  }
  return std::isfinite(std::round(value * 20.0));
}

/// Emits `BrtWsFmtInfo` for `defaults`.
void emit_ws_fmt_info(std::vector<std::uint8_t>& dst, const SheetFormatDefaults& defaults);

/// Emits the worksheet-properties record which starts the mandatory
/// worksheet prefix.
void emit_ws_prop(std::vector<std::uint8_t>& dst, const Sheet& sheet);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_
