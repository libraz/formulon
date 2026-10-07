// BrtWsProp codec: the sheet's `<sheetPr>` code name and tab colour to and
// from the worksheet-properties record.

#ifndef FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_
#define FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_

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

/// Emits `BrtWsFmtInfo` for `defaults`.
void emit_ws_fmt_info(std::vector<std::uint8_t>& dst, const SheetFormatDefaults& defaults);

/// Emits the worksheet-properties record which starts the mandatory
/// worksheet prefix.
void emit_ws_prop(std::vector<std::uint8_t>& dst, const Sheet& sheet);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_
