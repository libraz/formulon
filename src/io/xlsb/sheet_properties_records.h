// BrtWsProp codec: the sheet's `<sheetPr>` code name and tab colour to and
// from the worksheet-properties record; with it the other sheet-level records
// the model owns: BrtWsFmtInfo, BrtPhoneticInfo and the sheet's revision uid.

#ifndef FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_
#define FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "io/xlsb/record.h"
#include "phonetic.h"
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

/// Reads `BrtPhoneticInfo` (u16 ifnt, u32 phType, u32 alcH); false for any
/// other shape or an ordinal past 3.
bool decode_phonetic_info(ByteSpan payload, PhoneticProperties& out);

/// Emits `BrtPhoneticInfo` for the sheet's `<phoneticPr>`; nothing for a
/// sheet without one.
void emit_phonetic_info(std::vector<std::uint8_t>& dst, const Sheet& sheet);

/// Stores the 16-byte GUID of a sheet's revision-uid wrapper as the
/// worksheet root's `xr:uid`.
void set_sheet_uid(Sheet& sheet, const std::uint8_t* guid);

/// Emits the sheet's revision-uid wrapper (BrtACBegin, BrtUid, BrtACEnd),
/// which ends the worksheet. The GUID is the root's `xr:uid`; a sheet
/// without one gets `{00000000-0001-0000-XX00-000000000000}`, `XX` its
/// zero-based position, which is what Excel assigns a sheet it loads
/// without a uid.
void emit_sheet_uid(std::vector<std::uint8_t>& dst, const Sheet& sheet, std::size_t sheet_index);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_
