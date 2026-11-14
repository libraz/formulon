// BrtWsProp codec: the sheet's `<sheetPr>` code name, tab colour, outline and
// fit-to-page flags to and from the worksheet-properties record; with it the other sheet-level records
// the model owns: BrtWsFmtInfo, BrtPhoneticInfo and the sheet's revision uid.

#ifndef FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_
#define FORMULON_IO_XLSB_SHEET_PROPERTIES_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
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

/// Decodes `BrtWsFmtInfo` into the sheet's default column/row metrics, a
/// default row height past what Excel opens clamped as Excel does. False,
/// leaving the defaults untouched, for a truncated record, a reserved flag
/// bit, or a width, base width or outline level Excel does not open.
bool decode_ws_fmt_info(const XlsbRecord& rec, Sheet& sheet);

/// Emits `BrtWsFmtInfo` for `defaults`, sizes truncated to the record's
/// units. Fails with `kIoFeatureNotWritable` (context `kind=format_defaults`
/// or `kind=outline_level`, and `sheet_index`) for a metric or outline level
/// Excel does not hold.
Expected<void, Error> emit_ws_fmt_info(std::vector<std::uint8_t>& dst, const SheetFormatDefaults& defaults,
                                       std::size_t sheet_index);

/// True when BrtWsProp writes every member of the `<sheetPr>` fragment `raw`:
/// the code name, the tab colour, `<outlinePr>` applyStyles and
/// showOutlineSymbols (summaryBelow / summaryRight only at their default) and
/// `<pageSetUpPr>` fitToPage. An empty fragment is fully written.
bool sheet_pr_writes_all(std::string_view raw);

/// Emits the worksheet-properties record which starts the mandatory
/// worksheet prefix.
void emit_ws_prop(std::vector<std::uint8_t>& dst, const Sheet& sheet);

/// Reads `BrtPhoneticInfo` (u16 ifnt, u32 phType, u32 alcH); false for any
/// other shape or an ordinal past 3.
bool decode_phonetic_info(ByteSpan payload, PhoneticProperties& out);

/// Emits `BrtPhoneticInfo` for the sheet's `<phoneticPr>`, else for
/// `absent` (`FontSavePlan::absent_sheet_phonetic`); nothing when both are empty.
void emit_phonetic_info(std::vector<std::uint8_t>& dst, const Sheet& sheet,
                        const std::optional<PhoneticProperties>& absent);

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
