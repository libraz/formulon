//
// MS-XLSB data-validation container codec: BrtBeginDVals .. BrtEndDVals to
// and from `Sheet::validations()`. Layouts follow Excel 365 (macOS)
// re-saves against the .xlsx Excel wrote for the same book:
//
//   * BrtBeginDVals (573): fWnClosed u16, xLeft, yTop, unused u32 (all 0
//     in every sample), then idvMac u32 = number of BrtDVal records.
//   * BrtDVal (64): a u32 flag word -- valType bits 0-3, errStyle 4-6,
//     fStrLookup 7 (set exactly for an inline `"a,b"` list), fAllowBlank 8,
//     fSuppressCombo 9 (OOXML showDropDown), mdImeMode 10-17 (always 0),
//     fShowInputMsg 18, fShowErrorMsg 19, typOperator 20-23 -- then an
//     Sqrfx, errorTitle / error / promptTitle / prompt as nullable wide
//     strings, and two formulas (cce 0 when absent). valType, errStyle and
//     typOperator number their values as `DataValidation` does.
//   * Each BrtDVal may be preceded by an alternate-content wrapper holding
//     its revision uid (BrtACBegin, BrtUid, BrtACEnd); the model does not
//     carry `xr:uid`, as with .xlsx, so the wrapper is consumed.
//   * A validation naming another sheet is stored only in the future-record
//     container BrtFRTBegin, BrtBeginDVals14 (1054: a zero FRTHeader, then
//     BrtBeginDVals's fields), BrtDVal14 (1053) records, BrtEndDVals14
//     (1154), BrtFRTEnd. BrtDVal14 is an FRTHeader (fSqref, plus fFormula
//     when a formula is present), one FRTSqref (flags 0x02, Sqrfx), the
//     FRTFormulas in slot order, then BrtDVal's flag word with bit 24 / 25
//     set for a present formula1 / formula2, and its four strings. There is
//     no separate formula size field.
//
// A list validation's formula is reference class, every other one value
// class, all relative to the sqref's bounding-box top-left.

#ifndef FORMULON_IO_XLSB_DV_RECORDS_H_
#define FORMULON_IO_XLSB_DV_RECORDS_H_

#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

#include "io/xlsb/feature_formula.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "sheet_passthrough.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kBrtBeginDVals = 573;
inline constexpr std::uint16_t kBrtEndDVals = 574;

/// Decodes one framed container. Returns nothing when any validation in it
/// holds content outside the supported set; the caller then keeps the bytes.
std::optional<std::vector<DataValidation>> decode_dv_block(ByteSpan block, const FeatureFormulaReadContext& ctx);

/// Emits the validations of `validations` that `data_validation_in_x14`
/// leaves in the main records as one framed container; nothing when none
/// covers a cell.
Expected<void, Error> emit_dv_block(std::vector<std::uint8_t>& dst, const std::vector<DataValidation>& validations,
                                    const FeatureFormulaWriteContext& ctx, std::string_view own_sheet);

/// Appends every BrtDVal14 of the retained tail `records` to `validations`,
/// marked `x14_placement`. The records stay in the tail as the slots
/// `place_x14_validations` rewrites: a retained future record is what keeps
/// the source BrtExternSheet table seeding the written one, which the 3-D
/// references of a record it cannot adopt index.
void take_x14_validations(const std::vector<std::uint8_t>& records, const FeatureFormulaReadContext& ctx,
                          std::vector<DataValidation>& validations);

/// Writes the validations `data_validation_in_x14` selects into `tail`'s
/// BrtDVal14 records: the first retained container's adoptable records are
/// rewritten from them in order and the rest appended to it, or to a new
/// container after the x14 CF one. A record `take_x14_validations` could not
/// adopt stays verbatim; a container left empty goes.
Expected<void, Error> place_x14_validations(XlsbSheetTail& tail, const std::vector<DataValidation>& validations,
                                            const FeatureFormulaWriteContext& ctx, std::string_view own_sheet);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_DV_RECORDS_H_
