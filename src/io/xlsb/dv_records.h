//
// MS-XLSB data-validation container codec: BrtBeginDVals .. BrtEndDVals to
// and from `Sheet::validations()`. Measured from Excel 365 (macOS)
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
//
// A list validation's formula is reference class, every other one value
// class, all relative to the sqref's bounding-box top-left.

#ifndef FORMULON_IO_XLSB_DV_RECORDS_H_
#define FORMULON_IO_XLSB_DV_RECORDS_H_

#include <cstdint>
#include <optional>
#include <vector>

#include "io/xlsb/feature_formula.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kBrtBeginDVals = 573;
inline constexpr std::uint16_t kBrtEndDVals = 574;

/// Decodes one framed container. Returns nothing when any validation in it
/// holds content outside the measured set; the caller then keeps the bytes.
std::optional<std::vector<DataValidation>> decode_dv_block(ByteSpan block, const FeatureFormulaReadContext& ctx);

/// Emits `validations` as one framed container; nothing when none covers a cell.
Expected<void, Error> emit_dv_block(std::vector<std::uint8_t>& dst, const std::vector<DataValidation>& validations,
                                    const FeatureFormulaWriteContext& ctx);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_DV_RECORDS_H_
