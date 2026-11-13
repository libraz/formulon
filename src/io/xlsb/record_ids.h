//
// XLSB record ids that more than one translation unit dispatches on, for the
// records the `XlsbRecordType` reader enum does not list.

#ifndef FORMULON_IO_XLSB_RECORD_IDS_H_
#define FORMULON_IO_XLSB_RECORD_IDS_H_

#include <cstdint>

#include "io/xlsb/record.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kBrtEndBook = static_cast<std::uint16_t>(XlsbRecordType::BrtEndBook);

// AutoFilter block.
inline constexpr std::uint16_t kBrtEndAFilter = 162;
inline constexpr std::uint16_t kBrtBeginFilterColumn = 163;  // u32 column id
inline constexpr std::uint16_t kBrtEndFilterColumn = 164;
inline constexpr std::uint16_t kBrtBeginFilters = 165;
inline constexpr std::uint16_t kBrtEndFilters = 166;
inline constexpr std::uint16_t kBrtFilter = 167;
inline constexpr std::uint16_t kBrtColorFilter = 168;
inline constexpr std::uint16_t kBrtIconFilter = 169;
inline constexpr std::uint16_t kBrtTop10Filter = 170;
inline constexpr std::uint16_t kBrtDynamicFilter = 171;
inline constexpr std::uint16_t kBrtBeginCustomFilters = 172;
inline constexpr std::uint16_t kBrtEndCustomFilters = 173;
inline constexpr std::uint16_t kBrtCustomFilter = 174;
inline constexpr std::uint16_t kBrtAFilterDateGroupItem = 175;

// Row and column page breaks.
inline constexpr std::uint16_t kBrtBeginRwBrk = 392;  // u32 count, u32 manual count
inline constexpr std::uint16_t kBrtEndRwBrk = 393;
inline constexpr std::uint16_t kBrtBeginColBrk = 394;
inline constexpr std::uint16_t kBrtEndColBrk = 395;
inline constexpr std::uint16_t kBrtBrk = 396;  // u32 id, min, max, u32 manual

// Scenarios.
inline constexpr std::uint16_t kBrtBeginScenMan = 500;  // u16 current, u16 show
inline constexpr std::uint16_t kBrtEndScenMan = 501;
inline constexpr std::uint16_t kBrtBeginSct = 502;  // u16 input-cell count
inline constexpr std::uint16_t kBrtEndSct = 503;
inline constexpr std::uint16_t kBrtSlc = 504;  // u32 row, u32 col

// Sort state and range protection.
inline constexpr std::uint16_t kBrtBeginSortState = 530;  // RfX at 2
inline constexpr std::uint16_t kBrtEndSortState = 531;
inline constexpr std::uint16_t kBrtBeginSortCond = 532;  // RfX at 2
inline constexpr std::uint16_t kBrtEndSortCond = 533;
inline constexpr std::uint16_t kBrtRangeProtection = 536;  // Sqrfx at 2
inline constexpr std::uint16_t kBrtRangeProtectionIso = 680;

// Worksheet records between the merge block and the hyperlinks.
inline constexpr std::uint16_t kBrtDVal = 64;
inline constexpr std::uint16_t kBrtBeginCFRule = 463;
inline constexpr std::uint16_t kBrtEndCFRule = 464;
inline constexpr std::uint16_t kBrtBeginIconSet = 465;
inline constexpr std::uint16_t kBrtEndIconSet = 466;
inline constexpr std::uint16_t kBrtBeginDataBar = 467;
inline constexpr std::uint16_t kBrtEndDataBar = 468;
inline constexpr std::uint16_t kBrtBeginColorScale = 469;
inline constexpr std::uint16_t kBrtEndColorScale = 470;
inline constexpr std::uint16_t kBrtCFVO = 471;
inline constexpr std::uint16_t kBrtPhoneticInfo = 537;
inline constexpr std::uint16_t kBrtColor = 564;  // conditional-formatting colour

// Worksheet tail grammar slots after the hyperlinks.
inline constexpr std::uint16_t kBrtDrawing = 550;
inline constexpr std::uint16_t kBrtLegacyDrawing = 551;
inline constexpr std::uint16_t kBrtLegacyDrawingHF = 552;
inline constexpr std::uint16_t kBrtBeginWebPubItems = 554;  // u32 item count
inline constexpr std::uint16_t kBrtEndWebPubItems = 555;
inline constexpr std::uint16_t kBrtBeginWebPubItem = 556;  // u32 source type, RfX at 9
inline constexpr std::uint16_t kBrtEndWebPubItem = 557;
inline constexpr std::uint16_t kBrtBkhim = 562;
inline constexpr std::uint16_t kBrtCellSmartTagProperty = 589;
inline constexpr std::uint16_t kBrtBeginCellSmartTag = 590;
inline constexpr std::uint16_t kBrtEndCellSmartTag = 591;
inline constexpr std::uint16_t kBrtBeginCellSmartTags = 592;
inline constexpr std::uint16_t kBrtEndCellSmartTags = 593;
inline constexpr std::uint16_t kBrtBeginSmartTags = 594;
inline constexpr std::uint16_t kBrtEndSmartTags = 595;
inline constexpr std::uint16_t kBrtBeginCellWatches = 605;
inline constexpr std::uint16_t kBrtEndCellWatches = 606;
inline constexpr std::uint16_t kBrtCellWatch = 607;  // u32 row, u32 col
inline constexpr std::uint16_t kBrtBigName = 625;
inline constexpr std::uint16_t kBrtBeginOleObjects = 638;
inline constexpr std::uint16_t kBrtOleObject = 639;
inline constexpr std::uint16_t kBrtEndOleObjects = 640;
inline constexpr std::uint16_t kBrtBeginActiveXControls = 643;
inline constexpr std::uint16_t kBrtActiveX = 644;
inline constexpr std::uint16_t kBrtEndActiveXControls = 645;
inline constexpr std::uint16_t kBrtBeginCellIgnoreECs = 648;
inline constexpr std::uint16_t kBrtCellIgnoreEC = 649;  // u32 flags, Sqrfx at 4
inline constexpr std::uint16_t kBrtEndCellIgnoreECs = 650;
inline constexpr std::uint16_t kBrtBeginTableParts = 660;
inline constexpr std::uint16_t kBrtTablePart = 661;
inline constexpr std::uint16_t kBrtEndTableParts = 662;

// Table part records.
inline constexpr std::uint16_t kBrtBeginList = 343;
inline constexpr std::uint16_t kBrtBeginListCol = 347;

// Custom table styles (styles part). BrtTableStyleElement: u32 tseType, u32 size, u32 dxf index.
inline constexpr std::uint16_t kBrtBeginTableStyles = 508;
inline constexpr std::uint16_t kBrtEndTableStyles = 509;
inline constexpr std::uint16_t kBrtBeginTableStyle = 510;
inline constexpr std::uint16_t kBrtEndTableStyle = 511;
inline constexpr std::uint16_t kBrtTableStyleElement = 512;

// Page setup block.
inline constexpr std::uint16_t kBrtMargins = 476;
inline constexpr std::uint16_t kBrtPrintOptions = 477;
inline constexpr std::uint16_t kBrtPageSetup = 478;
inline constexpr std::uint16_t kBrtBeginHeaderFooter = 479;
inline constexpr std::uint16_t kBrtEndHeaderFooter = 480;

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_RECORD_IDS_H_
