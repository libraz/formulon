//
// Record ids of the three XLSB pivot parts (pivot table, cache definition,
// cache records) and the verbatim record a codec hands on to another one.
// Names follow [MS-XLSB] where its grammar gives the number, otherwise the
// grammar slot Excel's own parts show the record in.

#ifndef FORMULON_IO_XLSB_PIVOT_RECORDS_COMMON_H_
#define FORMULON_IO_XLSB_PIVOT_RECORDS_COMMON_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <vector>

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_records {

// Shared wrappers.
inline constexpr std::uint16_t kACBegin = 37;
inline constexpr std::uint16_t kACEnd = 38;
inline constexpr std::uint16_t kUid = 3072;

// Pivot table part.
inline constexpr std::uint16_t kBeginSXView = 280;
inline constexpr std::uint16_t kEndSXView = 315;
inline constexpr std::uint16_t kBeginSXLocation = 314;
inline constexpr std::uint16_t kEndSXLocation = 313;
inline constexpr std::uint16_t kBeginSXVDs = 287;
inline constexpr std::uint16_t kEndSXVDs = 288;
inline constexpr std::uint16_t kBeginSXVD = 285;
inline constexpr std::uint16_t kEndSXVD = 286;
inline constexpr std::uint16_t kBeginSXVIs = 283;
inline constexpr std::uint16_t kEndSXVIs = 284;
inline constexpr std::uint16_t kBeginSXVI = 282;
inline constexpr std::uint16_t kEndSXVI = 281;
inline constexpr std::uint16_t kBeginAutoSortScope = 459;
inline constexpr std::uint16_t kEndAutoSortScope = 460;
inline constexpr std::uint16_t kBeginISXVDRws = 309;
inline constexpr std::uint16_t kEndISXVDRws = 310;
inline constexpr std::uint16_t kBeginSXLIRws = 299;
inline constexpr std::uint16_t kEndSXLIRws = 300;
inline constexpr std::uint16_t kBeginSXLI = 297;
inline constexpr std::uint16_t kEndSXLI = 298;
inline constexpr std::uint16_t kBeginISXVIs = 388;
inline constexpr std::uint16_t kEndISXVIs = 389;
inline constexpr std::uint16_t kBeginISXVDCols = 311;
inline constexpr std::uint16_t kEndISXVDCols = 312;
inline constexpr std::uint16_t kBeginSXLICols = 301;
inline constexpr std::uint16_t kEndSXLICols = 302;
inline constexpr std::uint16_t kBeginSXPIs = 291;
inline constexpr std::uint16_t kEndSXPIs = 292;
inline constexpr std::uint16_t kBeginSXPI = 289;
inline constexpr std::uint16_t kEndSXPI = 290;
inline constexpr std::uint16_t kBeginSXDIs = 295;
inline constexpr std::uint16_t kEndSXDIs = 296;
inline constexpr std::uint16_t kBeginSXDI = 293;
inline constexpr std::uint16_t kEndSXDI = 294;
inline constexpr std::uint16_t kTableStyleClient = 513;
inline constexpr std::uint16_t kBeginSXFilters = 599;
inline constexpr std::uint16_t kEndSXFilters = 600;
inline constexpr std::uint16_t kBeginSXFilter = 601;
inline constexpr std::uint16_t kEndSXFilter = 602;
inline constexpr std::uint16_t kBeginSXView14 = 1062;
inline constexpr std::uint16_t kEndSXView14 = 1063;
inline constexpr std::uint16_t kBeginSXView16 = 5000;
inline constexpr std::uint16_t kEndSXView16 = 5001;
inline constexpr std::uint16_t kBeginSXFormats = 305;
inline constexpr std::uint16_t kEndSXFormats = 306;
inline constexpr std::uint16_t kBeginSXFormat = 303;
inline constexpr std::uint16_t kEndSXFormat = 304;
inline constexpr std::uint16_t kBeginSXCrtFormats = 483;
inline constexpr std::uint16_t kEndSXCrtFormats = 484;
inline constexpr std::uint16_t kBeginSXCrtFormat = 481;
inline constexpr std::uint16_t kEndSXCrtFormat = 482;
/// The x14 `<pivotField>` record (FRTSXVD slot), a lone record inside an FRT block.
inline constexpr std::uint16_t kSXVD14 = 1061;
/// The x14 `<dataField>` record (FRTSXDI slot), a lone record inside an FRT block.
inline constexpr std::uint16_t kSXDI14 = 1044;

// Pivot rules (pivotArea), used by the table's autoSortScope, formats and chart formats, and the cache's calculated
// items.
inline constexpr std::uint16_t kBeginPRule = 247;
inline constexpr std::uint16_t kEndPRule = 248;
inline constexpr std::uint16_t kBeginPRFilters = 249;
inline constexpr std::uint16_t kEndPRFilters = 250;
inline constexpr std::uint16_t kBeginPRFilter = 251;
inline constexpr std::uint16_t kEndPRFilter = 252;
inline constexpr std::uint16_t kBeginPRFItem = 382;
inline constexpr std::uint16_t kEndPRFItem = 383;

// Cache definition part.
inline constexpr std::uint16_t kBeginPivotCacheDef = 179;
inline constexpr std::uint16_t kEndPivotCacheDef = 180;
inline constexpr std::uint16_t kBeginPCDFields = 181;
inline constexpr std::uint16_t kEndPCDFields = 182;
inline constexpr std::uint16_t kBeginPCDField = 183;
inline constexpr std::uint16_t kEndPCDField = 184;
inline constexpr std::uint16_t kBeginPCDSource = 185;
inline constexpr std::uint16_t kEndPCDSource = 186;
inline constexpr std::uint16_t kBeginPCDSRange = 187;
inline constexpr std::uint16_t kEndPCDSRange = 188;
inline constexpr std::uint16_t kBeginPCDFAtbl = 189;
inline constexpr std::uint16_t kEndPCDFAtbl = 190;
inline constexpr std::uint16_t kPCDIMissing = 20;
inline constexpr std::uint16_t kPCDINumber = 21;
inline constexpr std::uint16_t kPCDIError = 23;
inline constexpr std::uint16_t kPCDIString = 24;
inline constexpr std::uint16_t kPCDIDatetime = 25;
inline constexpr std::uint16_t kPCDIAMissing = 27;
inline constexpr std::uint16_t kPCDIAString = 31;
inline constexpr std::uint16_t kBeginPCDIRun = 191;
inline constexpr std::uint16_t kEndPCDIRun = 192;
inline constexpr std::uint16_t kBeginPCDFGroup = 219;
inline constexpr std::uint16_t kEndPCDFGroup = 220;
inline constexpr std::uint16_t kBeginPCDFGItems = 221;
inline constexpr std::uint16_t kEndPCDFGItems = 222;
inline constexpr std::uint16_t kBeginPCDFGRange = 223;
inline constexpr std::uint16_t kEndPCDFGRange = 224;
inline constexpr std::uint16_t kBeginPCDFGDiscrete = 225;
inline constexpr std::uint16_t kEndPCDFGDiscrete = 226;

/// Excel writes 4 or more numbers, strings or dates in a row as one BrtBeginPCDIRun and fewer as single records
/// (3 numbers and 3 group items were single records; 4 strings, 4 dates and 12 dates one run each); booleans and
/// errors always single.
inline constexpr std::size_t kMinItemRunLength = 4;

inline constexpr std::uint16_t kBeginPCDCalcItems = 243;
inline constexpr std::uint16_t kEndPCDCalcItems = 244;
inline constexpr std::uint16_t kBeginPCDCalcItem = 245;
inline constexpr std::uint16_t kEndPCDCalcItem = 246;
inline constexpr std::uint16_t kBeginPNames = 253;
inline constexpr std::uint16_t kEndPNames = 254;
inline constexpr std::uint16_t kBeginPName = 255;
inline constexpr std::uint16_t kEndPName = 256;
inline constexpr std::uint16_t kBeginPNPairs = 257;
inline constexpr std::uint16_t kEndPNPairs = 258;
inline constexpr std::uint16_t kBeginPNPair = 259;
inline constexpr std::uint16_t kEndPNPair = 260;
inline constexpr std::uint16_t kBeginPCD14 = 1066;
inline constexpr std::uint16_t kEndPCD14 = 1067;
inline constexpr std::uint16_t kDataSourceClean = 5153;
/// The version-info blocks Excel writes on a cache (lastRefreshFeature) and a table (lastUpdateFeature) once rich
/// values reach them, each holding one feature record, and the cache's richInfo.
inline constexpr std::uint16_t kBeginPCDVersionInfo = 5111;
inline constexpr std::uint16_t kEndPCDVersionInfo = 5112;
inline constexpr std::uint16_t kBeginSXVersionInfo = 5109;
inline constexpr std::uint16_t kEndSXVersionInfo = 5110;
inline constexpr std::uint16_t kVersionInfoFeature = 5114;
inline constexpr std::uint16_t kPCDRichInfo = 5105;
/// BrtFRTBegin payload around both version-info blocks.
inline constexpr std::array<std::uint8_t, 4> kVersionInfoFrtVersion = {0x41, 0x00, 0x02, 0x00};
/// The feature record's payload: a reserved u32 0, then the XLWideString `RichData`, the only feature measured.
inline constexpr std::array<std::uint8_t, 24> kRichDataFeature = {0x00, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00,
                                                                  0x52, 0x00, 0x69, 0x00, 0x63, 0x00, 0x68, 0x00,
                                                                  0x44, 0x00, 0x61, 0x00, 0x74, 0x00, 0x61, 0x00};

// Cache records part.
inline constexpr std::uint16_t kBeginPivotCacheRecords = 193;
inline constexpr std::uint16_t kEndPivotCacheRecords = 194;
inline constexpr std::uint16_t kPCRRecord = 33;
/// BrtPCRRecordDt: opens a row spelt as one item record per database field.
inline constexpr std::uint16_t kPCRRecordDt = 34;
/// BrtPCDIIndex: a shared-item index inside a BrtPCRRecordDt row, or a group item index inside a
/// BrtBeginPCDFGDiscrete map.
inline constexpr std::uint16_t kPCDIIndex = 26;

/// One record kept verbatim: a codec hands a block it does not own to the
/// codec that does, and re-emits it unchanged until that codec decodes it.
struct RawRecord {
  std::uint16_t type = 0;
  std::vector<std::uint8_t> payload;
};

}  // namespace pivot_records
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_RECORDS_COMMON_H_
