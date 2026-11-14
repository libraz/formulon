// Pivot rule (`pivotArea`) records: BrtBeginPRule, its PRFilters / PRFilter
// list and the PRFItem indices under each filter. The same block appears
// under a cache calculated item and under a table field's autoSortScope.
//
// A block is decoded from a span of records starting at BrtBeginPRule and
// encoded back to the identical records. Payload fields whose meaning is not
// established are carried as raw integers.

#ifndef FORMULON_IO_XLSB_PIVOT_AREA_RECORDS_H_
#define FORMULON_IO_XLSB_PIVOT_AREA_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "io/xlsb/pivot_records_common.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_area {

/// Field index meaning "no field" (pivotArea@field absent).
inline constexpr std::uint32_t kNoField = 0xFFFFFFFFU;
/// Field index of the data pseudo-field (`4294967294`).
inline constexpr std::uint32_t kDataField = 0xFFFFFFFEU;

/// BrtBeginPRule `type` byte for `normal`, the only type Excel's parts show.
inline constexpr std::uint8_t kAreaTypeNormal = 1U;
/// BrtBeginPRule flag bits (bytes 5-7, little-endian): `dataOnly` and `cacheIndex`.
inline constexpr std::uint32_t kAreaDataOnly = 0x000001U;
inline constexpr std::uint32_t kAreaCacheIndex = 0x000010U;

/// BrtBeginPRFilter flag bits (bytes 8-10, little-endian): `selected`.
inline constexpr std::uint32_t kFilterSelected = 0x010000U;

/// One `reference` (BrtBeginPRFilter .. BrtEndPRFilter).
struct PivotAreaFilter {
  std::uint32_t field = 0;           ///< Field index, or kDataField.
  std::uint32_t flags = 0;           ///< 24-bit flag block; see kFilterSelected.
  std::vector<std::uint32_t> items;  ///< One BrtBeginPRFItem index per `x`.
};

/// One `pivotArea` (BrtBeginPRule .. BrtEndPRule).
struct PivotArea {
  std::uint32_t field = kNoField;
  std::uint8_t type = kAreaTypeNormal;
  std::uint32_t flags = 0;  ///< 24-bit flag block; see kArea* bits.
  /// The BrtBeginPRFilters list; nullopt when the rule carries none.
  std::optional<std::vector<PivotAreaFilter>> filters;
};

/// Decodes the block starting at `records[*pos]` (BrtBeginPRule) and advances
/// `*pos` past its BrtEndPRule. Any record outside the grammar is an error.
Expected<PivotArea, Error> decode_pivot_area(const std::vector<pivot_records::RawRecord>& records, std::size_t* pos);

/// Appends the records of `area` to `*out`.
void encode_pivot_area(const PivotArea& area, std::vector<pivot_records::RawRecord>* out);

}  // namespace pivot_area
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_AREA_RECORDS_H_
