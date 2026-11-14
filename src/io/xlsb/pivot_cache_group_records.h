// Cache field grouping records (`fieldGroup`): BrtBeginPCDFGroup with an
// optional BrtBeginPCDFGRange (`rangePr`) and an optional BrtBeginPCDFGItems
// list whose `groupItems` sit in one BrtBeginPCDIRun of strings.
//
// A block is decoded from a span of records starting at BrtBeginPCDFGroup and
// encoded back to the identical records. Discrete grouping and item runs of
// any type other than strings have no measured layout and are refused.

#ifndef FORMULON_IO_XLSB_PIVOT_CACHE_GROUP_RECORDS_H_
#define FORMULON_IO_XLSB_PIVOT_CACHE_GROUP_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/pivot_records_common.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_group {

/// `par` / `base` value meaning the attribute is absent.
inline constexpr std::uint32_t kNoField = 0xFFFFFFFFU;

/// BrtBeginPCDFGRange byte 0 (`groupBy`): `days`, the only value measured.
inline constexpr std::uint8_t kGroupByDays = 4U;
/// BrtBeginPCDFGRange byte 1 bits: `autoStart`, `autoEnd`, and a date range
/// (`startDate`/`endDate` rather than `startNum`/`endNum`).
inline constexpr std::uint8_t kRangeAutoStart = 0x01U;
inline constexpr std::uint8_t kRangeAutoEnd = 0x02U;
inline constexpr std::uint8_t kRangeDates = 0x04U;

/// One `rangePr` (BrtBeginPCDFGRange .. BrtEndPCDFGRange).
struct RangeGroup {
  std::uint8_t group_by = kGroupByDays;
  std::uint8_t flags = 0;  ///< kRange* bits.
  double start = 0.0;      ///< Serial for a date range, else the number.
  double end = 0.0;
  double interval = 1.0;
};

/// One `fieldGroup` (BrtBeginPCDFGroup .. BrtEndPCDFGroup).
struct FieldGroup {
  std::uint32_t par = kNoField;
  std::uint32_t base = kNoField;
  std::optional<RangeGroup> range;
  /// `groupItems` strings; nullopt when the group carries no item list.
  std::optional<std::vector<std::string>> items;
};

/// Decodes the block starting at `records[*pos]` (BrtBeginPCDFGroup) and
/// advances `*pos` past its BrtEndPCDFGroup.
Expected<FieldGroup, Error> decode_field_group(const std::vector<pivot_records::RawRecord>& records, std::size_t* pos);

/// Appends the records of `group` to `*out`.
void encode_field_group(const FieldGroup& group, std::vector<pivot_records::RawRecord>* out);

}  // namespace pivot_group
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_CACHE_GROUP_RECORDS_H_
