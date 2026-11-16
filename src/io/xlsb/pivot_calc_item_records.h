// Cache calculated-item records (`calculatedItems`): BrtBeginPCDCalcItems
// and one BrtBeginPCDCalcItem per item, each holding the item's formula as a
// PivotParsedFormula, its pivot rule (`pivotArea`) and the PNames table that
// the formula's PtgSxName tokens index. A PName lists (field, item) pairs.
// A calculated field's formula carries its own PNames table, whose names hold
// a cache field index and no pairs list.
//
// A block is decoded from a span of records starting at BrtBeginPCDCalcItems
// and encoded back to the identical records. The formula stays as rgce/rgcb
// bytes, read and written through the shared Ptg codec.

#ifndef FORMULON_IO_XLSB_PIVOT_CALC_ITEM_RECORDS_H_
#define FORMULON_IO_XLSB_PIVOT_CALC_ITEM_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "io/xlsb/pivot_area_records.h"
#include "io/xlsb/pivot_records_common.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_calc {

/// Field index meaning the attribute is absent (`calculatedItem@field`).
inline constexpr std::uint32_t kNoField = 0xFFFFFFFFU;

/// One (field, item) pair of a name (BrtBeginPNPair .. BrtEndPNPair).
struct PNPair {
  std::uint8_t flags = 0;  ///< Byte 0, carried as is.
  std::uint32_t field = 0;
  std::uint32_t item = 0;
};

/// One name (BrtBeginPName with its BrtBeginPNPairs list).
struct PName {
  std::uint32_t field = kNoField;  ///< Bytes 0-3, carried as is.
  std::uint16_t flags = 0;         ///< Bytes 4-5, carried as is.
  /// nullopt when the name has no BrtBeginPNPairs list, as a calculated field's names have not.
  std::optional<std::vector<PNPair>> pairs;
};

/// One `calculatedItem` (BrtBeginPCDCalcItem .. BrtEndPCDCalcItem).
struct CalcItem {
  std::uint32_t field = kNoField;
  std::vector<std::uint8_t> rgce;  ///< PivotParsedFormula token stream.
  std::vector<std::uint8_t> rgcb;  ///< PivotParsedFormula extra data.
  std::optional<pivot_area::PivotArea> area;
  /// The PNames table; nullopt when the item carries none.
  std::optional<std::vector<PName>> names;
};

/// Decodes the block starting at `records[*pos]` (BrtBeginPCDCalcItems) and
/// advances `*pos` past its BrtEndPCDCalcItems.
Expected<std::vector<CalcItem>, Error> decode_calc_items(const std::vector<pivot_records::RawRecord>& records,
                                                         std::size_t* pos);

/// Appends the records of `items` to `*out`.
void encode_calc_items(const std::vector<CalcItem>& items, std::vector<pivot_records::RawRecord>* out);

/// Decodes a PNames table (BrtBeginPNames .. BrtEndPNames) filling `records` exactly.
Expected<std::vector<PName>, Error> decode_names(const std::vector<pivot_records::RawRecord>& records);

/// Appends the records of a PNames table to `*out`.
void encode_names(const std::vector<PName>& names, std::vector<pivot_records::RawRecord>* out);

}  // namespace pivot_calc
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_CALC_ITEM_RECORDS_H_
