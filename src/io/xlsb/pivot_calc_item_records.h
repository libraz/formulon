// Cache calculated-item records (`calculatedItems`): BrtBeginPCDCalcItems
// and one BrtBeginPCDCalcItem per item, each holding the item's formula as a
// PivotParsedFormula, its pivot rule (`pivotArea`) and the PNames table that
// the formula's PtgSxName tokens index. A PName lists (field, item) pairs.
//
// A block is decoded from a span of records starting at BrtBeginPCDCalcItems
// and encoded back to the identical records. The formula stays as rgce/rgcb
// bytes; walk_calc_item_formula() tokenises it for callers that need names.

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

/// PtgSxName: ptg 0x18, eptg 0x1D, then a u32 index into the PNames table.
inline constexpr std::uint8_t kPtgSxName = 0x18U;
inline constexpr std::uint8_t kEptgSxName = 0x1DU;
/// PtgInt: ptg 0x1E, then a u16 integer.
inline constexpr std::uint8_t kPtgInt = 0x1EU;

/// One (field, item) pair of a name (BrtBeginPNPair .. BrtEndPNPair).
struct PNPair {
  std::uint32_t field = 0;
  std::uint8_t flags = 0;  ///< Byte 4, carried as is.
  std::uint32_t item = 0;
};

/// One name (BrtBeginPName with its BrtBeginPNPairs list).
struct PName {
  std::uint32_t field = kNoField;  ///< Bytes 0-3, carried as is.
  std::uint16_t flags = 0;         ///< Bytes 4-5, carried as is.
  std::vector<PNPair> pairs;
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

/// One token of a calculated-item formula.
struct FormulaToken {
  std::uint8_t ptg = 0;
  std::uint32_t sx_name = 0;  ///< PNames index; PtgSxName only.
  std::uint16_t integer = 0;  ///< PtgInt only.
};

/// Decodes the block starting at `records[*pos]` (BrtBeginPCDCalcItems) and
/// advances `*pos` past its BrtEndPCDCalcItems.
Expected<std::vector<CalcItem>, Error> decode_calc_items(const std::vector<pivot_records::RawRecord>& records,
                                                         std::size_t* pos);

/// Appends the records of `items` to `*out`.
void encode_calc_items(const std::vector<CalcItem>& items, std::vector<pivot_records::RawRecord>* out);

/// Tokenises `rgce`: PtgSxName, PtgInt and the payload-free operators 0x03-0x15.
/// Any other token is refused with kIoXlsbUnsupportedPtg.
Expected<std::vector<FormulaToken>, Error> walk_calc_item_formula(const std::vector<std::uint8_t>& rgce);

}  // namespace pivot_calc
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_CALC_ITEM_RECORDS_H_
