// Row and column line items of an XLSB pivot table part: the SXLI block
// (BrtBeginSXLIRws / BrtBeginSXLICols) and its BrtBeginSXLI rows, each with
// an optional BrtBeginISXVIs index array. One row is one `<i>` of
// `<rowItems>` / `<colItems>`.
//
// BrtBeginSXLI payload (12 bytes): u16 r (repeated leading members), u8 t
// (ST_ItemType order: data, default, sum, countA, avg, max, min, product,
// count, stdDev, stdDevP, var, varP, grand, blank), u8 reserved (0), u32
// count of `x`, u32 i (data-field index). BrtBeginISXVIs follows iff the
// count is non-zero and holds that many u32 item indices.

#ifndef FORMULON_IO_XLSB_PIVOT_LINE_ITEM_RECORDS_H_
#define FORMULON_IO_XLSB_PIVOT_LINE_ITEM_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/xlsb/pivot_records_common.h"
#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Largest `t` value a line item may carry (`blank`).
inline constexpr std::uint8_t kSxLineItemTypeMax = 14;
/// `t` of a grand-total row.
inline constexpr std::uint8_t kSxLineItemTypeGrand = 13;

/// Which axis a line-item block belongs to.
enum class SxLineItemAxis : std::uint8_t { kRows, kCols };

/// One `<i>`: repeat count, item type, data-field index and member indices.
struct SxLineItem {
  std::uint16_t r = 0;
  std::uint8_t t = 0;
  std::uint32_t i = 0;
  std::vector<std::uint32_t> x;
};

/// The rows of one `<rowItems>` / `<colItems>` block, in stream order.
struct SxLineItems {
  std::vector<SxLineItem> items;
};

/// Decodes the line-item block whose begin record is `records[*pos]`
/// (BrtBeginSXLIRws or BrtBeginSXLICols, matching `axis`) through its end
/// record; on success `*pos` indexes the record after the end record.
///
/// Errors: `kIoXlsbUnknownRecord` for a record the block does not admit,
/// `kIoXlsbRecordCorrupt` for a malformed payload or a count mismatch.
Expected<SxLineItems, Error> decode_line_items(const std::vector<XlsbRecord>& records, std::size_t* pos,
                                               SxLineItemAxis axis);

/// Appends the framed line-item block for `axis` to `dst`.
void encode_line_items(std::vector<std::uint8_t>& dst, const SxLineItems& items, SxLineItemAxis axis);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_LINE_ITEM_RECORDS_H_
