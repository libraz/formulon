//
// The workbook part's list of PivotCaches, [MS-XLSB] PIVOTCACHEIDS:
// BrtBeginPivotCacheIDs (384), one BrtBeginPivotCacheID (386, u32 idSx and an
// XLWideString naming the workbook relationship of the cache definition) /
// BrtEndPivotCacheID (387) pair per cache, BrtEndPivotCacheIDs (385). Excel
// refuses a package whose pivot tables name a cache this list lacks. An
// `.xlsb` source's block is kept as `Workbook::xlsb_pivot_cache_ids()` and
// re-emitted with the relationship ids the writer assigns.

#ifndef FORMULON_IO_XLSB_PIVOT_CACHE_IDS_H_
#define FORMULON_IO_XLSB_PIVOT_CACHE_IDS_H_

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "io/zip_reader.h"

namespace formulon {
namespace io {
namespace xlsb {

/// The framed PIVOTCACHEIDS block of a `xl/workbook.bin` body; empty when the
/// body has none.
std::vector<std::uint8_t> read_book_pivot_cache_ids(ByteSpan workbook_bin);

/// Splices `retained` (from `read_book_pivot_cache_ids`) into the generated
/// `workbook_bin` ahead of BrtFileRecover, the future-record blocks and
/// BrtEndBook, rewriting each relationship id through `id_map` (source id,
/// written id). A cache whose relationship was not written is dropped, and the
/// block with it when none is left.
void splice_book_pivot_cache_ids(std::vector<std::uint8_t>& workbook_bin, const std::vector<std::uint8_t>& retained,
                                 const std::vector<std::pair<std::string, std::string>>& id_map);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_CACHE_IDS_H_
