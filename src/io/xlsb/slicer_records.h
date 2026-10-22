//
// Slicers in an XLSB package. The slicer and slicer-cache parts
// (`xl/slicers/slicerN.bin`, `xl/slicerCaches/slicerCacheN.bin`), the sheet's
// TABLESLICERSEX / SLICERSEX future records, the drawing frame and the
// `Slicer_*` defined name all ride the existing retention paths (passthrough
// parts, the sheet tail, the retained drawing, the name table). What the
// workbook part adds is the list of slicer caches, [MS-XLSB] 2.1.7.61
// FRTWORKBOOK: SLICERCACHEIDS (BrtBeginSlicerCacheIDs 1070,
// BrtBeginSlicerCacheID 1072 / BrtEndSlicerCacheID 1073,
// BrtEndSlicerCacheIDs 1071) for pivot slicers and TABLESLICERCACHEIDS
// (BrtTableSlicerCacheIDs 2075, BrtTableSlicerCacheID 2076) for table
// slicers, each ID record an FRTHeader with fRelID naming a workbook
// relationship. Those blocks are kept as `Workbook::xlsb_book_frt()` and
// re-emitted with the relationship ids the writer assigns.
//
// The XML slicer parts an `.xlsx` source carries have no binary form here;
// the writer leaves them out and counts them as deferred.

#ifndef FORMULON_IO_XLSB_SLICER_RECORDS_H_
#define FORMULON_IO_XLSB_SLICER_RECORDS_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/zip_reader.h"

namespace formulon {
namespace io {
namespace xlsb {

/// The framed SLICERCACHEIDS / TABLESLICERCACHEIDS blocks of a
/// `xl/workbook.bin` body, in stream order.
std::vector<std::uint8_t> read_book_slicer_cache_ids(ByteSpan workbook_bin);

/// Splices `retained` (from `read_book_slicer_cache_ids`) into the generated
/// `workbook_bin` ahead of BrtEndBook, in FRTWORKBOOK order, rewriting each
/// relationship id through `id_map` (source id, written id). An ID record
/// whose relationship was not written is dropped, and a block left empty.
void splice_book_slicer_cache_ids(std::vector<std::uint8_t>& workbook_bin, const std::vector<std::uint8_t>& retained,
                                  const std::vector<std::pair<std::string, std::string>>& id_map);

/// True for an `.xlsx` slicer or slicer-cache part (`xl/slicers/*.xml`,
/// `xl/slicerCaches/*.xml`).
bool is_xml_slicer_part(std::string_view path);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SLICER_RECORDS_H_
