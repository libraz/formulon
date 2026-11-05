//
// Internal seam between the table part codec and the worksheet-tail LISTPARTS
// handling (`table_list_parts.cpp`). Not part of the public API.

#ifndef FORMULON_IO_XLSB_TABLE_RECORDS_INTERNAL_H_
#define FORMULON_IO_XLSB_TABLE_RECORDS_INTERNAL_H_

#include <cstddef>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/record.h"
#include "sheet.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace detail {

/// The LISTPARTS block of a tail slot: record index range and its ids.
struct ListPartsBlock {
  std::size_t begin = 0;  // index of BrtBeginListParts
  std::size_t end = 0;    // index of BrtEndListParts
  std::vector<std::string> rel_ids;
};

/// The LISTPARTS block among `recs`, or nullopt when there is none or it is
/// malformed.
std::optional<ListPartsBlock> FindListParts(const std::vector<FramedRecord>& recs);

/// Removes the BrtListPart records naming `rel_ids` from the sheet (see the
/// definition).
void ReleaseSheetTables(Sheet& sheet, const std::vector<std::string>& rel_ids);

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_TABLE_RECORDS_INTERNAL_H_
