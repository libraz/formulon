#ifndef FORMULON_IO_SHEET_IDS_H_
#define FORMULON_IO_SHEET_IDS_H_

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <unordered_set>
#include <vector>

#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// The sheet id each sheet is written with: its source id when that is
/// non-zero and no earlier sheet holds it, else the next id past every one in
/// use, so a retained part citing a sheet by `tabId` still finds it.
inline std::vector<std::uint32_t> assign_sheet_ids(const Workbook& wb) {
  std::vector<std::uint32_t> ids(wb.sheet_count(), 0U);
  std::unordered_set<std::uint32_t> used;
  std::uint32_t next = 0;
  for (std::size_t i = 0; i < ids.size(); ++i) {
    const std::uint32_t source = wb.sheet(i).source_sheet_id();
    if (source != 0U && used.insert(source).second) {
      ids[i] = source;
      next = std::max(next, source);
    }
  }
  for (std::uint32_t& id : ids) {
    while (id == 0U) {
      next = next == std::numeric_limits<std::uint32_t>::max() ? 1U : next + 1U;
      if (used.insert(next).second) {
        id = next;
      }
    }
  }
  return ids;
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SHEET_IDS_H_
