// Which sheet tabs a save marks selected. Excel never saves a book without a selected tab: when no sheet's view
// carries `tabSelected`, the active sheet's tab is written selected, as `Workbook::apply_open_active_sheet_rule`
// selects it when the book is opened, so a first save and the save after a reload agree.

#pragma once

#include <cstddef>

#include "workbook.h"

namespace formulon {
namespace io {

/// Whether a save selects sheet `index`'s tab although its view does not: no sheet of `wb` is selected and `index`
/// is the active sheet.
inline bool tab_selected_by_save(const Workbook& wb, std::size_t index) {
  if (index != wb.active_sheet()) {
    return false;
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).view().tab_selected) {
      return false;
    }
  }
  return true;
}

}  // namespace io
}  // namespace formulon
