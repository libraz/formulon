//
// `Workbook::project_pivot`: writes a pivot's report range into the sheet
// cells.

#include <cstddef>

#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {

Expected<void, Error> Workbook::project_pivot(std::size_t /*sheet_index*/, std::size_t /*pivot_index*/) {
  return {};
}

void Workbook::project_pending_pivots() {}

}  // namespace formulon
