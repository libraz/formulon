//
// Checks and inputs shared by the OOXML and XLSB package writers.

#ifndef FORMULON_IO_WRITER_COMMON_H_
#define FORMULON_IO_WRITER_COMMON_H_

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Fails with `kIoWriteFailed` when a table names a sheet the workbook does
/// not have; `context` is the writer's error-context prefix.
inline Expected<void, Error> check_tables_have_sheets(const Workbook& wb, std::string_view context) {
  const std::size_t sheet_count = wb.sheet_count();
  for (const TableMetadata& table : wb.tables()) {
    if (table.sheet_index >= sheet_count) {
      return make_error(FormulonErrorCode::kIoWriteFailed, "table references a missing worksheet",
                        std::string(context) + " table=" + table.name + " sheet_index=" +
                            std::to_string(table.sheet_index) + " sheet_count=" + std::to_string(sheet_count));
    }
  }
  return Expected<void, Error>::Ok();
}

/// The sheet names of `wb` in sheet order.
inline std::vector<std::string> sheet_names_of(const Workbook& wb) {
  std::vector<std::string> names;
  names.reserve(wb.sheet_count());
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    names.push_back(wb.sheet(i).name());
  }
  return names;
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_WRITER_COMMON_H_
