#ifndef FORMULON_IO_PACKAGE_READER_COMMON_H_
#define FORMULON_IO_PACKAGE_READER_COMMON_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <tuple>
#include <vector>

#include "io/read_options.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Marks as dynamic the formulas of `sheet` whose `(row, col, entry)` in `cell_metadata` names
/// `dynamic_array_entry`; 0 marks none.
void mark_sheet_dynamic_arrays(
    Sheet& sheet, std::uint32_t dynamic_array_entry,
    const std::vector<std::tuple<std::uint32_t, std::uint32_t, std::uint32_t>>& cell_metadata);

/// The error a failed `Workbook::add_sheet_validated` surfaces: an invalid or colliding name becomes
/// `kIoSheetCorrupt` with `message` and `context`, any other error passes through.
Error sheet_add_error(Error error, std::string_view message, std::string context);

/// Final steps of a package read, once every part is in the workbook: legacy implicit intersections,
/// phonetic fonts, pivot refresh, the requested profile, then the recalc marks.
void finish_loaded_workbook(Workbook& wb, const ReadOptions& options,
                            const std::vector<std::vector<CellAddress>>& recalc_marked);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PACKAGE_READER_COMMON_H_
