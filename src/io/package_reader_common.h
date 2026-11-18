#ifndef FORMULON_IO_PACKAGE_READER_COMMON_H_
#define FORMULON_IO_PACKAGE_READER_COMMON_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <tuple>
#include <vector>

#include "io/read_options.h"
#include "passthrough_part.h"
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

/// Tags the parts derived from modelled content: the legacy drawing a sheet's `comment_vml_path()` names, the binary
/// index a sheet relationship names, and the part of each external-book link with its relationships part.
void mark_model_derived_parts(const Workbook& wb, std::vector<PassthroughPart>& parts);

/// Completes the cell fonts of the styles just read from a package, under the profile the load assigns (the one in
/// `options`, else the workbook's), before anything fingerprints them.
void complete_loaded_fonts(StylesTable& styles, const Workbook& wb, const ReadOptions& options);

/// Final steps of a package read, once every part is in the workbook: legacy implicit intersections,
/// phonetic fonts, pivot refresh, the requested profile, the scheme font names (the theme is in place by now), the
/// column widths that are only the sheet default, then the recalc marks.
void finish_loaded_workbook(Workbook& wb, const ReadOptions& options,
                            const std::vector<std::vector<CellAddress>>& recalc_marked);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PACKAGE_READER_COMMON_H_
