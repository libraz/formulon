#include "io/package_reader_common.h"

#include <unordered_set>
#include <utility>

#include "io/dynamic_array_formula.h"
#include "io/sheet_phonetic.h"

namespace formulon {
namespace io {

void mark_sheet_dynamic_arrays(
    Sheet& sheet, std::uint32_t dynamic_array_entry,
    const std::vector<std::tuple<std::uint32_t, std::uint32_t, std::uint32_t>>& cell_metadata) {
  std::unordered_set<std::uint64_t> marked;
  for (const auto& [row, col, entry] : cell_metadata) {
    if (dynamic_array_entry != 0U && entry == dynamic_array_entry) {
      marked.insert(dynamic_array_cell_key(row, col));
    }
  }
  apply_loaded_dynamic_array_marks(sheet, marked);
}

Error sheet_add_error(Error error, std::string_view message, std::string context) {
  if (error.code != FormulonErrorCode::kInvalidSheetName && error.code != FormulonErrorCode::kAlreadyExists) {
    return error;
  }
  return make_error(FormulonErrorCode::kIoSheetCorrupt, std::string(message), std::move(context));
}

void finish_loaded_workbook(Workbook& wb, const ReadOptions& options,
                            const std::vector<std::vector<CellAddress>>& recalc_marked) {
  wb.apply_legacy_implicit_intersections();
  clamp_sheet_phonetic_fonts(wb);
  wb.request_pivot_refresh_on_load();
  if (options.profile.has_value()) {
    wb.assign_loaded_profile(*options.profile);
  }
  wb.apply_loaded_recalc_marks(recalc_marked);
}

}  // namespace io
}  // namespace formulon
