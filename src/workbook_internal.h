//
// Helpers shared by the `Workbook` translation units: grid-target checks,
// the external-book name pre-filter and defined-name lookup. Internal to the
// workbook implementation; not part of the public API.

#ifndef FORMULON_WORKBOOK_INTERNAL_H_
#define FORMULON_WORKBOOK_INTERNAL_H_

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "defined_name.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"

namespace formulon {
namespace detail {

// Hidden sheet-scoped name Excel keeps over a sheet AutoFilter's range.
inline constexpr std::string_view kFilterDatabaseName = "_xlnm._FilterDatabase";

// `op`-prefixed target check for a cell setter: the sheet index, then the grid coordinate.
inline Expected<void, Error> check_cell_target(const char* op, std::size_t sheet_index, std::size_t sheet_count,
                                               std::uint32_t row, std::uint32_t col) {
  if (sheet_index >= sheet_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheet_count));
  }
  if (!Sheet::coord_in_grid(row, col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": coordinate out of grid",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  return Expected<void, Error>::Ok();
}

// Whether `formula` can contain a cross-workbook reference at all: every
// spelling has a `[`, or a `!` after a workbook extension or a quoted path.
inline bool MayNameExternalBook(std::string_view formula) noexcept {
  if (formula.find('[') != std::string_view::npos) {
    return true;
  }
  if (formula.find('!') == std::string_view::npos) {
    return false;
  }
  return strings::case_insensitive_contains(formula, ".xl") ||
         (formula.find('\'') != std::string_view::npos && formula.find_first_of("/\\") != std::string_view::npos);
}

/// The defined name `name' (ASCII case-insensitive) in scope `local_sheet_id`
/// (-1 is workbook scope), or `names.end()`.
inline std::vector<DefinedName>::iterator find_defined_name(std::vector<DefinedName>& names, std::string_view name,
                                                            std::int32_t local_sheet_id) {
  return std::find_if(names.begin(), names.end(), [name, local_sheet_id](const DefinedName& entry) {
    return entry.local_sheet_id == local_sheet_id && strings::case_insensitive_eq(entry.name, name);
  });
}

}  // namespace detail
}  // namespace formulon

#endif  // FORMULON_WORKBOOK_INTERNAL_H_
