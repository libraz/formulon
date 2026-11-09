//
// Helpers shared by the `Workbook` translation units: grid-target checks,
// the external-book name pre-filter and defined-name lookup. Internal to the
// workbook implementation; not part of the public API.

#ifndef FORMULON_WORKBOOK_INTERNAL_H_
#define FORMULON_WORKBOOK_INTERNAL_H_

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "auto_filter.h"
#include "defined_name.h"
#include "eval/recalc_engine.h"
#include "sheet.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"

namespace formulon {
namespace pivot {
class PivotTable;
}  // namespace pivot

namespace detail {

// Hidden sheet-scoped name Excel keeps over a sheet AutoFilter's range.
inline constexpr std::string_view kFilterDatabaseName = "_xlnm._FilterDatabase";

// Eagerly marks every existing dependent of `cell` dirty in `mutator`'s
// engine. The next `recalc()` pass would discover them via BFS anyway,
// but eager marking keeps the dirty set self-consistent between
// mutations and matches what callers see when they introspect the
// engine via `recalc_engine()`.
//
// Precondition: caller holds the engine mutex for the lifetime of the
// `LockedMutator&`. The helper routes through the facade so the
// compound mutation in `set_cell_value` / `set_cell_formula` stays
// under a single critical section, which keeps the `Sheet` write that
// follows from racing against a concurrent `recalc_parallel`.
inline void mark_dependents_dirty(const eval::RecalcEngine::LockedMutator& mutator, eval::CellNodeId cell) {
  for (eval::CellNodeId dep : mutator.dep_graph().dependents_of_ref(cell)) {
    mutator.mark_dirty(dep);
  }
  mutator.mark_range_dependents_dirty(cell);
}

// Index of the sheet named `name` (Unicode simple fold), or `Workbook::kNoSheet`.
inline std::size_t find_sheet_index(const std::vector<Sheet>& sheets, std::string_view name) noexcept {
  for (std::size_t i = 0; i < sheets.size(); ++i) {
    if (sheet_names::equal(sheets[i].name(), name)) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}

// `op`-prefixed sheet index range check whose detail names the sheet count.
inline Expected<void, Error> check_sheet_in_range(const char* op, std::size_t sheet_index, std::size_t sheet_count) {
  if (sheet_index >= sheet_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheet_count));
  }
  return Expected<void, Error>::Ok();
}

// `op`-prefixed target check for a cell setter: the sheet index, then the grid coordinate.
inline Expected<void, Error> check_cell_target(const char* op, std::size_t sheet_index, std::size_t sheet_count,
                                               std::uint32_t row, std::uint32_t col) {
  RETURN_IF_ERROR(check_sheet_in_range(op, sheet_index, sheet_count));
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

/// The defined name `name' (`collection_names::equal`) in scope `local_sheet_id`
/// (-1 is workbook scope), or `names.end()`.
inline std::vector<DefinedName>::iterator find_defined_name(std::vector<DefinedName>& names, std::string_view name,
                                                            std::int32_t local_sheet_id) {
  return std::find_if(names.begin(), names.end(), [name, local_sheet_id](const DefinedName& entry) {
    return entry.local_sheet_id == local_sheet_id && collection_names::equal(entry.name, name);
  });
}

/// Whether a workbook-scope name spelled `name` (`collection_names::equal`) exists.
inline bool has_book_namesake(const std::vector<DefinedName>& names, std::string_view name) {
  return std::any_of(names.begin(), names.end(), [name](const DefinedName& entry) {
    return entry.local_sheet_id < 0 && collection_names::equal(entry.name, name);
  });
}

/// The table's A1 rectangle when it lives on `sheet_index`, else nullopt.
inline std::optional<MergeRange> table_rect_on_sheet(const TableMetadata& table, std::size_t sheet_index) {
  return table.sheet_index == sheet_index ? parse_a1_rectangle(table.ref) : std::nullopt;
}

/// The rectangle a pivot occupies: its last projection, else its location
/// with the page fields and their gap row above it.
MergeRange pivot_occupied_extent(const pivot::PivotTable& table);

}  // namespace detail
}  // namespace formulon

#endif  // FORMULON_WORKBOOK_INTERNAL_H_
