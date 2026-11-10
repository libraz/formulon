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
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "defined_name.h"
#include "eval/recalc_engine.h"
#include "pivot/pivot_layout.h"
#include "sheet.h"
#include "sheet_name.h"
#include "styles.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
class Workbook;
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

// Index of the sheet `name` resolves to (`sheet_names::resolve_by`), or `Workbook::kNoSheet`.
inline std::size_t find_sheet_index(const std::vector<Sheet>& sheets, std::string_view name) noexcept {
  return sheet_names::resolve_by(sheets.size(), [&](std::size_t i, sheet_names::Fold fold) {
    return sheet_names::same(sheets[i].name(), name, fold);
  });
}

// Whether writing `candidate` as the name of the sheet at `self_index` (`Workbook::kNoSheet` for a new sheet)
// collides with another sheet's name (`sheet_names::write_collides`).
inline bool sheet_name_collides(const std::vector<Sheet>& sheets, std::string_view candidate, std::size_t self_index) {
  for (std::size_t i = 0; i < sheets.size(); ++i) {
    if (i != self_index && sheet_names::write_collides(sheets[i].name(), candidate)) {
      return true;
    }
  }
  return false;
}

// The sheets' names, for transforms and formatters that match a formula's sheet spellings as they resolve.
inline sheet_names::NameList sheet_name_list(const std::vector<Sheet>& sheets) {
  std::vector<std::string> names;
  names.reserve(sheets.size());
  for (const Sheet& sheet : sheets) {
    names.emplace_back(sheet.name());
  }
  return sheet_names::NameList(std::move(names));
}

// Sheet index range check whose detail names the sheet count.
inline Expected<void, Error> check_sheet_in_range(std::size_t sheet_index, std::size_t sheet_count) {
  if (sheet_index >= sheet_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet_index is out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheet_count));
  }
  return Expected<void, Error>::Ok();
}

// Target check for a cell setter: the sheet index, then the grid coordinate.
inline Expected<void, Error> check_cell_target(std::size_t sheet_index, std::size_t sheet_count, std::uint32_t row,
                                               std::uint32_t col) {
  RETURN_IF_ERROR(check_sheet_in_range(sheet_index, sheet_count));
  if (!Sheet::coord_in_grid(row, col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "row and col lie outside the grid",
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

/// The defined name `name` resolves to (`sheet_names::resolve_by`) in scope
/// `local_sheet_id` (-1 is workbook scope), or `names.end()`.
inline std::vector<DefinedName>::iterator find_defined_name(std::vector<DefinedName>& names, std::string_view name,
                                                            std::int32_t local_sheet_id) {
  const std::size_t index = sheet_names::resolve_by(names.size(), [&](std::size_t i, sheet_names::Fold fold) {
    return names[i].local_sheet_id == local_sheet_id && sheet_names::same(names[i].name, name, fold);
  });
  return index == sheet_names::kNotFound ? names.end() : names.begin() + static_cast<std::ptrdiff_t>(index);
}

/// Whether writing `name` into workbook scope collides with a workbook-scope
/// name (`sheet_names::write_collides`).
inline bool has_book_namesake(const std::vector<DefinedName>& names, std::string_view name) {
  return std::any_of(names.begin(), names.end(), [name](const DefinedName& entry) {
    return entry.local_sheet_id < 0 && sheet_names::write_collides(entry.name, name);
  });
}

/// The table's A1 rectangle when it lives on `sheet_index`, else nullopt.
inline std::optional<MergeRange> table_rect_on_sheet(const TableMetadata& table, std::size_t sheet_index) {
  return table.sheet_index == sheet_index ? parse_a1_rectangle(table.ref) : std::nullopt;
}

/// The rectangle a pivot occupies: its last projection, else its location
/// with the page fields and their gap row above it. A blocked report holds
/// its anchor alone.
MergeRange pivot_occupied_extent(const pivot::PivotTable& table);

/// A pivot report laid out at its table's current anchor, before anything is written.
struct PivotReportPlan {
  /// False for a cache saved without records, which cannot be re-evaluated; nothing else is set then.
  bool evaluated = false;
  pivot::PivotCells cells;
  /// The report's range, clipped to the grid; nullopt for an empty layout.
  std::optional<MergeRange> extent;
  /// The unclipped range overflows the grid's last row or column.
  bool past_edge = false;
  /// Page-field rows above the anchor inside `extent`.
  std::uint32_t rows_above = 0;
};

/// Evaluates and lays out pivot `pivot_index` on sheet `sheet_index`, publishing the result on the table and
/// changing no cell. Ok with no `extent` for a cache saved without records. Errors: `kInvalidArgument` for an
/// index out of range, `kEvalPivotMissing` for an absent cache, an evaluation or layout error.
Expected<PivotReportPlan, Error> plan_pivot_report(Workbook& wb, std::size_t sheet_index, std::size_t pivot_index);

/// `plan_pivot_report` for `table` as it stands, which need not be on a sheet yet (a moved copy).
Expected<PivotReportPlan, Error> plan_pivot_report(const Workbook& wb, const pivot::PivotTable& table);

/// Something a pivot report cannot be drawn over.
struct ReportOccupant {
  enum class Kind : std::uint8_t { kCell, kSpill, kMerge, kArray, kDataTable, kTable, kPivot };
  Kind kind = Kind::kCell;
  /// First occupied cell of the target, row-major.
  std::uint32_t row = 0;
  std::uint32_t col = 0;
};

/// The first occupant of `target` outside `own` for pivot `pivot_index` on sheet `sheet_index`: a stored cell with
/// a formula or a non-blank value, a cell of a committed spill region, a merge not inside `own`, a cell of a legacy
/// array block, a result cell of a data table, a table, or another pivot's occupied extent. A cell holding only a
/// format is not one. Cost follows the rows of `target`, not the sheet.
std::optional<ReportOccupant> find_report_occupant(const Workbook& wb, std::size_t sheet_index, std::size_t pivot_index,
                                                   const MergeRange& target, const MergeRange& own);

/// Why `target` is blocked for pivot `pivot_index`, over the occupants `find_report_occupant` sees: `kMerged` when
/// the first obstructed cell in row-major order lies in a merge not inside `own`, else `kOccupied`.
SpillCause report_blocker_cause(const Workbook& wb, std::size_t sheet_index, std::size_t pivot_index,
                                const MergeRange& target, const MergeRange& own);

/// Whether pivot `a` (at `a_index` on its sheet) was created before pivot `b` (at `b_index` on the same sheet):
/// the smaller `creation_order`, else the later position in the sheet's list, as Excel lists newer pivots first.
inline bool created_before(const pivot::PivotTable& a, std::size_t a_index, const pivot::PivotTable& b,
                           std::size_t b_index) {
  return a.creation_order() != b.creation_order() ? a.creation_order() < b.creation_order() : a_index > b_index;
}

/// `stem` followed by the smallest N >= `first` whose result collides with none
/// of `names` (`sheet_names::write_collides`).
std::string first_free_numbered_name(std::string_view stem, std::size_t first, const std::vector<std::string>& names);

/// Names a new table's columns: a blank one takes `column_word` plus the
/// smallest free N >= 1, and one colliding with a column to its left takes its
/// own spelling plus the smallest free N >= 2. Every given name, later ones
/// included, stays reserved while numbering.
void name_new_table_columns(std::vector<std::string>& names, std::string_view column_word);

/// Name of a table column inserted right of `left_name` (empty for none):
/// a name ending in ASCII digits counts on (`Q2` -> `Q3`, `Q09` -> `Q10`),
/// falling back to `left_name` plus the smallest free N >= 2 when that is
/// taken; any other takes `column_word` plus the smallest free N >= 1. Taken
/// means colliding with one of `taken` (`sheet_names::write_collides`).
std::string next_table_column_name(std::string_view left_name, std::string_view column_word,
                                   const std::vector<std::string>& taken);

/// Inserts `count` column entries at `offset` in `table.columns`, with fresh
/// ids and names from `next_table_column_name` applied left to right. With a
/// header row, writes each name into its header cell on `rect`, the table's
/// extent after the insert, and marks the cell's dependents dirty. The new
/// data and totals cells are left blank.
void add_table_columns_locked(TableMetadata& table, std::size_t offset, std::uint32_t count, const MergeRange& rect,
                              Sheet& sheet, std::size_t sheet_index, const eval::RecalcEngine::LockedMutator& mutator,
                              std::string_view column_word);

/// Writes the calculated-column formula of each column in table-relative
/// `[first_column, last_column]` into rows `[first_row, last_row]` of `table`
/// on `sheet`, its relative references moved from the first data row to the
/// target row; the array flag stores a one-cell array formula. A formula that
/// does not parse leaves its cells as they are.
void fill_calculated_columns_locked(const Workbook& wb, Sheet& sheet, const eval::RecalcEngine::LockedMutator& mutator,
                                    const TableMetadata& table, std::uint32_t first_row, std::uint32_t last_row,
                                    std::size_t first_column, std::size_t last_column);

/// Moves the relative references of every calculated-column formula of
/// `table` by `row_delta` rows, re-anchoring them on the first data row after
/// rows were inserted at it. A formula that does not parse stays as it is.
void shift_calculated_column_anchor(TableMetadata& table, std::int32_t row_delta);

/// A built-in `ST_TotalsRowFunction` token and its `SUBTOTAL` function number;
/// the 1xx forms skip hidden rows, as the desktop application writes them.
struct TotalsFunction {
  std::string_view token;
  int subtotal_code;
};
inline constexpr TotalsFunction kTotalsFunctions[] = {
    {"sum", 109},   {"min", 105},       {"max", 104},    {"average", 101},
    {"count", 103}, {"countNums", 102}, {"stdDev", 107}, {"var", 110},
};

/// The `=SUBTOTAL(code,Table[Column])` formula the totals row of `table`
/// holds for function number `code` in column `column`.
std::string totals_subtotal_formula(const TableMetadata& table, std::size_t column, int code);

/// The bookkeeping after table `index`'s extent changed: its AutoFilter
/// spans the header through the last data row, the placeholder-row and
/// creation-shift marks clear, formulas naming it are re-indexed, spill
/// barriers re-synced and row-visibility dependents dirtied.
void finish_table_extent_change_locked(Workbook& wb, std::vector<Sheet>& sheets,
                                       const eval::RecalcEngine::LockedMutator& mutator,
                                       const std::vector<DefinedName>& defined_names, std::size_t index);

/// Whether `value`, read from a cell formatted with `xf`, is a number shown
/// as a date or time (elapsed `[h]:mm` included) under `profile`.
bool cell_is_date_value(const StylesTable& styles, const CellXf* xf, ExcelProfile profile, const Value& value);

/// Whether (`row`, `col`) holds a formula, a non-blank value or part of a
/// committed spill: what keeps a cell from becoming a table's new blank row.
bool cell_is_occupied(const Sheet& sheet, std::uint32_t row, std::uint32_t col);

/// Where cell (`row`, `col`) after the shifts of `pending` sits now, or
/// nullopt for a cell one of them inserts.
std::optional<CellAddress> cell_before_shifts(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                              const std::vector<PendingShift>& pending);

/// The edit `insert_cells` (`kDown` / `kRight`) or `delete_cells` (`kUp` /
/// `kLeft`) makes for `range`, after their argument checks.
Expected<StructuralEdit, Error> band_edit_of(const MergeRange& range, CellShift shift);

/// Where `rect` on `sheet_index` lies after the shifts of `pending`, in
/// order, or nullopt once one deletes it. A rectangle straddling a shift's
/// band edge stays put, or is gone with `straddle_drops` (a merge the shift
/// unmerges).
std::optional<MergeRange> rect_after_shifts(const MergeRange& rect, std::size_t sheet_index,
                                            const std::vector<PendingShift>& pending, bool straddle_drops);

/// The structure-intersection preflight every row/column edit passes before
/// anything moves (data tables, legacy arrays, PivotTables, tables, and
/// content an insert pushes off the grid).
Expected<void, Error> check_structure_edit(const Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit);

/// `Workbook::check_band_edit` with the shifts of `pending` taken as applied
/// first. The edit is mapped back onto the current cells, and an insert
/// counts the inserts before it on the same band towards what it pushes off.
Expected<void, Error> check_band_edit_after(const Workbook& wb, std::size_t sheet_index, const MergeRange& range,
                                            CellShift shift, const std::vector<PendingShift>& pending);

/// Refusal of `rect` on `sheet_index` as the range of a table other than
/// `self`, on the first structure Excel keeps a table from overlapping, read
/// where `pending` would move it. With `unmerge_inside`, as on creation, a
/// merge wholly inside is not one.
Expected<void, Error> check_table_range_free(const Workbook& wb, std::size_t sheet_index, const MergeRange& rect,
                                             std::optional<std::size_t> self, bool unmerge_inside,
                                             const std::vector<PendingShift>& pending);

/// Whether a merge, table, PivotTable, multi-cell array formula or data table
/// reaches `edit`'s moving region across its band edge after `pending`, which
/// keeps a totals row toggle from shifting cells.
bool band_meets_straddler(const Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit,
                          const std::vector<PendingShift>& pending);

/// Refusal of the row below `assumed.rect` (within its columns) as a new
/// totals row of table `index` after `pending`, on the first thing occupying
/// it. `assumed.rect` must end above the grid's last row.
Expected<void, Error> check_totals_row_free(const Workbook& wb, std::size_t index, const TableShape& assumed,
                                            const std::vector<PendingShift>& pending);

}  // namespace detail
}  // namespace formulon

#endif  // FORMULON_WORKBOOK_INTERNAL_H_
