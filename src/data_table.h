//
// What-If data table (`<f t="dataTable">`): the model a sheet keeps for one
// TABLE() block. The result cells are ordinary literal cells inside `ref`;
// this record is what makes them one protected, recalculated unit.

#ifndef FORMULON_DATA_TABLE_H_
#define FORMULON_DATA_TABLE_H_

#include <cstdint>
#include <string>

#include "merge_range.h"
#include "utils/a1_ref.h"

namespace formulon {

/// One input cell of a data table (`r1` / `r2`), on the table's own sheet.
struct DataTableInput {
  bool present = false;  ///< The attribute was given and parsed.
  bool deleted = false;  ///< `del1` / `del2`: the input cell was deleted.
  std::uint32_t row = 0;
  std::uint32_t col = 0;

  friend bool operator==(const DataTableInput& a, const DataTableInput& b) noexcept {
    return a.present == b.present && a.deleted == b.deleted && a.row == b.row && a.col == b.col;
  }
};

/// Attributes that were spelled out in the source, so a save re-emits them
/// even at their default value.
enum DataTableExplicitAttr : std::uint8_t {
  kDataTableExplicitDt2D = 1U << 0,
  kDataTableExplicitDtr = 1U << 1,
  kDataTableExplicitDel1 = 1U << 2,
  kDataTableExplicitDel2 = 1U << 3,
  kDataTableExplicitCa = 1U << 4,
  kDataTableExplicitAca = 1U << 5,
};

/// True when `(row, col)` lies inside `range`.
inline bool range_contains(const MergeRange& range, std::uint32_t row, std::uint32_t col) noexcept {
  return row >= range.first_row && row <= range.last_row && col >= range.first_col && col <= range.last_col;
}

/// A data table. `ref` is the result range; its top-left cell is the anchor.
///
/// Where the substitution values and formulas sit relative to `ref`:
///   * one-variable, column input (`row_oriented == false`): values in the
///     column left of `ref`, formulas in the row above it; `r1` is the input.
///   * one-variable, row input (`row_oriented == true`): values in the row
///     above `ref`, formulas in the column left of it; `r1` is the input.
///   * two-variable: the formula is the corner cell above-left of `ref`,
///     row-header values lie in the row above, column-header values in the
///     column left; `r1` is the row input cell, `r2` the column input cell.
struct DataTable {
  MergeRange ref;
  bool two_dimensional = false;  ///< `dt2D`.
  bool row_oriented = false;     ///< `dtr`: only a one-variable table reads it.
  DataTableInput r1;
  DataTableInput r2;
  bool always_calc = false;         ///< `ca`.
  bool always_calc_array = false;   ///< `aca`.
  std::uint8_t explicit_attrs = 0;  ///< `DataTableExplicitAttr` bits.

  /// True when `(row, col)` lies inside the result range.
  bool contains(std::uint32_t row, std::uint32_t col) const noexcept { return range_contains(ref, row, col); }

  /// The input fed from the row above `ref`; absent for a column-input table.
  DataTableInput row_input() const noexcept { return two_dimensional || row_oriented ? r1 : DataTableInput{}; }
  /// The input fed from the column left of `ref`; absent for a row-input table.
  DataTableInput col_input() const noexcept { return two_dimensional ? r2 : row_oriented ? DataTableInput{} : r1; }
};

/// How a row/column insert or delete meets a table's block: `ref` plus the
/// header row above it and the header column left of it.
enum class DataTableEditCut : std::uint8_t {
  kMisses,   ///< The edit lies outside the block along its axis; the table shifts.
  kRemoves,  ///< A delete spans the whole block along its axis; the table goes.
  kSplits,   ///< The edit cuts into the block; it is refused.
};

inline DataTableEditCut data_table_edit_cut(const DataTable& table, bool row_axis, bool is_delete, std::uint32_t index,
                                            std::uint32_t count) noexcept {
  const std::uint32_t first = row_axis ? table.ref.first_row : table.ref.first_col;
  const std::uint32_t last = row_axis ? table.ref.last_row : table.ref.last_col;
  const std::uint32_t lo = first == 0U ? 0U : first - 1U;
  if (!is_delete) {
    return index > lo && index <= last ? DataTableEditCut::kSplits : DataTableEditCut::kMisses;
  }
  const std::uint64_t end = static_cast<std::uint64_t>(index) + count;  // exclusive
  if (end <= lo || index > last) {
    return DataTableEditCut::kMisses;
  }
  return index <= lo && end > last ? DataTableEditCut::kRemoves : DataTableEditCut::kSplits;
}

/// The formula every result cell shows: `=TABLE(row_input,col_input)` in A1,
/// an absent input left empty and a deleted one as `#REF!`.
inline std::string data_table_formula_text(const DataTable& table) {
  const auto text = [](const DataTableInput& input) {
    if (input.deleted) {
      return std::string("#REF!");
    }
    return input.present ? a1::encode_a1(input.row, input.col) : std::string();
  };
  return "=TABLE(" + text(table.row_input()) + "," + text(table.col_input()) + ")";
}

}  // namespace formulon

#endif  // FORMULON_DATA_TABLE_H_
