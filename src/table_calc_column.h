//
// Calculated columns of tables: Excel's rule for marking a column whose cells
// already hold one formula. The `Workbook` entries that set, clear and
// re-infer marks (`set_table_calculated_column` and its siblings) are
// defined in table_calc_column.cpp.

#ifndef FORMULON_TABLE_CALC_COLUMN_H_
#define FORMULON_TABLE_CALC_COLUMN_H_

#include <cstddef>
#include <optional>
#include <string>

namespace formulon {

class Sheet;
struct TableMetadata;

/// The calculated-column mark Excel gives column `column` of `table` on
/// `sheet` when the table is created or resized over it. With W the smaller
/// of 5 and the data-row count, the last data row must hold a formula F that
/// is not an array formula, and at least W - 1 of the last W data rows must
/// hold F in R1C1 terms (compared as text after moving each to the last
/// row). Returns F moved to the first data row, without `=`; nullopt when the
/// rule fails, the table has no data row, `column` is out of range, or the
/// move would take a reference off the grid.
std::optional<std::string> infer_calculated_column(const Sheet& sheet, const TableMetadata& table, std::size_t column);

}  // namespace formulon

#endif  // FORMULON_TABLE_CALC_COLUMN_H_
