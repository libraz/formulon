//
// Calculated columns of tables. A column is marked by Excel's window rule
// when a table is created or resized over formulas, or explicitly through
// `Workbook::set_table_calculated_column`, which fills the column and may
// register a shared-formula group the `.xlsx` writer verifies at save. "Same
// formula" is R1C1 identity, judged as the formatted text of both formulas
// moved to one row, the mechanism the calculated-column fill uses.

// @size-budget: 6 KB

#include "table_calc_column.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "eval/builtin_names.h"
#include "eval/recalc_engine.h"
#include "io/dynamic_array_formula.h"
#include "io/xlsb/ptg_writer.h"
#include "merge_range.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/external_book_spelling.h"
#include "parser/parse_status.h"
#include "sheet.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/status_macros.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"

namespace formulon {

namespace {

Error invalid_argument(std::string message, std::string context = {}) {
  Error err = make_error(FormulonErrorCode::kInvalidArgument, std::move(message));
  err.context = std::move(context);
  return err;
}

/// A table's data rows: `count` rows from `first`.
struct DataRows {
  std::uint32_t first = 0;
  std::uint32_t count = 0;

  std::uint32_t last() const noexcept { return first + count - 1U; }
};

DataRows data_rows_of(const TableMetadata& table, const MergeRange& rect) {
  const std::uint32_t first = rect.first_row + (table.header_row ? 1U : 0U);
  const std::uint32_t end = rect.last_row + 1U - (table.totals_row ? 1U : 0U);
  return DataRows{first, end > first ? end - first : 0U};
}

/// Whether `cell` holds a formula that is neither a legacy array block nor a
/// dynamic-array formula.
bool holds_plain_formula(const Cell* cell) {
  return cell != nullptr && !cell->formula_text.empty() && cell->array_block_rows == 0U && !cell->dynamic_array;
}

bool is_blank_cell(const Cell* cell) {
  return cell == nullptr || (cell->formula_text.empty() && cell->cached_value.is_blank());
}

/// `root` moved by `row_delta` rows and formatted, or nullopt when the move
/// collapses a reference off the grid (moving back does not restore `root`).
/// `arena` is scratch, reset here.
std::optional<std::string> moved_text(const parser::AstNode& root, std::int32_t row_delta, Arena& arena) {
  arena.reset();
  const parser::AstNode* moved = parser::shift_relative_refs(root, arena, row_delta, 0);
  if (moved == nullptr) {
    return std::nullopt;
  }
  const parser::AstNode* back = parser::shift_relative_refs(*moved, arena, -row_delta, 0);
  if (back == nullptr || parser::format_formula(*back) != parser::format_formula(root)) {
    return std::nullopt;
  }
  return parser::format_formula(*moved);
}

bool contains_structured_ref(const parser::AstNode& node) {
  return node.kind() == parser::NodeKind::StructuredRef ||
         parser::any_child_node(node, [](const parser::AstNode& child) { return contains_structured_ref(child); });
}

/// The range of table `table_index`, which must be a table of `sheet_index`.
Expected<MergeRange, Error> table_rect(const Workbook& wb, std::size_t sheet_index, std::size_t table_index) {
  if (sheet_index >= wb.sheet_count()) {
    return invalid_argument("sheet index out of range", "sheet=" + std::to_string(sheet_index));
  }
  if (table_index >= wb.tables().size() || wb.tables()[table_index].sheet_index != sheet_index) {
    return invalid_argument("table index out of range", "table=" + std::to_string(table_index));
  }
  const TableMetadata& table = wb.tables()[table_index];
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (!rect.has_value()) {
    return invalid_argument("table ref is not a range", "ref=" + table.ref);
  }
  return *rect;
}

/// `table_rect`, refusing a `column` the table does not have.
Expected<MergeRange, Error> table_column_rect(const Workbook& wb, std::size_t sheet_index, std::size_t table_index,
                                              std::size_t column) {
  ASSIGN_OR_RETURN(const MergeRange rect, table_rect(wb, sheet_index, table_index));
  const std::size_t columns = wb.tables()[table_index].columns.size();
  if (column >= columns) {
    return invalid_argument("column index out of range",
                            "column=" + std::to_string(column) + " column_count=" + std::to_string(columns));
  }
  return rect;
}

}  // namespace

std::optional<std::string> infer_calculated_column(const Sheet& sheet, const TableMetadata& table, std::size_t column) {
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (!rect.has_value() || column >= table.columns.size() ||
      column > static_cast<std::size_t>(rect->last_col - rect->first_col)) {
    return std::nullopt;
  }
  const DataRows rows = data_rows_of(table, *rect);
  if (rows.count == 0U) {
    return std::nullopt;
  }
  const std::uint32_t col = rect->first_col + static_cast<std::uint32_t>(column);
  const std::uint32_t last = rows.last();
  const Cell* last_cell = sheet.cell_at(last, col);
  if (!holds_plain_formula(last_cell)) {
    return std::nullopt;
  }
  Arena last_arena;
  const parser::AstNode* formula = parse_stored_formula(last_cell->formula_text, last_arena);
  if (formula == nullptr) {
    return std::nullopt;
  }
  const std::string target = parser::format_formula(*formula);
  const std::uint32_t window = std::min<std::uint32_t>(5U, rows.count);
  Arena row_arena;
  Arena shift_arena;
  std::uint32_t matches = 1U;
  for (std::uint32_t row = last - window + 1U; row < last; ++row) {
    const Cell* cell = sheet.cell_at(row, col);
    if (!holds_plain_formula(cell)) {
      continue;
    }
    const parser::AstNode* root = parse_stored_formula(cell->formula_text, row_arena);
    if (root == nullptr) {
      continue;
    }
    const std::optional<std::string> moved = moved_text(*root, static_cast<std::int32_t>(last - row), shift_arena);
    if (moved.has_value() && *moved == target) {
      ++matches;
    }
  }
  if (matches + 1U < window) {
    return std::nullopt;
  }
  return moved_text(*formula, -static_cast<std::int32_t>(last - rows.first), shift_arena);
}

Expected<void, Error> Workbook::set_table_calculated_column(std::size_t sheet_index, std::size_t table_index,
                                                            std::size_t column, std::string_view formula) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  ASSIGN_OR_RETURN(const MergeRange rect, table_column_rect(*this, sheet_index, table_index, column));
  const std::string_view body = strip_formula_prefix(formula);
  if (body.empty()) {
    return invalid_argument("formula is empty");
  }
  const DataRows rows = data_rows_of(tables_[table_index], rect);
  const std::uint32_t col = rect.first_col + static_cast<std::uint32_t>(column);
  // Bound inside the table, so an unqualified `[@x]` names it even with no data row.
  const std::uint32_t bind_row = rows.count > 0U ? rows.first : rect.first_row;
  std::string text = normalize_formula_text(
      bind_entered_structured_refs("=" + std::string(body), sheet_index, bind_row, col), FormulaSource::kEntered);
  if (!file_name_.empty()) {
    text = parser::bind_self_book_names(text, file_name_);
  }
  Arena arena;
  parser::ParseFailure failure;
  const parser::AstNode* root = eval::parse_formula_entry_with_diagnostics(strip_formula_prefix(text), arena, &failure);
  if (root == nullptr) {
    return make_error(parser::parse_error_status(failure.code), "formula does not parse", std::string(formula));
  }
  if (io::entered_as_dynamic_array(*this, sheet_index, *root)) {
    return invalid_argument("formula must not be a dynamic-array formula", std::string(formula));
  }
  Arena shift_arena;
  if (rows.count > 1U && !moved_text(*root, static_cast<std::int32_t>(rows.count - 1U), shift_arena).has_value()) {
    return invalid_argument("formula moves a reference off the grid", std::string(formula));
  }

  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  if (detail::MayNameExternalBook(text)) {
    bind_external_books_unlocked(*root);
  }
  TableMetadata& table = tables_[table_index];
  table.columns[column].calculated_column_formula = parser::format_formula(*root);
  table.columns[column].calculated_column_formula_array = false;
  if (rows.count > 0U) {
    const MergeRange cells{rows.first, col, rows.last(), col};
    Sheet& sheet = sheets_[sheet_index];
    // A group over these cells no longer describes what they hold.
    std::vector<SharedFormulaGroup>& groups = sheet.mutable_shared_formula_groups();
    groups.erase(std::remove_if(groups.begin(), groups.end(),
                                [&cells](const SharedFormulaGroup& group) { return intersects(group.ref, cells); }),
                 groups.end());
    detail::fill_calculated_columns_locked(*this, sheet, mutator, table, rows.first, rows.last(), column, column);
    if (!contains_structured_ref(*root) &&
        !io::xlsb::formula_always_calculates(*root, io::name_shapes(*this, sheet_index))) {
      groups.push_back(SharedFormulaGroup{cells});
    }
  }
  reindex_formulas_referencing_table(sheets_, mutator, *this, {table.name, table.display_name});
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::clear_table_calculated_column(std::size_t sheet_index, std::size_t table_index,
                                                              std::size_t column) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(table_column_rect(*this, sheet_index, table_index, column));
  TableMetadata& table = tables_[table_index];
  table.columns[column].calculated_column_formula.clear();
  table.columns[column].calculated_column_formula_array = false;
  reindex_formulas_referencing_table(sheets_, engine_->locked_mutator(), *this, {table.name, table.display_name});
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::drop_calculated_column_if_empty(std::size_t sheet_index, std::size_t table_index,
                                                                std::size_t column) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  ASSIGN_OR_RETURN(const MergeRange rect, table_column_rect(*this, sheet_index, table_index, column));
  TableMetadata& table = tables_[table_index];
  const DataRows rows = data_rows_of(table, rect);
  if (rows.count == 0U || table.columns[column].calculated_column_formula.empty()) {
    return Expected<void, Error>::Ok();
  }
  const Sheet& sheet = sheets_[sheet_index];
  const std::uint32_t col = rect.first_col + static_cast<std::uint32_t>(column);
  for (std::uint32_t row = rows.first; row <= rows.last(); ++row) {
    if (!is_blank_cell(sheet.cell_at(row, col))) {
      return Expected<void, Error>::Ok();
    }
  }
  table.columns[column].calculated_column_formula.clear();
  table.columns[column].calculated_column_formula_array = false;
  reindex_formulas_referencing_table(sheets_, engine_->locked_mutator(), *this, {table.name, table.display_name});
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::infer_table_calculated_columns(std::size_t sheet_index, std::size_t table_index,
                                                               std::uint32_t first_new_row) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  ASSIGN_OR_RETURN(const MergeRange rect, table_rect(*this, sheet_index, table_index));
  TableMetadata& table = tables_[table_index];
  const DataRows rows = data_rows_of(table, rect);
  Sheet& sheet = sheets_[sheet_index];
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  for (std::size_t column = 0; column < table.columns.size(); ++column) {
    TableColumn& entry = table.columns[column];
    if (entry.calculated_column_formula.empty()) {
      if (std::optional<std::string> inferred = infer_calculated_column(sheet, table, column)) {
        entry.calculated_column_formula = std::move(*inferred);
        entry.calculated_column_formula_array = false;
      }
      continue;
    }
    // Each run of blank cells among the added data rows takes the column's formula.
    const std::uint32_t col = rect.first_col + static_cast<std::uint32_t>(column);
    for (std::uint32_t row = std::max(rows.first, first_new_row); rows.count > 0U && row <= rows.last(); ++row) {
      if (!is_blank_cell(sheet.cell_at(row, col))) {
        continue;
      }
      std::uint32_t run_last = row;
      while (run_last < rows.last() && is_blank_cell(sheet.cell_at(run_last + 1U, col))) {
        ++run_last;
      }
      detail::fill_calculated_columns_locked(*this, sheet, mutator, table, row, run_last, column, column);
      row = run_last;
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
