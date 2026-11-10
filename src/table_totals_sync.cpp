// Totals-row cell classification and its application to column metadata.
//
// @size-budget: 2 KB

#include "table_totals_sync.h"

#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "auto_filter.h"
#include "eval/recalc_engine.h"
#include "parser/ast_format.h"
#include "parser/structured_ref_payload.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/number_text.h"
#include "value.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"

namespace formulon {

namespace {

std::string_view without_equals(std::string_view text) {
  return !text.empty() && text.front() == '=' ? text.substr(1) : text;
}

/// Parsed and re-formatted `text`; empty when it does not parse.
std::string canonical_formula(std::string_view text, Arena& arena) {
  const parser::AstNode* root = parse_stored_formula(text, arena);
  return root == nullptr ? std::string() : parser::format_formula(*root);
}

std::string value_label(const Value& value) {
  if (value.is_text()) {
    return std::string(value.as_text());
  }
  if (value.is_number()) {
    std::string out;
    append_xml_number(out, value.as_number());
    return out;
  }
  if (value.is_boolean()) {
    return value.as_boolean() ? "TRUE" : "FALSE";
  }
  if (value.is_error()) {
    return error_info(value.as_error()).display_name;
  }
  return {};
}

}  // namespace

TotalsCellClass classify_totals_cell(const Workbook& /*wb*/, const TableMetadata& table, std::size_t column,
                                     const Cell& cell) {
  TotalsCellClass result;
  if (cell.formula_text.empty()) {
    result.text = value_label(cell.cached_value);
    result.kind = result.text.empty() ? TotalsCellClass::Kind::kNone : TotalsCellClass::Kind::kLabel;
    return result;
  }
  if (column < table.columns.size() && cell.array_block_rows == 0U) {
    Arena cell_arena;
    const std::string actual = canonical_formula(cell.formula_text, cell_arena);
    if (!actual.empty()) {
      Arena expected_arena;
      for (const detail::TotalsFunction& entry : detail::kTotalsFunctions) {
        const std::string expected = Workbook::normalize_formula_text(
            detail::totals_subtotal_formula(table, column, entry.subtotal_code), Workbook::FormulaSource::kEntered);
        if (canonical_formula(expected, expected_arena) == actual) {
          result.kind = TotalsCellClass::Kind::kFunction;
          result.text = std::string(entry.token);
          return result;
        }
      }
    }
  }
  result.kind = TotalsCellClass::Kind::kCustom;
  result.text = std::string(without_equals(cell.formula_text));
  result.array = cell.array_block_rows != 0U;
  return result;
}

void apply_totals_cell_locked(Workbook& wb, const eval::RecalcEngine::LockedMutator& mutator, std::size_t table_index,
                              std::size_t column) {
  if (table_index >= wb.tables().size()) {
    return;
  }
  const TableMetadata& table = wb.tables()[table_index];
  if (!table.totals_row || column >= table.columns.size()) {
    return;
  }
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (!rect.has_value()) {
    return;
  }
  const std::size_t sheet_index = table.sheet_index;
  const std::uint32_t row = rect->last_row;
  const std::uint32_t col = rect->first_col + static_cast<std::uint32_t>(column);
  TotalsCellClass cls;
  bool rewrite_as_text = false;
  if (const Cell* cell = wb.sheet(sheet_index).cell_at(row, col); cell != nullptr) {
    cls = classify_totals_cell(wb, table, column, *cell);
    rewrite_as_text = cls.kind == TotalsCellClass::Kind::kLabel && !cell->cached_value.is_text();
  }

  TableColumn& meta = wb.mutable_tables()[table_index].columns[column];
  meta.totals_function.clear();
  meta.totals_label.clear();
  meta.totals_row_formula.clear();
  meta.totals_row_formula_array = false;
  switch (cls.kind) {
    case TotalsCellClass::Kind::kNone:
      break;
    case TotalsCellClass::Kind::kLabel:
      meta.totals_label = cls.text;
      break;
    case TotalsCellClass::Kind::kFunction:
      meta.totals_function = cls.text;
      break;
    case TotalsCellClass::Kind::kCustom:
      meta.totals_function = "custom";
      meta.totals_row_formula = cls.text;
      meta.totals_row_formula_array = cls.array;
      break;
  }
  if (rewrite_as_text) {
    wb.sheet(sheet_index).set_cell_text(row, col, cls.text);
    detail::mark_dependents_dirty(mutator, make_node(sheet_index, row, col));
  }
}

void normalize_table_totals_cells_locked(Workbook& wb, const eval::RecalcEngine::LockedMutator& mutator) {
  for (std::size_t i = 0; i < wb.tables().size(); ++i) {
    const TableMetadata& table = wb.tables()[i];
    const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
    if (!table.totals_row || !rect.has_value()) {
      continue;
    }
    for (std::size_t column = 0; column < table.columns.size(); ++column) {
      // A file that stores no cell there keeps the settings it declares.
      const std::uint32_t col = rect->first_col + static_cast<std::uint32_t>(column);
      if (wb.sheet(table.sheet_index).cell_at(rect->last_row, col) != nullptr) {
        apply_totals_cell_locked(wb, mutator, i, column);
      }
    }
  }
}

void Workbook::sync_table_totals_cell(std::size_t table_index, std::size_t column) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  apply_totals_cell_locked(*this, mutator, table_index, column);
}

void Workbook::normalize_table_totals_cells() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  normalize_table_totals_cells_locked(*this, mutator);
}

}  // namespace formulon
