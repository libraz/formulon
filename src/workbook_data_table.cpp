//
// What-If data tables on `Workbook`: creating, removing and querying them,
// and the calculation that fills their result cells.
//
// A result cell is found by writing the substitution value into the input
// cell, recalculating through the engine (which re-evaluates only the cells
// that read the input), and reading the table's formula cell. The inputs are
// restored afterwards and one more recalc returns their dependents to their
// real values.

#include <cstddef>
#include <cstdint>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {

namespace {

// A cell value that owns its text, so it survives the cell it was read from
// being rewritten.
struct OwnedValue {
  Value value = Value::blank();
  std::string text;
  bool is_text = false;

  static OwnedValue from(const Value& v) {
    OwnedValue out;
    if (v.is_text()) {
      out.is_text = true;
      out.text = std::string(v.as_text());
    } else if (v.is_array() || v.is_ref() || v.is_lambda()) {
      out.value = Value::error(ErrorCode::Value);
    } else {
      out.value = v;
    }
    return out;
  }

  static OwnedValue error(ErrorCode code) { return from(Value::error(code)); }

  // A formula cell that is empty reads as 0 in a data table.
  static OwnedValue result_of(const Cell* cell) {
    if (cell == nullptr || cell->cached_value.is_blank()) {
      return from(Value::number(0.0));
    }
    return from(cell->cached_value);
  }
};

bool is_literal_scalar_slot(const Cell* cell) {
  return cell == nullptr || (!cell->dynamic_array && cell->array_block_rows == 0U);
}

// Whether any cell of `ref` is part of a legacy array block or a spill.
bool intersects_array_or_spill(const Sheet& sheet, const MergeRange& ref) {
  for (std::uint32_t row = ref.first_row; row <= ref.last_row; ++row) {
    for (std::uint32_t col = ref.first_col; col <= ref.last_col; ++col) {
      if (!is_literal_scalar_slot(sheet.cell_at(row, col)) || sheet.array_block_covering(row, col).has_value() ||
          sheet.committed_spill_footprint_covering(row, col).has_value()) {
        return true;
      }
    }
  }
  return false;
}

// One input cell of a table, swapped to substitution values and put back.
class InputCell {
 public:
  InputCell(Workbook& wb, std::size_t sheet_index, const DataTableInput& input)
      : wb_(wb), sheet_(sheet_index), row_(input.row), col_(input.col) {
    const Cell* cell = wb.sheet(sheet_index).cell_at(row_, col_);
    present_ = cell != nullptr;
    if (cell != nullptr) {
      formula_ = cell->formula_text;
      dynamic_ = cell->dynamic_array;
      original_ = OwnedValue::from(cell->cached_value);
    }
  }

  Expected<void, Error> set(const OwnedValue& v) { return write(v); }

  Expected<void, Error> restore() {
    if (formula_.empty()) {
      RETURN_IF_ERROR(write(original_));
      if (!present_) {
        wb_.sheet(sheet_).erase_cell(row_, col_);
      }
      return Expected<void, Error>::Ok();
    }
    RETURN_IF_ERROR(wb_.set_cell_formula(sheet_, row_, col_, formula_, Workbook::FormulaSource::kRendered));
    wb_.sheet(sheet_).set_cell_dynamic_array(row_, col_, dynamic_);
    return Expected<void, Error>::Ok();
  }

 private:
  Expected<void, Error> write(const OwnedValue& v) {
    return v.is_text ? wb_.set_cell_text(sheet_, row_, col_, v.text) : wb_.set_cell_value(sheet_, row_, col_, v.value);
  }

  Workbook& wb_;
  std::size_t sheet_;
  std::uint32_t row_;
  std::uint32_t col_;
  std::string formula_;
  bool dynamic_ = false;
  bool present_ = false;
  OwnedValue original_;
};

// Whether `input` can take substitution values: present, not deleted, on the
// grid, outside every data table's results, and not part of an array or spill.
bool input_usable(const Sheet& sheet, const DataTableInput& input) {
  return input.present && !input.deleted && Sheet::coord_in_grid(input.row, input.col) &&
         sheet.data_table_covering(input.row, input.col) == nullptr &&
         sheet.spill_region_covering(input.row, input.col) == nullptr &&
         is_literal_scalar_slot(sheet.cell_at(input.row, input.col));
}

OwnedValue read_header(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = sheet.cell_at(row, col);
  return OwnedValue::from(cell != nullptr ? cell->cached_value : Value::blank());
}

OwnedValue read_result(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  return OwnedValue::result_of(sheet.cell_at(row, col));
}

std::uint64_t cell_count(const MergeRange& ref) {
  return static_cast<std::uint64_t>(ref.last_row - ref.first_row + 1U) * (ref.last_col - ref.first_col + 1U);
}

// The result grid of `table`, row-major over its `ref`. A table without a
// header row and column, or one whose recalc passes would exceed `passes`,
// is skipped (`computed == false`) and keeps its cached results; one whose
// input cell is unusable yields errors.
struct TableResult {
  bool computed = false;
  std::uint32_t cols = 0;
  std::vector<OwnedValue> cells;
};

Expected<TableResult, Error> compute_table(Workbook& wb, std::size_t sheet_index, const DataTable& table,
                                           const eval::FunctionRegistry& registry, ResourceBudget& passes) {
  TableResult result;
  const MergeRange& ref = table.ref;
  if (ref.first_row == 0U || ref.first_col == 0U || ResourceBudget(kMaxDataTableCells).would_exceed(cell_count(ref))) {
    return result;
  }
  const std::uint32_t rows = ref.last_row - ref.first_row + 1U;
  const std::uint32_t cols = ref.last_col - ref.first_col + 1U;
  const std::uint32_t top = ref.first_row - 1U;
  const std::uint32_t left = ref.first_col - 1U;
  const DataTableInput row_in = table.row_input();
  const DataTableInput col_in = table.col_input();
  const bool needs_row_in = table.two_dimensional || table.row_oriented;
  const bool needs_col_in = table.two_dimensional || !table.row_oriented;
  // One settle before the headers, then one per substitution.
  const std::uint64_t substitutions = needs_row_in && needs_col_in ? cell_count(ref) : needs_col_in ? rows : cols;
  if (!passes.consume(1U + substitutions)) {
    return result;
  }
  result.computed = true;
  result.cols = cols;
  result.cells.assign(static_cast<std::size_t>(rows) * cols, OwnedValue::error(ErrorCode::Ref));

  const Sheet& sheet = wb.sheet(sheet_index);
  if ((needs_row_in && (!input_usable(sheet, row_in) || range_contains(ref, row_in.row, row_in.col))) ||
      (needs_col_in && (!input_usable(sheet, col_in) || range_contains(ref, col_in.row, col_in.col)))) {
    return result;
  }

  const auto settle = [&]() -> Expected<void, Error> {
    auto recalculated = wb.recalc_engine().recalc(wb, registry);
    if (!recalculated) {
      return std::move(recalculated.error());
    }
    return Expected<void, Error>::Ok();
  };
  // An earlier table leaves its inputs' dependents stale; settle them before
  // reading headers, which may depend on the inputs.
  RETURN_IF_ERROR(settle());

  // Headers are read before any substitution.
  std::vector<OwnedValue> top_values;
  std::vector<OwnedValue> left_values;
  for (std::uint32_t j = 0; j < cols; ++j) {
    top_values.push_back(read_header(sheet, top, ref.first_col + j));
  }
  for (std::uint32_t i = 0; i < rows; ++i) {
    left_values.push_back(read_header(sheet, ref.first_row + i, left));
  }

  const auto at = [&](std::uint32_t i, std::uint32_t j) -> OwnedValue& {
    return result.cells[static_cast<std::size_t>(i) * cols + j];
  };

  // The loops stop at the first failure; the inputs are restored either way.
  Expected<void, Error> status = Expected<void, Error>::Ok();
  if (needs_row_in && needs_col_in) {
    InputCell row_cell(wb, sheet_index, row_in);
    InputCell col_cell(wb, sheet_index, col_in);
    for (std::uint32_t i = 0; i < rows && status; ++i) {
      for (std::uint32_t j = 0; j < cols && status; ++j) {
        status = row_cell.set(top_values[j]);
        if (status) {
          status = col_cell.set(left_values[i]);
        }
        if (status) {
          status = settle();
        }
        if (status) {
          at(i, j) = read_result(sheet, top, left);
        }
      }
    }
    auto restored_col = col_cell.restore();
    auto restored_row = row_cell.restore();
    if (status && !restored_col) {
      status = std::move(restored_col);
    }
    if (status && !restored_row) {
      status = std::move(restored_row);
    }
  } else if (needs_col_in) {
    InputCell cell(wb, sheet_index, col_in);
    for (std::uint32_t i = 0; i < rows && status; ++i) {
      status = cell.set(left_values[i]);
      if (status) {
        status = settle();
      }
      for (std::uint32_t j = 0; j < cols && status; ++j) {
        at(i, j) = read_result(sheet, top, ref.first_col + j);
      }
    }
    auto restored = cell.restore();
    if (status && !restored) {
      status = std::move(restored);
    }
  } else {
    InputCell cell(wb, sheet_index, row_in);
    for (std::uint32_t j = 0; j < cols && status; ++j) {
      status = cell.set(top_values[j]);
      if (status) {
        status = settle();
      }
      for (std::uint32_t i = 0; i < rows && status; ++i) {
        at(i, j) = read_result(sheet, ref.first_row + i, left);
      }
    }
    auto restored = cell.restore();
    if (status && !restored) {
      status = std::move(restored);
    }
  }
  if (!status) {
    return std::move(status.error());
  }
  return result;
}

}  // namespace

Expected<void, Error> Workbook::add_data_table(std::size_t sheet_index, MergeRange ref, DataTableInput row_input,
                                               DataTableInput col_input) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "add_data_table: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  ref = normalized(ref);
  if (!Sheet::coord_in_grid(ref.first_row, ref.first_col) || !Sheet::coord_in_grid(ref.last_row, ref.last_col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "add_data_table: ref outside grid");
  }
  for (const DataTableInput* input : {&row_input, &col_input}) {
    if (input->present && !Sheet::coord_in_grid(input->row, input->col)) {
      return make_error(FormulonErrorCode::kInvalidArgument, "add_data_table: input cell outside grid");
    }
  }
  if (!row_input.present && !col_input.present) {
    return make_error(FormulonErrorCode::kDataTableInvalid, "add_data_table: no input cell");
  }
  if (ref.first_row == 0U || ref.first_col == 0U) {
    return make_error(FormulonErrorCode::kDataTableInvalid,
                      "add_data_table: ref needs a row above and a column left of it");
  }
  if (auto charged = ResourceBudget(kMaxDataTableCells).consume(cell_count(ref)); !charged) {
    return std::move(charged.error());
  }
  for (const DataTableInput* input : {&row_input, &col_input}) {
    if (input->present && range_contains(ref, input->row, input->col)) {
      return make_error(FormulonErrorCode::kDataTableInvalid, "add_data_table: input cell inside ref");
    }
  }
  for (const DataTable& other : sheets_[sheet_index].data_tables()) {
    if (other.ref.first_row <= ref.last_row && ref.first_row <= other.ref.last_row &&
        other.ref.first_col <= ref.last_col && ref.first_col <= other.ref.last_col) {
      return make_error(FormulonErrorCode::kDataTableInvalid, "add_data_table: ref overlaps another data table");
    }
  }

  if (intersects_array_or_spill(sheets_[sheet_index], ref)) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "add_data_table: cannot change part of an array");
  }

  DataTable table;
  table.ref = ref;
  table.two_dimensional = row_input.present && col_input.present;
  // A two-variable table is stored as a row table, as the desktop app writes it.
  table.row_oriented = row_input.present;
  table.r1 = table.two_dimensional || table.row_oriented ? row_input : col_input;
  if (table.two_dimensional) {
    table.r2 = col_input;
  }
  table.r1.deleted = false;
  table.r2.deleted = false;
  table.always_calc = true;
  table.explicit_attrs = kDataTableExplicitDt2D | kDataTableExplicitDtr | kDataTableExplicitDel1 |
                         kDataTableExplicitDel2 | kDataTableExplicitCa;
  // The cells become table results (values only, no formulas) before the
  // table exists, so the write refusal does not apply to them.
  for (std::uint32_t row = ref.first_row; row <= ref.last_row; ++row) {
    for (std::uint32_t col = ref.first_col; col <= ref.last_col; ++col) {
      put_cell_value(sheet_index, row, col, Value::number(0.0));
    }
  }
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    sheets_[sheet_index].mutable_data_tables().push_back(table);
  }
  data_tables_pending_ = true;
  return Expected<void, Error>::Ok();
}

const DataTable* Workbook::data_table_at(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) const noexcept {
  return sheet_index < sheets_.size() ? sheets_[sheet_index].data_table_covering(row, col) : nullptr;
}

Expected<void, Error> Workbook::remove_data_table(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const DataTable* found = data_table_at(sheet_index, row, col);
  if (found == nullptr) {
    return make_error(FormulonErrorCode::kDataTableInvalid, "remove_data_table: no data table at cell",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  const MergeRange ref = found->ref;
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    std::vector<DataTable>& tables = sheets_[sheet_index].mutable_data_tables();
    tables.erase(tables.begin() + (found - tables.data()));
  }
  for (std::uint32_t r = ref.first_row; r <= ref.last_row; ++r) {
    for (std::uint32_t c = ref.first_col; c <= ref.last_col; ++c) {
      put_cell_value(sheet_index, r, c, Value::blank());
    }
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::compute_data_tables(const eval::FunctionRegistry& registry) {
  data_tables_pending_ = false;
  bool wrote = false;
  ResourceBudget passes(kMaxDataTableRecalcPasses);
  for (std::size_t sheet_index = 0; sheet_index < sheets_.size(); ++sheet_index) {
    // Copied: computing a table rewrites cells, never the table list.
    const std::vector<DataTable> tables = sheets_[sheet_index].data_tables();
    for (const DataTable& table : tables) {
      ASSIGN_OR_RETURN(TableResult result, compute_table(*this, sheet_index, table, registry, passes));
      if (!result.computed) {
        continue;
      }
      for (std::size_t k = 0; k < result.cells.size(); ++k) {
        const std::uint32_t row = table.ref.first_row + static_cast<std::uint32_t>(k / result.cols);
        const std::uint32_t col = table.ref.first_col + static_cast<std::uint32_t>(k % result.cols);
        const OwnedValue& v = result.cells[k];
        if (v.is_text) {
          put_cell_text(sheet_index, row, col, v.text);
        } else {
          put_cell_value(sheet_index, row, col, v.value);
        }
      }
      wrote = true;
    }
  }
  if (!wrote) {
    return Expected<void, Error>::Ok();
  }
  // One pass for the inputs' restored dependents and the cells reading results.
  auto settled = engine_->recalc(*this, registry);
  if (!settled) {
    return std::move(settled.error());
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::recalc_data_tables(const eval::FunctionRegistry& registry) {
  return compute_data_tables(registry);
}

bool Workbook::has_data_tables() const {
  // A concurrent sheet removal moves `sheets_` under the engine mutex.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  for (const Sheet& sheet : sheets_) {
    if (!sheet.data_tables().empty()) {
      return true;
    }
  }
  return false;
}

Expected<void, Error> Workbook::recalc_data_tables_after_pass(bool evaluated_any,
                                                              const eval::FunctionRegistry& registry) {
  if (!has_data_tables()) {
    data_tables_pending_ = false;
    return Expected<void, Error>::Ok();
  }
  if (calc_mode_ == CalcMode::kAutoNoTable) {
    data_tables_pending_ = true;
    return Expected<void, Error>::Ok();
  }
  if (!evaluated_any && !data_tables_pending_) {
    return Expected<void, Error>::Ok();
  }
  return compute_data_tables(registry);
}

}  // namespace formulon
