//
// C ABI - What-If data tables: creating, reading and removing them, and the
// explicit table calculation.

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/function_registry.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

namespace {

formulon::DataTableInput input_of(int32_t has, uint32_t row, uint32_t col) {
  formulon::DataTableInput input;
  input.present = has != 0;
  input.row = row;
  input.col = col;
  return input;
}

}  // namespace

extern "C" fm_status_t fm_sheet_add_data_table(fm_workbook_t* wb, size_t sheet_index, const fm_data_table* spec) {
  clear_last_error();
  if (spec == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_add_data_table: NULL spec");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_add_data_table"); rc != 0) {
    return rc;
  }
  const formulon::MergeRange ref{spec->first_row, spec->first_col, spec->last_row, spec->last_col};
  auto added = wb->workbook().add_data_table(sheet_index, ref,
                                             input_of(spec->has_row_input, spec->row_input_row, spec->row_input_col),
                                             input_of(spec->has_col_input, spec->col_input_row, spec->col_input_col));
  if (!added) {
    return set_last_error(added.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_data_table_at(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              int32_t* out_present, fm_data_table* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_data_table_at: NULL out argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_data_table_at"); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out = fm_data_table{};
  const formulon::DataTable* table = wb->workbook().data_table_at(sheet_index, row, col);
  if (table == nullptr) {
    return 0;
  }
  const formulon::DataTableInput row_in = table->row_input();
  const formulon::DataTableInput col_in = table->col_input();
  *out_present = 1;
  out->first_row = table->ref.first_row;
  out->first_col = table->ref.first_col;
  out->last_row = table->ref.last_row;
  out->last_col = table->ref.last_col;
  out->has_row_input = row_in.present ? 1 : 0;
  out->has_col_input = col_in.present ? 1 : 0;
  out->row_input_row = row_in.row;
  out->row_input_col = row_in.col;
  out->col_input_row = col_in.row;
  out->col_input_col = col_in.col;
  out->two_dimensional = table->two_dimensional ? 1 : 0;
  out->row_oriented = table->row_oriented ? 1 : 0;
  out->row_input_deleted = row_in.deleted ? 1 : 0;
  out->col_input_deleted = col_in.deleted ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_data_table(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_remove_data_table"); rc != 0) {
    return rc;
  }
  auto removed = wb->workbook().remove_data_table(sheet_index, row, col);
  if (!removed) {
    return set_last_error(removed.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_recalc_data_tables(fm_workbook_t* wb) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_recalc_data_tables: wb is NULL");
  }
  auto r = wb->workbook().recalc_data_tables(formulon::eval::default_registry());
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}
