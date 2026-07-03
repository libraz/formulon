// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// C ABI - ad-hoc, side-effect-free formula evaluation. Bridges the
// engine-side `eval::evaluate_formula_text` / `eval::evaluate_cf_formula`
// drivers into the stable C ABI. Both entry points observe the workbook
// read-only (the `const fm_workbook_t*` is the purity contract) and return
// a scalar `fm_value_t`; text payloads borrow the handle's read scratch.

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/adhoc_eval.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::value_to_fm;

extern "C" fm_status_t fm_workbook_evaluate_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                    uint32_t col, const char* formula, fm_value_t* out) {
  clear_last_error();
  if (out == nullptr || formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_evaluate_formula: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_evaluate_formula"); rc != 0) {
    return rc;
  }
  // Per-call arena backs any text payload in the result until it is copied
  // into the read scratch below; both are discarded when this call returns.
  formulon::Arena arena;
  const formulon::Value v =
      formulon::eval::evaluate_formula_text(wb->workbook(), wb->workbook().sheet(sheet_index), row, col,
                                            std::string_view(formula), arena, formulon::eval::default_registry());
  // Read-path scratch: cleared per call so returned text pointers stay
  // valid only until the next read (see `fm_workbook_get_value`). Cast away
  // const because the scratch is logically internal — it never affects the
  // workbook's observable state.
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  value_to_fm(v, store, out);
  return 0;
}

extern "C" fm_status_t fm_workbook_evaluate_cf_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                       uint32_t col, uint32_t anchor_row, uint32_t anchor_col,
                                                       const char* formula, fm_value_t* out) {
  clear_last_error();
  if (out == nullptr || formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_evaluate_cf_formula: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_evaluate_cf_formula"); rc != 0) {
    return rc;
  }
  formulon::Arena arena;
  const bool fired = formulon::eval::evaluate_cf_formula(wb->workbook(), wb->workbook().sheet(sheet_index), row, col,
                                                         anchor_row, anchor_col, std::string_view(formula), arena,
                                                         formulon::eval::default_registry());
  *out = fm_value_t{};
  out->kind = FM_VAL_BOOL;
  out->u.boolean = fired ? 1 : 0;
  return 0;
}
