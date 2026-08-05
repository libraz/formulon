//
// C ABI - ad-hoc, side-effect-free formula evaluation. Bridges the
// engine-side `eval::evaluate_formula_text` / `eval::evaluate_cf_formula`
// drivers into the stable C ABI. Both entry points observe the workbook
// read-only (the `const fm_workbook_t*` is the purity contract) and return
// a scalar `fm_value_t`; text payloads borrow the handle's read scratch.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/adhoc_eval.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::AdhocArrayStash;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::value_to_fm;

namespace {

// Returns a copy of `v` whose `Text` payload (if any) is owned by `owner`
// rather than borrowing the evaluation arena. Non-text kinds are returned
// unchanged; Array / Ref / Lambda cells keep their arena pointers, which the
// C ABI never dereferences (it reports those kinds by tag only).
formulon::Value own_cell_text(const formulon::Value& v, TextStore& owner) {
  if (v.kind() != formulon::ValueKind::Text) {
    return v;
  }
  const std::string_view text = v.as_text();
  owner.emplace_back(text.data(), text.size());
  return formulon::Value::text(std::string_view(owner.back()));
}

}  // namespace

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
  if (arena.exhausted()) {
    return set_binding_error(formulon::FormulonErrorCode::kOutOfMemory,
                             "fm_workbook_evaluate_formula: evaluation arena exhausted");
  }
  // Read-path scratch: cleared per call so returned text pointers stay
  // valid only until the next read (see `fm_workbook_get_value`). Cast away
  // const because the scratch is logically internal — it never affects the
  // workbook's observable state.
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  value_to_fm(v, store, out);
  return 0;
}

extern "C" fm_status_t fm_workbook_evaluate_formula_array(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                          uint32_t col, const char* formula, uint32_t* out_rows,
                                                          uint32_t* out_cols) {
  clear_last_error();
  if (out_rows == nullptr || out_cols == nullptr || formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_evaluate_formula_array: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_evaluate_formula_array"); rc != 0) {
    return rc;
  }
  // Per-call arena backs any text payload in the evaluated result until it is
  // deep-copied into the per-handle stash below; the arena is discarded when
  // this call returns.
  formulon::Arena arena;
  const formulon::Value v =
      formulon::eval::evaluate_formula_text_array(wb->workbook(), wb->workbook().sheet(sheet_index), row, col,
                                                  std::string_view(formula), arena, formulon::eval::default_registry());
  if (arena.exhausted()) {
    return set_binding_error(formulon::FormulonErrorCode::kOutOfMemory,
                             "fm_workbook_evaluate_formula_array: evaluation arena exhausted");
  }

  // Stash the whole result on the handle. Cast away const for the same reason
  // as the read scratch: the stash is internal handle state, never part of
  // the workbook's observable value graph, so the purity contract holds.
  AdhocArrayStash& stash = const_cast<AdhocArrayStash&>(wb->adhoc_array);
  stash.clear();

  if (v.is_array()) {
    const uint32_t rows = v.as_array_rows();
    const uint32_t cols = v.as_array_cols();
    if (rows == 0 || cols == 0) {
      // Degenerate empty array: surface as a single #VALUE! cell, matching
      // the scalar reduction's empty-array handling.
      stash.rows = 1;
      stash.cols = 1;
      stash.cells.push_back(formulon::Value::error(formulon::ErrorCode::Value));
    } else {
      stash.rows = rows;
      stash.cols = cols;
      const formulon::Value* cells = v.as_array_cells();
      const size_t count = static_cast<size_t>(rows) * static_cast<size_t>(cols);
      stash.cells.reserve(count);
      for (size_t i = 0; i < count; ++i) {
        stash.cells.push_back(own_cell_text(cells[i], stash.text_owner));
      }
    }
  } else {
    // Scalar result: a 1x1 array.
    stash.rows = 1;
    stash.cols = 1;
    stash.cells.push_back(own_cell_text(v, stash.text_owner));
  }

  *out_rows = stash.rows;
  *out_cols = stash.cols;
  return 0;
}

extern "C" fm_status_t fm_workbook_evaluate_formula_array_cell(const fm_workbook_t* wb, size_t index, fm_value_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_evaluate_formula_array_cell: NULL argument");
  }
  const AdhocArrayStash& stash = wb->adhoc_array;
  if (index >= stash.cells.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_evaluate_formula_array_cell: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(stash.cells.size()));
  }
  // Read-path scratch: cleared per call so returned text pointers stay valid
  // only until the next read (see `fm_workbook_evaluate_formula_array_cell`
  // contract in `formulon_c.h`). Cast away const for the same reason as the
  // other read-path entry points.
  TextStore& scratch = const_cast<TextStore&>(wb->read_scratch);
  scratch.clear();
  value_to_fm(stash.cells[index], scratch, out);
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
  if (arena.exhausted()) {
    return set_binding_error(formulon::FormulonErrorCode::kOutOfMemory,
                             "fm_workbook_evaluate_cf_formula: evaluation arena exhausted");
  }
  *out = fm_value_t{};
  out->kind = FM_VAL_BOOL;
  out->u.boolean = fired ? 1 : 0;
  return 0;
}
