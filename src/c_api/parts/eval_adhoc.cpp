//
// C ABI - ad-hoc, side-effect-free formula evaluation. Bridges the
// engine-side `eval::evaluate_parsed_formula` / `eval::evaluate_parsed_cf_formula`
// drivers into the stable C ABI. A formula that does not parse fails before
// evaluation. Every entry point observes the workbook read-only apart from
// bringing pending results up to date (`ensure_calculated`). Scalar
// results borrow the handle's read scratch for text; an array result is an
// owned `fm_value_array_t`.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/adhoc_eval.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::ensure_calculated;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::parse_evaluated_formula;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::value_to_fm;

// Owned result of `fm_sheet_evaluate_formula_array`. Every `Text` cell is
// re-pointed at `text_owner`, so the cells outlive the evaluation arena; the
// C ABI reports Array / Ref / Lambda cells by kind only and never
// dereferences their arena pointers.
struct fm_value_array {
  std::vector<formulon::Value> cells;  // row-major; size == rows * cols
  uint32_t rows = 0;
  uint32_t cols = 0;
  TextStore text_owner;
};

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

// Per-call arena for a caller-supplied formula. The byte ceiling keeps a
// hostile formula to a recoverable error instead of an allocation failure
// that takes the host down.
formulon::Arena make_eval_arena() {
  return formulon::Arena(formulon::kDefaultArenaChunkBytes, formulon::kMaxEvalArenaBytes);
}

fm_status_t check_adhoc_cell(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return check_sheet_rect(row, col, row, col);
}

// `<file_name>!Name` names this workbook and `[@Qty]` the table holding the
// cell; no external book is registered.
std::string bind_adhoc_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                               const char* formula) {
  return wb->workbook().bind_self_book_names(
      wb->workbook().bind_entered_structured_refs(formula, sheet_index, row, col));
}

fm_status_t arena_exhausted_error() {
  return set_binding_error(formulon::FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted");
}

// Status of a parse that yielded no AST: running out of arena is not the
// syntax failure `parse_evaluated_formula` recorded for it.
fm_status_t parse_failure_status(const formulon::Arena& arena, fm_status_t rc) {
  if (!arena.exhausted()) {
    return rc;
  }
  clear_last_error();
  return arena_exhausted_error();
}

}  // namespace

extern "C" fm_status_t fm_sheet_evaluate_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                 uint32_t col, const char* formula, fm_value_t* out) {
  clear_last_error();
  if (out == nullptr || formula == nullptr) {
    return null_argument();
  }
  if (auto rc = check_adhoc_cell(wb, sheet_index, row, col); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({formula}); rc != 0) {
    return rc;
  }
  // The arena backs any text payload in the result until it is copied into
  // the read scratch below.
  formulon::Arena arena = make_eval_arena();
  const std::string text = bind_adhoc_formula(wb, sheet_index, row, col, formula);
  fm_status_t rc = 0;
  const formulon::parser::AstNode* root = parse_evaluated_formula(formula, text, arena, &rc);
  if (root == nullptr) {
    return parse_failure_status(arena, rc);
  }
  if (rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  const formulon::Value v = formulon::eval::evaluate_parsed_formula(
      wb->workbook(), wb->workbook().sheet(sheet_index), row, col, *root, arena, formulon::eval::default_registry());
  if (arena.exhausted()) {
    return arena_exhausted_error();
  }
  // Returned text pointers stay valid only until the next read (see
  // `fm_sheet_get_value`).
  TextStore& store = fresh_read_scratch(wb);
  value_to_fm(v, store, out);
  return 0;
}

extern "C" fm_status_t fm_sheet_evaluate_formula_array(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                       uint32_t col, const char* formula, fm_value_array_t** out) {
  clear_last_error();
  if (out == nullptr || formula == nullptr) {
    return null_argument();
  }
  if (auto rc = check_adhoc_cell(wb, sheet_index, row, col); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({formula}); rc != 0) {
    return rc;
  }
  // The arena backs any text payload in the result until it is deep-copied
  // into the returned handle below.
  formulon::Arena arena = make_eval_arena();
  const std::string text = bind_adhoc_formula(wb, sheet_index, row, col, formula);
  fm_status_t rc = 0;
  const formulon::parser::AstNode* root = parse_evaluated_formula(formula, text, arena, &rc);
  if (root == nullptr) {
    return parse_failure_status(arena, rc);
  }
  if (rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  const formulon::Value v = formulon::eval::evaluate_parsed_formula_array(
      wb->workbook(), wb->workbook().sheet(sheet_index), row, col, *root, arena, formulon::eval::default_registry());
  if (arena.exhausted()) {
    return arena_exhausted_error();
  }

  auto result = std::unique_ptr<fm_value_array_t>(new fm_value_array_t{});
  if (v.is_array() && v.as_array_rows() != 0 && v.as_array_cols() != 0) {
    result->rows = v.as_array_rows();
    result->cols = v.as_array_cols();
    const formulon::Value* cells = v.as_array_cells();
    const size_t count = static_cast<size_t>(result->rows) * static_cast<size_t>(result->cols);
    result->cells.reserve(count);
    for (size_t i = 0; i < count; ++i) {
      result->cells.push_back(own_cell_text(cells[i], result->text_owner));
    }
  } else {
    // A scalar is a 1x1 array; a degenerate empty array is a single #VALUE!,
    // matching the scalar reduction's empty-array handling.
    result->rows = 1;
    result->cols = 1;
    result->cells.push_back(v.is_array() ? formulon::Value::error(formulon::ErrorCode::Value)
                                         : own_cell_text(v, result->text_owner));
  }
  *out = result.release();
  return 0;
}

extern "C" void fm_value_array_destroy(fm_value_array_t* array) {
  delete array;
}

extern "C" fm_status_t fm_value_array_get_dimensions(const fm_value_array_t* array, uint32_t* out_rows,
                                                     uint32_t* out_cols) {
  clear_last_error();
  if (array == nullptr || out_rows == nullptr || out_cols == nullptr) {
    return null_argument();
  }
  *out_rows = array->rows;
  *out_cols = array->cols;
  return 0;
}

extern "C" fm_status_t fm_value_array_get_value_count(const fm_value_array_t* array, size_t* out_count) {
  clear_last_error();
  if (array == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = array->cells.size();
  return 0;
}

extern "C" fm_status_t fm_value_array_get_value_at(const fm_value_array_t* array, size_t index, fm_value_t* out) {
  clear_last_error();
  if (array == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(index, array->cells.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::Value& cell = array->cells[index];
  if (cell.kind() == formulon::ValueKind::Text) {
    // The handle owns the NUL-terminated bytes behind every text cell.
    out->kind = FM_VAL_TEXT;
    out->u.text = cell.as_text().data();
    return 0;
  }
  TextStore unused;
  value_to_fm(cell, unused, out);
  return 0;
}

extern "C" fm_status_t fm_sheet_evaluate_cf_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                    uint32_t col, uint32_t anchor_row, uint32_t anchor_col,
                                                    const char* formula, fm_value_t* out) {
  clear_last_error();
  if (out == nullptr || formula == nullptr) {
    return null_argument();
  }
  if (auto rc = check_adhoc_cell(wb, sheet_index, row, col); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(anchor_row, anchor_col, anchor_row, anchor_col); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({formula}); rc != 0) {
    return rc;
  }
  formulon::Arena arena = make_eval_arena();
  const std::string text = bind_adhoc_formula(wb, sheet_index, row, col, formula);
  fm_status_t rc = 0;
  const formulon::parser::AstNode* root = parse_evaluated_formula(formula, text, arena, &rc);
  if (root == nullptr) {
    return parse_failure_status(arena, rc);
  }
  if (rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  const bool fired = formulon::eval::evaluate_parsed_cf_formula(wb->workbook(), wb->workbook().sheet(sheet_index), row,
                                                                col, anchor_row, anchor_col, *root, arena,
                                                                formulon::eval::default_registry());
  if (arena.exhausted()) {
    return arena_exhausted_error();
  }
  *out = fm_value_t{};
  out->kind = FM_VAL_BOOL;
  out->u.boolean = fired ? 1 : 0;
  return 0;
}
