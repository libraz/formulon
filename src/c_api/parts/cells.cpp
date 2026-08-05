//
// C ABI - cell mutation (set_*), cell read (get_value, lambda_text_at),
// flat-enumeration (cell_count, cell_at), and dynamic-array spill info.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cell.h"
#include "eval/lambda_format.h"
#include "eval/lambda_value.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_u32;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::value_to_fm;

// ---------------------------------------------------------------------------
// Cell mutation
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_set_number(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              double value) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_number"); rc != 0) {
    return rc;
  }
  auto r = wb->workbook().set_cell_value(sheet_index, row, col, formulon::Value::number(value));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_bool(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                            int32_t value) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_bool"); rc != 0) {
    return rc;
  }
  auto r = wb->workbook().set_cell_value(sheet_index, row, col, formulon::Value::boolean(value != 0));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_error(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             fm_error_code_t error) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_error"); rc != 0) {
    return rc;
  }
  if (error < 0 || error > static_cast<fm_error_code_t>(formulon::ErrorCode::Unknown)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_error: error code out of range", "error=" + std::to_string(error));
  }
  auto r = wb->workbook().set_cell_value(sheet_index, row, col,
                                         formulon::Value::error(static_cast<formulon::ErrorCode>(error)));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_text(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                            const char* utf8) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_text"); rc != 0) {
    return rc;
  }
  if (utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_set_text: utf8 is NULL");
  }
  // Copy directly into the destination cell. Keeping every overwritten
  // input in a handle-global deque leaked one string per C-ABI write.
  auto r = wb->workbook().set_cell_text(sheet_index, row, col, utf8);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_cell_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                     const char* utf8) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_cell_phonetic"); rc != 0) {
    return rc;
  }
  if (utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_cell_phonetic: utf8 is NULL");
  }
  if (row >= formulon::Sheet::kMaxRows || col >= formulon::Sheet::kMaxCols) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_cell_phonetic: cell coordinate out of range");
  }
  wb->workbook().sheet(sheet_index).set_cell_phonetic(row, col, utf8);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_blank(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_blank"); rc != 0) {
    return rc;
  }
  auto r = wb->workbook().set_cell_value(sheet_index, row, col, formulon::Value::blank());
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_formula(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                               const char* formula) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_formula"); rc != 0) {
    return rc;
  }
  if (formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_formula: formula is NULL");
  }
  auto r = wb->workbook().set_cell_formula(sheet_index, row, col, std::string(formula));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Cell read
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_get_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             fm_value_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_get_value: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_get_value"); rc != 0) {
    return rc;
  }
  // `resolve_cell_value` is the spill-aware accessor: phantoms of a
  // dynamic-array spill surface their array cell rather than the raw
  // (blank) cached value.
  const formulon::Value v = wb->workbook().sheet(sheet_index).resolve_cell_value(row, col);
  // Read-path strings go to `read_scratch`, which is cleared on each read
  // so it holds only this call's output. Cast away const because the
  // scratch store is logically internal: mutating it does not affect the
  // workbook's observable state. The returned text pointer is valid until
  // the next read on this handle (see `formulon_c.h`).
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  value_to_fm(v, store, out);
  return 0;
}

extern "C" fm_status_t fm_workbook_get_cell_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                     uint32_t col, const char** out_text) {
  clear_last_error();
  if (out_text == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_get_cell_phonetic: out_text is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_get_cell_phonetic"); rc != 0) {
    return rc;
  }
  const formulon::Cell* cell = wb->workbook().sheet(sheet_index).cell_at(row, col);
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  store.emplace_back(cell == nullptr ? std::string() : cell->phonetic_text);
  *out_text = store.back().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_lambda_text_at(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                  const char** out_text) {
  clear_last_error();
  if (out_text == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_lambda_text_at: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_lambda_text_at"); rc != 0) {
    return rc;
  }
  const formulon::Cell* cell = wb->workbook().sheet(sheet_index).cell_at(row, col);
  if (cell == nullptr || !cell->cached_value.is_lambda()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_lambda_text_at: cell does not hold a lambda value",
        "sheet_index=" + std::to_string(sheet_index) + " row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  const formulon::eval::LambdaValue* lv = cell->cached_value.as_lambda();
  if (lv == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_lambda_text_at: lambda payload is NULL");
  }
  std::string formatted = formulon::eval::format_lambda_value(*lv);
  // Read-path scratch: cleared per call so the lambda text does not
  // accumulate across repeated reads. The returned pointer is valid until
  // the next read on this handle (see `formulon_c.h`).
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  store.emplace_back(std::move(formatted));
  *out_text = store.back().c_str();
  return 0;
}

// ---------------------------------------------------------------------------
// Iteration / dump
// ---------------------------------------------------------------------------
//
// Sheets are stored row-sparse (`unordered_map<row, RowCells>`); we
// surface a flat enumeration to bindings by materialising a sorted
// `(row, col)` index on demand. The cache is purely an optimisation;
// correctness does not depend on it.

namespace {

// Returns the `(row, col)` indices of every stored cell on `sheet`,
// sorted by `(row, col)` ascending. Implicitly default-constructed cells
// (those that exist only because a later column was touched in the same
// row) are kept: the dump command may want to surface them as blank
// slots, and dropping them here would make the count returned by
// `fm_workbook_cell_count` mismatch the indexable range. The CLI
// filters them out at render time. Columns before a row's first populated
// one are never materialised, so they are neither counted nor enumerated —
// this walks the row's stored run, matching `Sheet::cell_count`.
std::vector<std::pair<std::uint32_t, std::uint32_t>> collect_cell_addresses(const formulon::Sheet& sheet) {
  std::vector<std::pair<std::uint32_t, std::uint32_t>> out;
  for (const auto& [row, cells] : sheet.rows()) {
    for (std::size_t col = cells.first_col(); col < cells.size(); ++col) {
      out.emplace_back(row, static_cast<std::uint32_t>(col));
    }
  }
  // Dynamic-array spill phantoms carry an effective value through
  // `resolve_cell_value` but live only in the spill table, absent from
  // `rows()`. Merge them so the flat enumeration surfaces spilled cells, then
  // sort + unique: a phantom that coincides with an implicitly default-
  // constructed slot must appear only once.
  for (const formulon::CellAddress& addr : sheet.spill_phantom_addresses()) {
    out.emplace_back(addr.row, addr.col);
  }
  std::sort(out.begin(), out.end());
  out.erase(std::unique(out.begin(), out.end()), out.end());
  return out;
}

const std::vector<std::pair<std::uint32_t, std::uint32_t>>& cached_cell_addresses(const fm_workbook_t* wb,
                                                                                  std::size_t sheet_index) {
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  const std::uint64_t revision = sheet.cell_enumeration_revision();
  auto& cache = wb->cell_enumeration_cache;
  if (cache.sheet_index != sheet_index || cache.revision != revision) {
    cache.addresses = collect_cell_addresses(sheet);
    cache.sheet_index = sheet_index;
    cache.revision = revision;
  }
  return cache.addresses;
}

}  // namespace

// `fm_workbook_cell_count` is now emitted by the binding codegen (see
// `src/c_api/generated/sheet_counts.cpp`).

extern "C" fm_status_t fm_workbook_cell_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, uint32_t* out_row,
                                           uint32_t* out_col, const char** out_formula, fm_value_t* out_value) {
  clear_last_error();
  if (out_row == nullptr || out_col == nullptr || out_value == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_cell_at: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_cell_at"); rc != 0) {
    return rc;
  }
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  // The sorted address vector is cached on the handle until the Sheet's
  // enumeration revision changes. A `cell_count` / `cell_at(0..N)` pass
  // therefore pays the O(N log N) collection once rather than per cell.
  const auto& addrs = cached_cell_addresses(wb, sheet_index);
  if (idx >= addrs.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_cell_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(addrs.size()));
  }
  const auto [row, col] = addrs[idx];
  *out_row = row;
  *out_col = col;
  // `(row, col)` may be a spill phantom, which has no stored `Cell`: that is
  // a normal case here, not an error. Phantoms carry no formula text and
  // their value is resolved below via `resolve_cell_value`. Only a stored
  // cell can contribute a formula pointer.
  const formulon::Cell* cell = sheet.cell_at(row, col);
  if (out_formula != nullptr) {
    *out_formula = (cell != nullptr && !cell->formula_text.empty()) ? cell->formula_text.c_str() : nullptr;
  }
  // Use the spill-aware accessor so phantoms surface their owning anchor's
  // value. The phantoms themselves are still indexed via `cell_at`'s
  // implicit default cells; users that want only stored formulae filter
  // by `out_formula != NULL`.
  const formulon::Value v = sheet.resolve_cell_value(row, col);
  // `out_value`'s text payload goes to `read_scratch` (cleared per call,
  // valid until the next read). `out_formula` above points into the
  // cell's own `formula_text`, not scratch, so clearing here is safe.
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  value_to_fm(v, store, out_value);
  return 0;
}

// ---------------------------------------------------------------------------
// Dynamic-array spill payload
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_spill_info(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                              std::uint32_t col, fm_spill_info_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_spill_info: NULL argument");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_workbook_spill_info"); rc != 0) {
    return rc;
  }
  *out = fm_spill_info_t{};
  const auto& s = wb->workbook().sheet(sheet);
  // Try anchor lookup first; fall back to phantom-coverage map.
  const formulon::SpillRegion* region = s.spill_region_at_anchor(row, col);
  if (region == nullptr) {
    region = s.spill_region_covering(row, col);
  }
  if (region == nullptr) {
    out->engaged = 0;
    return 0;
  }
  out->anchor_row = region->anchor_row;
  out->anchor_col = region->anchor_col;
  out->rows = region->rows;
  out->cols = region->cols;
  out->engaged = 1;
  return 0;
}
