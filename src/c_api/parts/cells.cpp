// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::intern_text;
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

extern "C" fm_status_t fm_workbook_set_text(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                            const char* utf8) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_text"); rc != 0) {
    return rc;
  }
  if (utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_set_text: utf8 is NULL");
  }
  // The cell stores a non-owning view; we must keep the bytes alive for
  // as long as the handle does.
  const std::string_view view = intern_text(wb->text_store, std::string_view(utf8));
  auto r = wb->workbook().set_cell_value(sheet_index, row, col, formulon::Value::text(view));
  if (!r) {
    return set_last_error(r.error());
  }
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
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_get_value: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_get_value",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  // `resolve_cell_value` is the spill-aware accessor: phantoms of a
  // dynamic-array spill surface their array cell rather than the raw
  // (blank) cached value.
  const formulon::Value v = wb->workbook().sheet(sheet_index).resolve_cell_value(row, col);
  // `text_store` is a `std::deque`, so prior pointers handed out by this
  // accessor remain valid even when this call appends a new entry.
  // Cast away const to write into the per-handle text store; the store
  // is logically internal scratch space whose mutation does not affect
  // the workbook's observable state.
  TextStore& store = const_cast<TextStore&>(wb->text_store);
  value_to_fm(v, store, out);
  return 0;
}

extern "C" fm_status_t fm_workbook_lambda_text_at(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                  const char** out_text) {
  clear_last_error();
  if (wb == nullptr || out_text == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_lambda_text_at: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_lambda_text_at: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
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
  TextStore& store = const_cast<TextStore&>(wb->text_store);
  store.emplace_back(std::move(formatted));
  *out_text = store.back().c_str();
  return 0;
}

// ---------------------------------------------------------------------------
// Iteration / dump
// ---------------------------------------------------------------------------
//
// Sheets are stored row-sparse (`unordered_map<row, vector<Cell>>`); we
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
// filters them out at render time.
std::vector<std::pair<std::uint32_t, std::uint32_t>> collect_cell_addresses(const formulon::Sheet& sheet) {
  std::vector<std::pair<std::uint32_t, std::uint32_t>> out;
  for (const auto& [row, cells] : sheet.rows()) {
    for (std::size_t col = 0; col < cells.size(); ++col) {
      out.emplace_back(row, static_cast<std::uint32_t>(col));
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

}  // namespace

extern "C" fm_status_t fm_workbook_cell_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_cell_count: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_cell_count: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  *out_count = wb->workbook().sheet(sheet_index).cell_count();
  return 0;
}

extern "C" fm_status_t fm_workbook_cell_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, uint32_t* out_row,
                                           uint32_t* out_col, const char** out_formula, fm_value_t* out_value) {
  clear_last_error();
  if (wb == nullptr || out_row == nullptr || out_col == nullptr || out_value == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_cell_at: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_cell_at: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  // Materialise the sorted address vector. This is O(N log N) in the
  // sheet's cell count; the CLI calls cell_at in a tight loop so a
  // future optimisation could cache the vector on the handle. For the
  // current scope (workbooks up to ~100k populated cells) the simple
  // path is fast enough and avoids invalidation bookkeeping.
  const auto addrs = collect_cell_addresses(sheet);
  if (idx >= addrs.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_cell_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(addrs.size()));
  }
  const auto [row, col] = addrs[idx];
  *out_row = row;
  *out_col = col;
  const formulon::Cell* cell = sheet.cell_at(row, col);
  // `cell_at` must succeed because `(row, col)` came from the sheet's
  // own row vector. Guard defensively just in case the contract drifts.
  if (cell == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInternalError,
                             "fm_workbook_cell_at: cell vanished mid-iteration",
                             "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  if (out_formula != nullptr) {
    *out_formula = cell->formula_text.empty() ? nullptr : cell->formula_text.c_str();
  }
  // Use the spill-aware accessor so phantoms surface their owning anchor's
  // value. The phantoms themselves are still indexed via `cell_at`'s
  // implicit default cells; users that want only stored formulae filter
  // by `out_formula != NULL`.
  const formulon::Value v = sheet.resolve_cell_value(row, col);
  TextStore& store = const_cast<TextStore&>(wb->text_store);
  value_to_fm(v, store, out_value);
  return 0;
}

// ---------------------------------------------------------------------------
// Dynamic-array spill payload
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_spill_info(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                              std::uint32_t col, fm_spill_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_spill_info: NULL argument");
  }
  if (sheet >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_spill_info: sheet out of range", "sheet=" + std::to_string(sheet));
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
