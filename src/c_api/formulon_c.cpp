// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the stable C ABI declared in `formulon_c.h`.
//
// Design notes:
//
//   * The opaque `fm_workbook_t` is a thin C++ struct that owns either a
//     bare `Workbook` (for `create` / `create_empty`) or a full
//     `io::OoxmlReadResult` (for `load`). The latter is required so the
//     `text_storage` deque outlives every `Value::text` view that the
//     reader populated; without it, the cell-text views would dangle the
//     moment the result was discarded.
//
//   * Every text byte that crosses the C boundary needs to be
//     NUL-terminated. `Value::text` carries a non-owning `string_view`
//     whose underlying storage is not guaranteed to be terminated, so
//     `fm_workbook_get_value` materialises an extra copy in a per-handle
//     `std::deque<std::string>`. The deque guarantees pointer stability,
//     which lets us hand out raw `c_str()` pointers that remain valid
//     until the handle is destroyed.
//
//   * `fm_workbook_set_text` likewise needs to own its UTF-8 payload
//     because the underlying `Value::text` is a non-owning view. We append
//     into the same per-handle deque.
//
//   * Diagnostics (`fm_last_error_*`) live in `thread_local std::string`s
//     and are written by every fallible entry point before it returns a
//     non-zero status. A successful call clears them so callers do not
//     observe stale state.
//
//   * Buffers returned by `fm_workbook_save` are allocated with `new
//     uint8_t[]` and matched with `fm_buffer_free`'s `delete[]`; this
//     pairing is documented in the public header.
//
//   * The implementation relies exclusively on `Expected<T, Error>` -
//     no exceptions, no RTTI, no `dynamic_cast` / `typeid`. Compiles
//     cleanly with `-fno-exceptions -fno-rtti`.

#include "c_api/formulon_c.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <memory>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_evaluator.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace {

// Thread-local diagnostics for the most recent C API call on this thread.
//
// Both buffers are always non-empty in the sense that they own valid string
// storage, but their `c_str()` may point to an empty string. Returning
// `c_str()` therefore satisfies the public "never NULL" contract.
thread_local std::string g_last_error_message;
thread_local std::string g_last_error_context;

// Resets the thread-local diagnostics. Called at the start of every
// fallible entry point so a successful return surfaces an empty error
// message rather than the previous call's residue.
void clear_last_error() {
  g_last_error_message.clear();
  g_last_error_context.clear();
}

// Captures `err` into the thread-local buffers and returns its numeric
// status code. The Error code is mirrored bit-for-bit into `fm_status_t`.
fm_status_t set_last_error(const formulon::Error& err) {
  g_last_error_message = err.message;
  g_last_error_context = err.context;
  return static_cast<fm_status_t>(err.code);
}

// Convenience wrapper for the very common "binding misuse" case (NULL
// pointer argument, unknown handle, ...). `context` is appended verbatim;
// callers should keep it short and machine-friendly (key=value).
fm_status_t set_binding_error(formulon::FormulonErrorCode code, const char* message, std::string context = {}) {
  formulon::Error err;
  err.code = code;
  err.message = message != nullptr ? message : "";
  err.context = std::move(context);
  return set_last_error(err);
}

// Per-handle UTF-8 storage. `std::deque` is used (matching the OOXML
// reader) for pointer stability: we hand out `c_str()` pointers from
// elements deep inside the container, so a `std::vector` reallocation
// would invalidate every previously surfaced view.
using TextStore = std::deque<std::string>;

// Inserts `text` (a non-owning UTF-8 view) into `store` and returns a
// non-owning `string_view` whose pointee is owned by `store`. Used by
// `fm_workbook_set_text` so the view stored on the cell remains valid
// for the lifetime of the handle.
std::string_view intern_text(TextStore& store, std::string_view text) {
  store.emplace_back(text.data(), text.size());
  return std::string_view(store.back());
}

}  // namespace

// Opaque handle definition.
//
// Holds either a bare `Workbook` plus its own text storage (for the
// `create` / `create_empty` paths) or the full `OoxmlReadResult` (for
// `load`). The `result_` slot is engaged for loaded workbooks so the
// reader's `text_storage` deque survives long enough for every cell view
// to remain valid.
struct fm_workbook {
  // For `load`: holds the entire OoxmlReadResult so its `text_storage`
  // outlives the workbook handle. The workbook reference below points
  // into `result_->workbook` when this slot is engaged.
  std::optional<formulon::io::OoxmlReadResult> result;

  // For `create` / `create_empty`: holds the workbook directly.
  // Ignored when `result_` is engaged.
  std::optional<formulon::Workbook> standalone;

  // Storage for UTF-8 strings owned by this handle: cell text inputs
  // (`fm_workbook_set_text`) and read-side NUL-terminated copies
  // (`fm_workbook_get_value` / `fm_workbook_sheet_name` for sheets that
  // would otherwise alias inline storage).
  TextStore text_store;

  // Returns the underlying workbook. Exactly one of the two `optional`
  // slots is engaged for any valid handle.
  formulon::Workbook& workbook() {
    if (result.has_value()) {
      return result->workbook;
    }
    return *standalone;
  }
  const formulon::Workbook& workbook() const {
    if (result.has_value()) {
      return result->workbook;
    }
    return *standalone;
  }
};

namespace {

// Translates a `Value` into a C-side `fm_value_t`. For text variants the
// payload is interned in `store` so the returned pointer is NUL-terminated
// and stable across other reads on the same handle.
void value_to_fm(const formulon::Value& v, TextStore& store, fm_value_t* out) {
  switch (v.kind()) {
    case formulon::ValueKind::Blank:
      out->kind = FM_VAL_BLANK;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Number:
      out->kind = FM_VAL_NUMBER;
      out->u.number = v.as_number();
      return;
    case formulon::ValueKind::Bool:
      out->kind = FM_VAL_BOOL;
      out->u.boolean = v.as_boolean() ? 1 : 0;
      return;
    case formulon::ValueKind::Text: {
      const std::string_view text = v.as_text();
      store.emplace_back(text.data(), text.size());
      out->kind = FM_VAL_TEXT;
      out->u.text = store.back().c_str();
      return;
    }
    case formulon::ValueKind::Error:
      out->kind = FM_VAL_ERROR;
      out->u.error_code = static_cast<int32_t>(v.as_error());
      return;
    case formulon::ValueKind::Array:
      // Array passthrough is reserved; the kind is reported but no
      // payload is exposed across the boundary in this bundle.
      out->kind = FM_VAL_ARRAY;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Ref:
      out->kind = FM_VAL_REF;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Lambda:
      out->kind = FM_VAL_LAMBDA;
      out->u.number = 0.0;
      return;
  }
  // Defensive default: surface as Blank rather than leaving uninitialised
  // bytes on the boundary. The switch above is exhaustive over every
  // `ValueKind` enumerator, so this branch is unreachable in practice.
  out->kind = FM_VAL_BLANK;
  out->u.number = 0.0;
}

}  // namespace

// ---------------------------------------------------------------------------
// Construction / lifecycle
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_create(fm_workbook_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_create: out is NULL");
  }
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->standalone.emplace(formulon::Workbook::create());
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_workbook_create_empty(fm_workbook_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_create_empty: out is NULL");
  }
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->standalone.emplace(formulon::Workbook::create_empty());
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_workbook_load(const uint8_t* bytes, size_t len, fm_workbook_t** out) {
  clear_last_error();
  if (bytes == nullptr || out == nullptr || len == 0) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_load: NULL or empty input");
  }
  formulon::io::ByteSpan span;
  span.data = bytes;
  span.size = len;
  auto result = formulon::io::read_ooxml(span);
  if (!result) {
    return set_last_error(result.error());
  }
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->result.emplace(std::move(result.value()));
  *out = handle.release();
  return 0;
}

extern "C" void fm_workbook_destroy(fm_workbook_t* wb) {
  // Mirrors `free(NULL)` semantics: silently accept NULL handles.
  delete wb;
}

// ---------------------------------------------------------------------------
// Save
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_save(const fm_workbook_t* wb, uint8_t** out_bytes, size_t* out_len) {
  clear_last_error();
  if (wb == nullptr || out_bytes == nullptr || out_len == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_save: NULL argument");
  }
  auto bytes = wb->workbook().save();
  if (!bytes) {
    return set_last_error(bytes.error());
  }
  const std::vector<std::uint8_t>& src = bytes.value();
  // Allocate with `new[]` so `fm_buffer_free`'s matching `delete[]` is
  // well-defined. The header documents the pairing.
  auto* buffer = new uint8_t[src.size()];
  if (!src.empty()) {
    std::memcpy(buffer, src.data(), src.size());
  }
  *out_bytes = buffer;
  *out_len = src.size();
  return 0;
}

extern "C" void fm_buffer_free(uint8_t* bytes) {
  delete[] bytes;
}

// ---------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------

extern "C" size_t fm_workbook_sheet_count(const fm_workbook_t* wb) {
  if (wb == nullptr) {
    return 0;
  }
  return wb->workbook().sheet_count();
}

extern "C" fm_status_t fm_workbook_sheet_name(const fm_workbook_t* wb, size_t index, const char** out_utf8) {
  clear_last_error();
  if (wb == nullptr || out_utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_sheet_name: NULL argument");
  }
  if (index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_sheet_name: sheet_index out of range",
        "sheet_index=" + std::to_string(index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  // `Sheet::name()` returns `const std::string&`, so `c_str()` is
  // NUL-terminated and stable until the sheet is mutated or destroyed.
  *out_utf8 = wb->workbook().sheet(index).name().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_add_sheet(fm_workbook_t* wb, const char* utf8_name) {
  clear_last_error();
  if (wb == nullptr || utf8_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_add_sheet: NULL argument");
  }
  wb->workbook().add_sheet(std::string(utf8_name));
  return 0;
}

extern "C" fm_status_t fm_workbook_move_sheet(fm_workbook_t* wb, uint32_t from_index, uint32_t to_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_move_sheet: wb is NULL");
  }
  auto r = wb->workbook().move_sheet(from_index, to_index);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_remove_sheet(fm_workbook_t* wb, uint32_t index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_remove_sheet: wb is NULL");
  }
  auto r = wb->workbook().remove_sheet(index);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_rename_sheet(fm_workbook_t* wb, uint32_t index, const char* new_name) {
  clear_last_error();
  if (wb == nullptr || new_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_rename_sheet: NULL argument");
  }
  auto r = wb->workbook().rename_sheet(index, std::string(new_name));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const char* name, const char* formula) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_defined_name: NULL argument");
  }
  auto r = wb->workbook().set_defined_name(std::string(name), std::string(formula));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Cell mutation
// ---------------------------------------------------------------------------

namespace {

// Validates a `(handle, sheet_index)` pair and returns the in-bounds
// status. On failure the diagnostic is already populated; callers should
// just propagate the returned code.
fm_status_t check_sheet_index(const fm_workbook_t* wb, std::size_t sheet_index, const char* fn) {
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, fn,
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  return 0;
}

}  // namespace

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

// ---------------------------------------------------------------------------
// Iteration / dump
// ---------------------------------------------------------------------------
//
// Sheets are stored row-sparse (`unordered_map<row, vector<Cell>>`); we
// surface a flat enumeration to bindings by materialising a sorted
// `(row, col)` index on demand and caching it on the handle. The cache
// is invalidated whenever any cell-mutating C API entry runs on the
// handle, so iteration always reflects the current store. The cache is
// purely an optimisation; correctness does not depend on it. We keep
// one cache per sheet to amortise sort cost across `cell_count_at`
// loops in the CLI's `dump` command.

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

extern "C" size_t fm_workbook_defined_name_count(const fm_workbook_t* wb) {
  if (wb == nullptr) {
    return 0;
  }
  return wb->workbook().defined_names().size();
}

extern "C" fm_status_t fm_workbook_defined_name_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                                   const char** out_formula) {
  clear_last_error();
  if (wb == nullptr || out_name == nullptr || out_formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_defined_name_at: NULL argument");
  }
  const auto& names = wb->workbook().defined_names();
  if (idx >= names.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_defined_name_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(names.size()));
  }
  *out_name = names[idx].name.c_str();
  *out_formula = names[idx].formula.c_str();
  return 0;
}

extern "C" size_t fm_workbook_table_count(const fm_workbook_t* wb) {
  if (wb == nullptr) {
    return 0;
  }
  return wb->workbook().tables().size();
}

extern "C" fm_status_t fm_workbook_table_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                            const char** out_display_name, const char** out_ref,
                                            size_t* out_sheet_index) {
  clear_last_error();
  if (wb == nullptr || out_name == nullptr || out_display_name == nullptr || out_ref == nullptr ||
      out_sheet_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_table_at: NULL argument");
  }
  const auto& tables = wb->workbook().tables();
  if (idx >= tables.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_table_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(tables.size()));
  }
  *out_name = tables[idx].name.c_str();
  *out_display_name = tables[idx].display_name.c_str();
  *out_ref = tables[idx].ref.c_str();
  *out_sheet_index = tables[idx].sheet_index;
  return 0;
}

extern "C" size_t fm_workbook_passthrough_count(const fm_workbook_t* wb) {
  if (wb == nullptr) {
    return 0;
  }
  return wb->workbook().passthrough_parts().size();
}

extern "C" fm_status_t fm_workbook_passthrough_at(const fm_workbook_t* wb, size_t idx, const char** out_path) {
  clear_last_error();
  if (wb == nullptr || out_path == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_passthrough_at: NULL argument");
  }
  const auto& parts = wb->workbook().passthrough_parts();
  if (idx >= parts.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_passthrough_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(parts.size()));
  }
  *out_path = parts[idx].path.c_str();
  return 0;
}

// ---------------------------------------------------------------------------
// Recalc
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_recalc(fm_workbook_t* wb) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_recalc: wb is NULL");
  }
  auto r = wb->workbook().recalc(formulon::eval::default_registry());
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_iterative(fm_workbook_t* wb, int32_t enabled, int32_t max_iterations,
                                                 double max_change) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_set_iterative: wb is NULL");
  }
  formulon::eval::IterativeOptions opts;
  opts.enabled = (enabled != 0);
  opts.max_iterations = max_iterations < 1 ? 1U : static_cast<std::uint32_t>(max_iterations);
  opts.max_change = max_change;
  wb->workbook().set_iterative_options(opts);
  return 0;
}

// ---------------------------------------------------------------------------
// Conditional formatting
// ---------------------------------------------------------------------------
//
// Bridges the engine-side `cf::evaluate_cf_for_range` walker into the
// stable C ABI. The opaque `fm_cf_results_t` owns the per-cell match
// list returned by the evaluator; index accessors copy the relevant
// fields out into the `fm_cf_match_t` POD so JS / Python / CLI consumers
// never see the internal `cf::CFMatch` layout.

// Opaque results handle. Owns the engine-produced match list verbatim so
// the address / index accessors can return data without re-walking the
// evaluator. The engine returns one `CFRangeCellMatches` entry per cell
// that produced at least one match.
struct fm_cf_results {
  std::vector<formulon::cf::CFRangeCellMatches> matches;
};

namespace {

// Translates an engine-side `cf::Color` into the C ABI's RGBA POD.
fm_cf_color_t to_c_color(formulon::cf::Color c) {
  fm_cf_color_t out{};
  out.r = c.r;
  out.g = c.g;
  out.b = c.b;
  out.a = c.a;
  return out;
}

// Projects a single `cf::CFMatch` into the wire-format `fm_cf_match_t`.
// The output is zero-initialised first so non-active fields are
// deterministic (the public ABI promises default-zero for all
// kind-irrelevant payload).
void fill_match(const formulon::cf::CFMatch& match, fm_cf_match_t* out) {
  *out = fm_cf_match_t{};
  out->priority = match.priority;
  switch (match.kind) {
    case formulon::cf::CFMatchKind::DifferentialFormat:
      out->kind = FM_CF_DIFFERENTIAL_FORMAT;
      if (match.dxf_id.has_value()) {
        out->dxf_id_engaged = 1;
        out->dxf_id = *match.dxf_id;
      }
      return;
    case formulon::cf::CFMatchKind::ColorScale:
      out->kind = FM_CF_COLOR_SCALE;
      if (match.resolved_fill_color.has_value()) {
        out->color = to_c_color(*match.resolved_fill_color);
      }
      return;
    case formulon::cf::CFMatchKind::DataBar:
      out->kind = FM_CF_DATA_BAR;
      if (match.data_bar_render.has_value()) {
        const auto& bar = *match.data_bar_render;
        out->bar_length_pct = bar.length_pct;
        out->bar_axis_position_pct = bar.axis_position_pct;
        out->bar_is_negative = bar.is_negative ? 1 : 0;
        out->bar_fill = to_c_color(bar.fill);
        if (bar.border.has_value()) {
          out->bar_border_engaged = 1;
          out->bar_border = to_c_color(*bar.border);
        }
        out->bar_gradient = bar.gradient ? 1 : 0;
      }
      return;
    case formulon::cf::CFMatchKind::IconSet:
      out->kind = FM_CF_ICON_SET;
      if (match.icon_render.has_value()) {
        out->icon_set_name = static_cast<std::int32_t>(match.icon_render->set_name);
        out->icon_index = match.icon_render->icon_index;
      }
      return;
  }
  // Defensive default — every enumerator above returns. If a future
  // kind enumerator is added without a corresponding case the output
  // surfaces as DifferentialFormat with no engaged fields.
  out->kind = FM_CF_DIFFERENTIAL_FORMAT;
}

}  // namespace

extern "C" fm_status_t fm_workbook_cf_evaluate_range(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                     uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                     double today_serial, fm_cf_results_t** out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_cf_evaluate_range: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_cf_evaluate_range: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  // Stack-allocated arena and eval context: the CF walker is purely
  // synchronous and the engine consumes both before returning. Binding
  // them here keeps the C ABI free of long-lived per-handle CF state.
  formulon::Arena arena;
  formulon::eval::EvalContext eval_ctx(wb->workbook().sheet(sheet_index));
  formulon::cf::CFHost host;
  host.arena = &arena;
  host.registry = &formulon::eval::default_registry();
  host.eval_ctx = &eval_ctx;
  host.today_serial = std::isnan(today_serial) ? std::optional<double>{} : std::optional<double>{today_serial};

  formulon::cf::CFCellRange range{};
  range.first = formulon::CellAddress{first_row, first_col};
  range.last = formulon::CellAddress{last_row, last_col};

  auto matches = formulon::cf::evaluate_cf_for_range(wb->workbook().sheet(sheet_index), range, host);

  auto handle = std::unique_ptr<fm_cf_results_t>(new fm_cf_results_t{});
  handle->matches = std::move(matches);
  *out = handle.release();
  return 0;
}

extern "C" void fm_cf_results_destroy(fm_cf_results_t* results) {
  // Mirrors `free(NULL)` semantics.
  delete results;
}

extern "C" size_t fm_cf_results_cell_count(const fm_cf_results_t* results) {
  if (results == nullptr) {
    return 0;
  }
  return results->matches.size();
}

extern "C" fm_status_t fm_cf_results_cell_at(const fm_cf_results_t* results, size_t cell_idx, uint32_t* out_row,
                                             uint32_t* out_col, size_t* out_match_count) {
  clear_last_error();
  if (results == nullptr || out_row == nullptr || out_col == nullptr || out_match_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cf_results_cell_at: NULL argument");
  }
  if (cell_idx >= results->matches.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_cf_results_cell_at: cell_idx out of range",
        "cell_idx=" + std::to_string(cell_idx) + " cell_count=" + std::to_string(results->matches.size()));
  }
  const auto& entry = results->matches[cell_idx];
  *out_row = entry.cell.row;
  *out_col = entry.cell.col;
  *out_match_count = entry.matches.size();
  return 0;
}

extern "C" fm_status_t fm_cf_results_match_at(const fm_cf_results_t* results, size_t cell_idx, size_t match_idx,
                                              fm_cf_match_t* out) {
  clear_last_error();
  if (results == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cf_results_match_at: NULL argument");
  }
  if (cell_idx >= results->matches.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_cf_results_match_at: cell_idx out of range",
        "cell_idx=" + std::to_string(cell_idx) + " cell_count=" + std::to_string(results->matches.size()));
  }
  const auto& entry = results->matches[cell_idx];
  if (match_idx >= entry.matches.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_cf_results_match_at: match_idx out of range",
        "match_idx=" + std::to_string(match_idx) + " match_count=" + std::to_string(entry.matches.size()));
  }
  fill_match(entry.matches[match_idx], out);
  return 0;
}

// ---------------------------------------------------------------------------
// Sheet view / layout
// ---------------------------------------------------------------------------
//
// Bridges `Sheet::view()` / `Sheet::layout()` over the stable C ABI.
// The `set_*` mutators upsert into the underlying `SheetLayout` vectors
// via small private helpers so concurrent column-span edits stay
// internally consistent (e.g. setting the width on a span that
// overlaps an existing entry replaces only the overlapping portion).

namespace {

// Splits any pre-existing column entries that intersect `[first, last]`
// so the resulting `columns` vector contains at most one entry whose
// span equals `[first, last]`. Pre-existing fields that fall outside
// the requested range are preserved on the residual entry. Returns a
// reference to the entry whose span is `[first, last]`; the caller
// then writes the field it wants to update on that entry.
formulon::ColumnLayout& upsert_column_span(formulon::SheetLayout& layout, std::uint32_t first, std::uint32_t last) {
  // First pass: split any entry that overlaps the target span. We
  // copy non-overlapping residuals into a fresh vector so the caller
  // sees a consistent state regardless of how many splits happened.
  std::vector<formulon::ColumnLayout> next;
  next.reserve(layout.columns.size() + 2);
  for (const formulon::ColumnLayout& entry : layout.columns) {
    // No overlap -> retain verbatim.
    if (entry.last < first || entry.first > last) {
      next.push_back(entry);
      continue;
    }
    // Left residual (entry.first .. first-1).
    if (entry.first < first) {
      formulon::ColumnLayout left = entry;
      left.last = first - 1U;
      next.push_back(left);
    }
    // Right residual (last+1 .. entry.last).
    if (entry.last > last) {
      formulon::ColumnLayout right = entry;
      right.first = last + 1U;
      next.push_back(right);
    }
    // The middle slice `[max(entry.first,first), min(entry.last,last)]`
    // is dropped here; we re-create one canonical entry below so the
    // caller can write its field of interest atomically.
  }
  formulon::ColumnLayout target;
  target.first = first;
  target.last = last;
  next.push_back(target);
  layout.columns = std::move(next);
  // The freshly inserted entry is the last element by construction.
  return layout.columns.back();
}

// Returns a pointer to the row override whose `row` equals `row`,
// creating one if none exists. The returned pointer is valid until
// the next mutation of `layout.row_overrides`.
formulon::RowLayout* upsert_row_override(formulon::SheetLayout& layout, std::uint32_t row) {
  for (formulon::RowLayout& entry : layout.row_overrides) {
    if (entry.row == row) {
      return &entry;
    }
  }
  formulon::RowLayout fresh;
  fresh.row = row;
  layout.row_overrides.push_back(fresh);
  return &layout.row_overrides.back();
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_column_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_column_count: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_column_count: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  *out_count = wb->workbook().sheet(sheet_index).layout().columns.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_column(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                           fm_column_layout_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_column: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_column: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  const auto& cols = wb->workbook().sheet(sheet_index).layout().columns;
  if (idx >= cols.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_column: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(cols.size()));
  }
  *out = fm_column_layout_t{};
  out->first = cols[idx].first;
  out->last = cols[idx].last;
  out->width = cols[idx].width;
  out->hidden = cols[idx].hidden ? 1 : 0;
  out->outline_level = cols[idx].outline_level;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_override_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_row_override_count: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_row_override_count: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  *out_count = wb->workbook().sheet(sheet_index).layout().row_overrides.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_override(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                 fm_row_layout_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_row_override: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_row_override: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  const auto& rows = wb->workbook().sheet(sheet_index).layout().row_overrides;
  if (idx >= rows.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_row_override: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(rows.size()));
  }
  *out = fm_row_layout_t{};
  out->row = rows[idx].row;
  out->height = rows[idx].height;
  out->hidden = rows[idx].hidden ? 1 : 0;
  out->outline_level = rows[idx].outline_level;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_view: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_view: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  const formulon::SheetView& v = wb->workbook().sheet(sheet_index).view();
  out->zoom_scale = v.zoom_scale;
  out->freeze_rows = v.freeze_rows;
  out->freeze_cols = v.freeze_cols;
  out->tab_hidden = v.tab_hidden ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                 double width) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_column_width: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_width: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  if (last < first) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_width: last < first",
                             "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  formulon::ColumnLayout& entry = upsert_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last);
  entry.width = width;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                  int32_t hidden) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_set_column_hidden: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_hidden: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  if (last < first) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_hidden: last < first",
                             "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  formulon::ColumnLayout& entry = upsert_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last);
  entry.hidden = (hidden != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                   uint8_t level) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_set_column_outline: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_outline: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  if (last < first) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_column_outline: last < first",
                             "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  formulon::ColumnLayout& entry = upsert_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last);
  entry.outline_level = level;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_row_height: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_row_height: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  formulon::RowLayout* entry = upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  entry->height = height;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t hidden) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_row_hidden: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_row_hidden: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  formulon::RowLayout* entry = upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  entry->hidden = (hidden != 0);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint8_t level) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_row_outline: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_row_outline: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  formulon::RowLayout* entry = upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  entry->outline_level = level;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, size_t sheet_index, uint32_t zoom_scale) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_zoom: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_zoom: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  // Clamp to the OOXML-valid `[10, 400]` interval; out-of-range values
  // are rounded to the nearest endpoint rather than rejected so JS
  // callers do not have to mirror the bound.
  std::uint32_t clamped = zoom_scale;
  if (clamped < 10U) {
    clamped = 10U;
  } else if (clamped > 400U) {
    clamped = 400U;
  }
  wb->workbook().sheet(sheet_index).mutable_view().zoom_scale = clamped;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_freeze(fm_workbook_t* wb, size_t sheet_index, uint32_t freeze_rows,
                                           uint32_t freeze_cols) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_freeze: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_freeze: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  view.freeze_rows = freeze_rows;
  view.freeze_cols = freeze_cols;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_tab_hidden(fm_workbook_t* wb, size_t sheet_index, int32_t hidden) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_tab_hidden: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_set_tab_hidden: sheet_index out of range",
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  wb->workbook().sheet(sheet_index).mutable_view().tab_hidden = (hidden != 0);
  return 0;
}

// ---------------------------------------------------------------------------
// Diagnostics
// ---------------------------------------------------------------------------

extern "C" const char* fm_last_error_message(void) {
  return g_last_error_message.c_str();
}

extern "C" const char* fm_last_error_context(void) {
  return g_last_error_context.c_str();
}

extern "C" const char* fm_status_string(fm_status_t status) {
  return formulon::to_cstring(static_cast<formulon::FormulonErrorCode>(status));
}

// ---------------------------------------------------------------------------
// Version
// ---------------------------------------------------------------------------

#ifndef FORMULON_VERSION_STRING
#define FORMULON_VERSION_STRING "0.0.0+phase3"
#endif

extern "C" const char* fm_version_string(void) {
  return FORMULON_VERSION_STRING;
}
