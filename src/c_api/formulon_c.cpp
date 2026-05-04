// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the stable C ABI declared in `formulon_c.h`.
//
// Design notes:
//
//   * The opaque `fm_workbook_t` is a thin C++ struct that owns a single
//     `Workbook`. The workbook itself owns the text-storage deque that
//     backs every `Value::text` view, so the load path simply moves the
//     workbook out of the `io::OoxmlReadResult`; the cell-text views
//     remain valid for the workbook's lifetime.
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
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_evaluator.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "eval/dep_graph.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/lambda_format.h"
#include "eval/lambda_value.h"
#include "eval/recalc_engine.h"
#include "io/ooxml_reader.h"
#include "io/styles_reader.h"
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
// Holds a single `Workbook`. The workbook itself owns the text-storage
// deque that backs every `Value::text` view, so the load path no longer
// needs to keep an `OoxmlReadResult` alive beyond the call that
// produced it — the workbook is move-constructed out of the result and
// retains the storage. (Earlier slices kept the result alive purely
// for that reason; the dual-slot layout is no longer needed.)
struct fm_workbook {
  std::optional<formulon::Workbook> wb;

  // Storage for UTF-8 strings owned by this handle: cell text inputs
  // (`fm_workbook_set_text`) and read-side NUL-terminated copies
  // (`fm_workbook_get_value` / `fm_workbook_sheet_name` for sheets that
  // would otherwise alias inline storage).
  TextStore text_store;

  formulon::Workbook& workbook() { return *wb; }
  const formulon::Workbook& workbook() const { return *wb; }
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
  handle->wb.emplace(formulon::Workbook::create());
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_workbook_create_empty(fm_workbook_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_create_empty: out is NULL");
  }
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->wb.emplace(formulon::Workbook::create_empty());
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
  // The workbook now owns the text-storage deque that backs every
  // Text-cell `string_view`, so we move only the workbook out of the
  // result; the rest of `OoxmlReadResult` (passthrough parts mirror,
  // audit counter) is discarded.
  handle->wb.emplace(std::move(result.value().workbook));
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

extern "C" fm_status_t fm_workbook_insert_rows(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_insert_rows: NULL argument");
  }
  auto r = wb->workbook().insert_rows(sheet, row, count);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_delete_rows(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_delete_rows: NULL argument");
  }
  auto r = wb->workbook().delete_rows(sheet, row, count);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_insert_cols(fm_workbook_t* wb, uint32_t sheet, uint32_t col, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_insert_cols: NULL argument");
  }
  auto r = wb->workbook().insert_cols(sheet, col, count);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_delete_cols(fm_workbook_t* wb, uint32_t sheet, uint32_t col, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_delete_cols: NULL argument");
  }
  auto r = wb->workbook().delete_cols(sheet, col, count);
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
// Sheet UI features (merges, hyperlinks, comments)
// ---------------------------------------------------------------------------

namespace {

// Bounds-check for the uint32-typed sheet index used by the sheet UI
// surface. Mirrors `check_sheet_index` but accepts the ABI's uint32.
fm_status_t check_sheet_u32(fm_workbook_t* wb, std::uint32_t sheet, const char* fn) {
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  if (static_cast<std::size_t>(sheet) >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, fn,
        "sheet=" + std::to_string(sheet) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_add_hyperlink(fm_workbook_t* wb, std::uint32_t sheet, fm_hyperlink hl) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_add_hyperlink"); rc != 0) {
    return rc;
  }
  formulon::Hyperlink out;
  out.row = hl.row;
  out.col = hl.col;
  out.target = (hl.target != nullptr) ? std::string(hl.target) : std::string();
  out.location = (hl.location != nullptr) ? std::string(hl.location) : std::string();
  out.display = (hl.display != nullptr) ? std::string(hl.display) : std::string();
  out.tooltip = (hl.tooltip != nullptr) ? std::string(hl.tooltip) : std::string();
  // rid stays empty; the writer mints a fresh rIdN on save.
  wb->workbook().sheet(sheet).mutable_hyperlinks().push_back(std::move(out));
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_hyperlink(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                                 std::uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_hyperlink"); rc != 0) {
    return rc;
  }
  auto& hls = wb->workbook().sheet(sheet).mutable_hyperlinks();
  hls.erase(std::remove_if(hls.begin(), hls.end(),
                           [&](const formulon::Hyperlink& h) { return h.row == row && h.col == col; }),
            hls.end());
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_hyperlink_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_hyperlink_at"); rc != 0) {
    return rc;
  }
  auto& hls = wb->workbook().sheet(sheet).mutable_hyperlinks();
  if (static_cast<std::size_t>(index) >= hls.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_remove_hyperlink_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(hls.size()));
  }
  hls.erase(hls.begin() + static_cast<std::ptrdiff_t>(index));
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_hyperlinks(fm_workbook_t* wb, std::uint32_t sheet) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_clear_hyperlinks"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet).mutable_hyperlinks().clear();
  return 0;
}

extern "C" fm_status_t fm_sheet_add_merge(fm_workbook_t* wb, std::uint32_t sheet, fm_merge_range merge) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_add_merge"); rc != 0) {
    return rc;
  }
  // Normalise corners so first <= last componentwise; mirrors the
  // reader's behaviour and keeps downstream consumers simple.
  formulon::MergeRange m;
  m.first_row = (merge.first_row < merge.last_row) ? merge.first_row : merge.last_row;
  m.first_col = (merge.first_col < merge.last_col) ? merge.first_col : merge.last_col;
  m.last_row = (merge.first_row < merge.last_row) ? merge.last_row : merge.first_row;
  m.last_col = (merge.first_col < merge.last_col) ? merge.last_col : merge.first_col;
  wb->workbook().sheet(sheet).mutable_merges().push_back(m);
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_merge(fm_workbook_t* wb, std::uint32_t sheet, fm_merge_range range) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_merge"); rc != 0) {
    return rc;
  }
  // Normalise corners so first <= last componentwise; mirrors fm_sheet_add_merge.
  formulon::MergeRange q;
  q.first_row = (range.first_row < range.last_row) ? range.first_row : range.last_row;
  q.first_col = (range.first_col < range.last_col) ? range.first_col : range.last_col;
  q.last_row = (range.first_row < range.last_row) ? range.last_row : range.first_row;
  q.last_col = (range.first_col < range.last_col) ? range.last_col : range.first_col;
  auto& merges = wb->workbook().sheet(sheet).mutable_merges();
  merges.erase(std::remove_if(merges.begin(), merges.end(),
                              [&](const formulon::MergeRange& a) {
                                return !(a.last_row < q.first_row || q.last_row < a.first_row ||
                                         a.last_col < q.first_col || q.last_col < a.first_col);
                              }),
               merges.end());
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_merge_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_merge_at"); rc != 0) {
    return rc;
  }
  auto& merges = wb->workbook().sheet(sheet).mutable_merges();
  if (static_cast<std::size_t>(index) >= merges.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_remove_merge_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(merges.size()));
  }
  merges.erase(merges.begin() + static_cast<std::ptrdiff_t>(index));
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_merges(fm_workbook_t* wb, std::uint32_t sheet) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_clear_merges"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet).mutable_merges().clear();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_comment_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                               std::uint32_t col, fm_comment* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_comment_at: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_comment_at"); rc != 0) {
    return rc;
  }
  for (const formulon::CellComment& c : wb->workbook().sheet(sheet).comments()) {
    if (c.row == row && c.col == col) {
      out->row = c.row;
      out->col = c.col;
      out->author = c.author.c_str();
      out->text = c.text.c_str();
      return 0;
    }
  }
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_comment_at: no comment at cell",
                           "row=" + std::to_string(row) + " col=" + std::to_string(col));
}

extern "C" fm_status_t fm_sheet_get_hyperlink_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index,
                                                 fm_hyperlink* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_hyperlink_at: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_hyperlink_at"); rc != 0) {
    return rc;
  }
  const auto& hls = wb->workbook().sheet(sheet).hyperlinks();
  if (static_cast<std::size_t>(index) >= hls.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_hyperlink_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(hls.size()));
  }
  const formulon::Hyperlink& h = hls[index];
  out->row = h.row;
  out->col = h.col;
  out->target = h.target.c_str();
  out->location = h.location.c_str();
  out->display = h.display.c_str();
  out->tooltip = h.tooltip.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_hyperlink_count(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_hyperlink_count: out_count is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_hyperlink_count"); rc != 0) {
    return rc;
  }
  *out_count = static_cast<std::uint32_t>(wb->workbook().sheet(sheet).hyperlinks().size());
  return 0;
}

extern "C" fm_status_t fm_sheet_get_merge_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index,
                                             fm_merge_range* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_merge_at: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_merge_at"); rc != 0) {
    return rc;
  }
  const auto& merges = wb->workbook().sheet(sheet).merges();
  if (static_cast<std::size_t>(index) >= merges.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_merge_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(merges.size()));
  }
  const formulon::MergeRange& m = merges[index];
  out->first_row = m.first_row;
  out->first_col = m.first_col;
  out->last_row = m.last_row;
  out->last_col = m.last_col;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_merge_count(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_merge_count: out_count is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_merge_count"); rc != 0) {
    return rc;
  }
  *out_count = static_cast<std::uint32_t>(wb->workbook().sheet(sheet).merges().size());
  return 0;
}

extern "C" fm_status_t fm_sheet_set_comment(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                            std::uint32_t col, const char* author, const char* text) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_set_comment"); rc != 0) {
    return rc;
  }
  auto& list = wb->workbook().sheet(sheet).mutable_comments();
  // Locate any existing entry first; mutating the list invalidates
  // iterators so we capture the index instead of an iterator.
  std::size_t found = list.size();
  for (std::size_t i = 0; i < list.size(); ++i) {
    if (list[i].row == row && list[i].col == col) {
      found = i;
      break;
    }
  }
  // Empty/NULL text => removal.
  if (text == nullptr || text[0] == '\0') {
    if (found < list.size()) {
      list.erase(list.begin() + static_cast<std::ptrdiff_t>(found));
    }
    return 0;
  }
  formulon::CellComment c;
  c.row = row;
  c.col = col;
  c.author = (author != nullptr) ? std::string(author) : std::string();
  c.text = std::string(text);
  if (found < list.size()) {
    list[found] = std::move(c);
  } else {
    list.push_back(std::move(c));
  }
  return 0;
}

// `fm_merge_range` and `formulon::MergeRange` must remain layout-compatible
// so `fm_sheet_get_validation_at` can hand out the engine-side
// `std::vector<MergeRange>::data()` directly via `reinterpret_cast`. Both
// types are POD with four `uint32_t` fields in identical order; the static
// asserts below pin that contract so any future change to either struct
// trips a compile-time error.
static_assert(sizeof(fm_merge_range) == sizeof(formulon::MergeRange),
              "fm_merge_range / MergeRange size mismatch breaks validation getter");
static_assert(alignof(fm_merge_range) == alignof(formulon::MergeRange),
              "fm_merge_range / MergeRange alignment mismatch breaks validation getter");
static_assert(offsetof(fm_merge_range, first_row) == offsetof(formulon::MergeRange, first_row),
              "fm_merge_range::first_row layout mismatch");
static_assert(offsetof(fm_merge_range, first_col) == offsetof(formulon::MergeRange, first_col),
              "fm_merge_range::first_col layout mismatch");
static_assert(offsetof(fm_merge_range, last_row) == offsetof(formulon::MergeRange, last_row),
              "fm_merge_range::last_row layout mismatch");
static_assert(offsetof(fm_merge_range, last_col) == offsetof(formulon::MergeRange, last_col),
              "fm_merge_range::last_col layout mismatch");

extern "C" fm_status_t fm_sheet_get_validation_count(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_validation_count: out_count is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_validation_count"); rc != 0) {
    return rc;
  }
  *out_count = static_cast<std::uint32_t>(wb->workbook().sheet(sheet).validations().size());
  return 0;
}

extern "C" fm_status_t fm_sheet_get_validation_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index,
                                                  fm_data_validation* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_validation_at: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_validation_at"); rc != 0) {
    return rc;
  }
  const auto& list = wb->workbook().sheet(sheet).validations();
  if (static_cast<std::size_t>(index) >= list.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_validation_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(list.size()));
  }
  const formulon::DataValidation& v = list[index];
  // Ranges are layout-compatible (see static_asserts above), so we hand
  // out the engine-side vector buffer directly. The lifetime contract
  // documented in the header (valid until the next mutation that
  // touches the validation list) follows naturally from
  // `std::vector::data()` invalidation rules.
  out->ranges = v.ranges.empty() ? nullptr : reinterpret_cast<const fm_merge_range*>(v.ranges.data());
  out->range_count = static_cast<std::uint32_t>(v.ranges.size());
  out->type = v.type;
  out->op = v.op;
  out->error_style = v.error_style;
  out->allow_blank = v.allow_blank ? 1 : 0;
  out->show_input_message = v.show_input_message ? 1 : 0;
  out->show_error_message = v.show_error_message ? 1 : 0;
  out->formula1 = v.formula1.c_str();
  out->formula2 = v.formula2.c_str();
  out->error_title = v.error_title.c_str();
  out->error_message = v.error_message.c_str();
  out->prompt_title = v.prompt_title.c_str();
  out->prompt_message = v.prompt_message.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_add_validation(fm_workbook_t* wb, std::uint32_t sheet, fm_data_validation v) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_add_validation"); rc != 0) {
    return rc;
  }
  if (v.range_count > 0 && v.ranges == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_add_validation: ranges is NULL but range_count > 0",
                             "range_count=" + std::to_string(v.range_count));
  }
  formulon::DataValidation out;
  out.ranges.reserve(v.range_count);
  for (std::uint32_t i = 0; i < v.range_count; ++i) {
    formulon::MergeRange r;
    r.first_row = v.ranges[i].first_row;
    r.first_col = v.ranges[i].first_col;
    r.last_row = v.ranges[i].last_row;
    r.last_col = v.ranges[i].last_col;
    out.ranges.push_back(r);
  }
  out.type = v.type;
  out.op = v.op;
  out.error_style = v.error_style;
  out.allow_blank = v.allow_blank != 0;
  out.show_input_message = v.show_input_message != 0;
  out.show_error_message = v.show_error_message != 0;
  out.formula1 = (v.formula1 != nullptr) ? std::string(v.formula1) : std::string();
  out.formula2 = (v.formula2 != nullptr) ? std::string(v.formula2) : std::string();
  out.error_title = (v.error_title != nullptr) ? std::string(v.error_title) : std::string();
  out.error_message = (v.error_message != nullptr) ? std::string(v.error_message) : std::string();
  out.prompt_title = (v.prompt_title != nullptr) ? std::string(v.prompt_title) : std::string();
  out.prompt_message = (v.prompt_message != nullptr) ? std::string(v.prompt_message) : std::string();
  wb->workbook().sheet(sheet).mutable_validations().push_back(std::move(out));
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_validation_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_validation_at"); rc != 0) {
    return rc;
  }
  auto& list = wb->workbook().sheet(sheet).mutable_validations();
  if (static_cast<std::size_t>(index) >= list.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_remove_validation_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(list.size()));
  }
  list.erase(list.begin() + static_cast<std::ptrdiff_t>(index));
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_validations(fm_workbook_t* wb, std::uint32_t sheet) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_clear_validations"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet).mutable_validations().clear();
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

extern "C" fm_status_t fm_workbook_calc_mode(const fm_workbook_t* wb, fm_calc_mode_t* out_mode) {
  clear_last_error();
  if (wb == nullptr || out_mode == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_calc_mode: NULL argument");
  }
  switch (wb->workbook().calc_mode()) {
    case formulon::Workbook::CalcMode::kAuto:
      *out_mode = FM_CALC_MODE_AUTO;
      break;
    case formulon::Workbook::CalcMode::kManual:
      *out_mode = FM_CALC_MODE_MANUAL;
      break;
    case formulon::Workbook::CalcMode::kAutoNoTable:
      *out_mode = FM_CALC_MODE_AUTO_NO_TABLE;
      break;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_calc_mode(fm_workbook_t* wb, fm_calc_mode_t mode) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_set_calc_mode: wb is NULL");
  }
  formulon::Workbook::CalcMode resolved = formulon::Workbook::CalcMode::kAuto;
  switch (mode) {
    case FM_CALC_MODE_AUTO:
      resolved = formulon::Workbook::CalcMode::kAuto;
      break;
    case FM_CALC_MODE_MANUAL:
      resolved = formulon::Workbook::CalcMode::kManual;
      break;
    case FM_CALC_MODE_AUTO_NO_TABLE:
      resolved = formulon::Workbook::CalcMode::kAutoNoTable;
      break;
    default:
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_set_calc_mode: unknown mode");
  }
  wb->workbook().set_calc_mode(resolved);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_protection(const fm_workbook_t* wb, uint32_t sheet_index,
                                               fm_sheet_protection_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_protection: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_protection: sheet_index out of range");
  }
  const formulon::SheetProtection& p = wb->workbook().sheet(sheet_index).protection();
  out->enabled = p.enabled ? 1 : 0;
  out->algorithm_name = p.algorithm_name.c_str();
  out->hash_value = p.hash_value.c_str();
  out->salt_value = p.salt_value.c_str();
  out->spin_count = p.spin_count;
  out->legacy_password = p.legacy_password.c_str();
  out->sheet = p.sheet ? 1 : 0;
  out->objects = p.objects ? 1 : 0;
  out->scenarios = p.scenarios ? 1 : 0;
  out->format_cells = p.format_cells ? 1 : 0;
  out->format_columns = p.format_columns ? 1 : 0;
  out->format_rows = p.format_rows ? 1 : 0;
  out->insert_columns = p.insert_columns ? 1 : 0;
  out->insert_rows = p.insert_rows ? 1 : 0;
  out->insert_hyperlinks = p.insert_hyperlinks ? 1 : 0;
  out->delete_columns = p.delete_columns ? 1 : 0;
  out->delete_rows = p.delete_rows ? 1 : 0;
  out->select_locked_cells = p.select_locked_cells ? 1 : 0;
  out->select_unlocked_cells = p.select_unlocked_cells ? 1 : 0;
  out->sort = p.sort ? 1 : 0;
  out->auto_filter = p.auto_filter ? 1 : 0;
  out->pivot_tables = p.pivot_tables ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_protection(fm_workbook_t* wb, uint32_t sheet_index,
                                               const fm_sheet_protection_t* in) {
  clear_last_error();
  if (wb == nullptr || in == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_set_protection: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_set_protection: sheet_index out of range");
  }
  formulon::SheetProtection& p = wb->workbook().sheet(sheet_index).mutable_protection();
  // Helper for pointer→string deep copy with NULL-as-empty semantics.
  const auto copy_or_empty = [](const char* s) -> std::string { return s == nullptr ? std::string() : std::string(s); };
  p.enabled = in->enabled != 0;
  p.algorithm_name = copy_or_empty(in->algorithm_name);
  p.hash_value = copy_or_empty(in->hash_value);
  p.salt_value = copy_or_empty(in->salt_value);
  p.spin_count = in->spin_count;
  p.legacy_password = copy_or_empty(in->legacy_password);
  p.sheet = in->sheet != 0;
  p.objects = in->objects != 0;
  p.scenarios = in->scenarios != 0;
  p.format_cells = in->format_cells != 0;
  p.format_columns = in->format_columns != 0;
  p.format_rows = in->format_rows != 0;
  p.insert_columns = in->insert_columns != 0;
  p.insert_rows = in->insert_rows != 0;
  p.insert_hyperlinks = in->insert_hyperlinks != 0;
  p.delete_columns = in->delete_columns != 0;
  p.delete_rows = in->delete_rows != 0;
  p.select_locked_cells = in->select_locked_cells != 0;
  p.select_unlocked_cells = in->select_unlocked_cells != 0;
  p.sort = in->sort != 0;
  p.auto_filter = in->auto_filter != 0;
  p.pivot_tables = in->pivot_tables != 0;
  return 0;
}

extern "C" fm_status_t fm_workbook_partial_recalc(fm_workbook_t* wb, const fm_viewport* viewport,
                                                  uint32_t* out_recomputed_count) {
  clear_last_error();
  if (wb == nullptr || viewport == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_partial_recalc: NULL argument");
  }
  formulon::eval::SheetCellRange range;
  range.sheet_id = static_cast<std::uint16_t>(viewport->sheet);
  range.first_row = viewport->first_row;
  range.last_row = viewport->last_row;
  range.first_col = viewport->first_col;
  range.last_col = viewport->last_col;
  auto r = wb->workbook().partial_recalc(formulon::eval::default_registry(), range);
  if (!r) {
    return set_last_error(r.error());
  }
  if (out_recomputed_count != nullptr) {
    *out_recomputed_count = r.value().cells_evaluated;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_iterative_progress(fm_workbook_t* wb, fm_iterative_progress_cb cb,
                                                          void* user_data) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_iterative_progress: wb is NULL");
  }
  // The C ABI callback signature
  //   `bool(*)(uint32_t, double, uint32_t, void*)`
  // is bit-identical to the engine's `IterativeProgressCb` typedef, so
  // a direct assignment is well-defined under both C and C++ rules.
  formulon::eval::IterativeProgressCb engine_cb = cb;
  wb->workbook().set_iterative_progress(engine_cb, user_data);
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
// Conditional formatting — mutation
// ---------------------------------------------------------------------------
//
// Each rule lives inside a `<conditionalFormatting>` block. The mutation
// API exposes a flattened, per-rule view: `fm_sheet_cf_count` totals
// rules across all blocks, `fm_sheet_cf_get_at` reaches into the
// containing block to surface its sqref. Adds always create a fresh
// single-rule block to keep insertion order deterministic and avoid
// merging with semantically distinct sqref unions; removes prune the
// block too when its rule list goes empty.

// `fm_cf_cell_range_t` mirrors `cf::CFCellRange` so the OOXML reader's
// pre-allocated vector buffer can be handed back as a borrowed
// `const fm_cf_cell_range_t*` view without per-call repacking.
static_assert(sizeof(fm_cf_cell_range_t) == sizeof(formulon::cf::CFCellRange),
              "fm_cf_cell_range_t / cf::CFCellRange size mismatch");
static_assert(alignof(fm_cf_cell_range_t) == alignof(formulon::cf::CFCellRange),
              "fm_cf_cell_range_t / cf::CFCellRange align mismatch");
static_assert(offsetof(fm_cf_cell_range_t, first_row) == offsetof(formulon::cf::CFCellRange, first),
              "fm_cf_cell_range_t::first_row layout mismatch");
static_assert(offsetof(fm_cf_cell_range_t, last_row) == offsetof(formulon::cf::CFCellRange, last),
              "fm_cf_cell_range_t::last_row layout mismatch");

namespace {

// Returns `true` for the three visual rule types whose payloads
// (color_scale / data_bar / icon_set sub-specs) are not yet creatable
// through the C ABI. The OOXML reader / writer still round-trip them
// verbatim — this gate only fires on the mutation entry point.
bool is_visual_rule_type(std::uint8_t type) {
  switch (static_cast<formulon::cf::RuleType>(type)) {
    case formulon::cf::RuleType::ColorScale:
    case formulon::cf::RuleType::DataBar:
    case formulon::cf::RuleType::IconSet:
      return true;
    default:
      return false;
  }
}

// Walks the sheet's `conditional_formats` vector and resolves the
// `flat_idx`-th rule into the (block_idx, rule_idx) pair. Returns
// `true` on success; `false` when `flat_idx` is past the end.
bool resolve_flat_index(const std::vector<formulon::cf::ConditionalFormat>& blocks, std::size_t flat_idx,
                        std::size_t* out_block, std::size_t* out_rule) {
  std::size_t cursor = 0;
  for (std::size_t b = 0; b < blocks.size(); ++b) {
    const auto& block = blocks[b];
    if (flat_idx < cursor + block.rules.size()) {
      *out_block = b;
      *out_rule = flat_idx - cursor;
      return true;
    }
    cursor += block.rules.size();
  }
  return false;
}

// Materialises a `cf::CFRule` view onto the wire-format `fm_cf_rule_t`.
// All non-engaged variant fields are zero-initialised first so callers
// observe deterministic defaults. String views borrow the engine's
// storage; the contract documented in the header is "valid until the
// next CF mutation".
void fill_rule(const formulon::cf::ConditionalFormat& block, const formulon::cf::CFRule& rule, fm_cf_rule_t* out) {
  *out = fm_cf_rule_t{};
  out->id = rule.id.c_str();
  out->type = static_cast<std::uint8_t>(rule.type);
  out->priority = rule.priority;
  out->stop_if_true = rule.stop_if_true ? 1 : 0;
  if (rule.dxf_id.has_value()) {
    out->dxf_id_engaged = 1;
    out->dxf_id = *rule.dxf_id;
  }
  out->sqref = block.sqref.empty() ? nullptr : reinterpret_cast<const fm_cf_cell_range_t*>(block.sqref.data());
  out->sqref_count = static_cast<std::uint32_t>(block.sqref.size());
  out->formula1 = rule.formula1.has_value() ? rule.formula1->c_str() : nullptr;
  out->formula2 = rule.formula2.has_value() ? rule.formula2->c_str() : nullptr;
  if (rule.op.has_value()) {
    out->op_engaged = 1;
    out->op = static_cast<std::uint8_t>(*rule.op);
  }
  if (rule.rank.has_value()) {
    out->rank_engaged = 1;
    out->rank = *rule.rank;
  }
  out->percent = rule.percent ? 1 : 0;
  out->bottom = rule.bottom ? 1 : 0;
  out->above_average = rule.above_average ? 1 : 0;
  out->equal_average = rule.equal_average ? 1 : 0;
  if (rule.std_dev.has_value()) {
    out->std_dev_engaged = 1;
    out->std_dev = *rule.std_dev;
  }
  out->text = rule.text.has_value() ? rule.text->c_str() : nullptr;
  if (rule.time_period.has_value()) {
    out->time_period_engaged = 1;
    out->time_period = static_cast<std::uint8_t>(*rule.time_period);
  }
}

}  // namespace

extern "C" fm_status_t fm_sheet_cf_count(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_count: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_count: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  std::size_t total = 0;
  for (const auto& block : wb->workbook().sheet(sheet_index).conditional_formats()) {
    total += block.rules.size();
  }
  *out_count = total;
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_get_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t idx,
                                          fm_cf_rule_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_get_at: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_get_at: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  const auto& blocks = wb->workbook().sheet(sheet_index).conditional_formats();
  std::size_t b = 0;
  std::size_t r = 0;
  if (!resolve_flat_index(blocks, idx, &b, &r)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_cf_get_at: idx out of range",
                             "idx=" + std::to_string(idx));
  }
  fill_rule(blocks[b], blocks[b].rules[r], out);
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_add_rule(fm_workbook_t* wb, std::size_t sheet_index, fm_cf_rule_t rule) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_add_rule: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  if (rule.sqref_count == 0) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: sqref_count must be >= 1");
  }
  if (rule.sqref == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_cf_add_rule: sqref is NULL while sqref_count > 0");
  }
  if (is_visual_rule_type(rule.type)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: visual rule types (ColorScale/DataBar/IconSet) "
                             "are not creatable through this API",
                             "type=" + std::to_string(rule.type));
  }
  if (rule.type > static_cast<std::uint8_t>(formulon::cf::RuleType::UniqueValues)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_cf_add_rule: unknown rule type",
                             "type=" + std::to_string(rule.type));
  }

  formulon::cf::ConditionalFormat new_block;
  new_block.sqref.reserve(rule.sqref_count);
  for (std::uint32_t i = 0; i < rule.sqref_count; ++i) {
    formulon::cf::CFCellRange r;
    r.first = formulon::CellAddress{rule.sqref[i].first_row, rule.sqref[i].first_col};
    r.last = formulon::CellAddress{rule.sqref[i].last_row, rule.sqref[i].last_col};
    new_block.sqref.push_back(r);
  }

  formulon::cf::CFRule out_rule;
  out_rule.type = static_cast<formulon::cf::RuleType>(rule.type);

  // Auto-assign priority to (max_existing + 1) when caller passed <= 0.
  std::int32_t max_priority = 0;
  for (const auto& block : wb->workbook().sheet(sheet_index).conditional_formats()) {
    for (const auto& existing : block.rules) {
      if (existing.priority > max_priority) {
        max_priority = existing.priority;
      }
    }
  }
  out_rule.priority = (rule.priority > 0) ? rule.priority : (max_priority + 1);
  out_rule.stop_if_true = rule.stop_if_true != 0;
  if (rule.dxf_id_engaged != 0) {
    out_rule.dxf_id = rule.dxf_id;
  }
  if (rule.id != nullptr && rule.id[0] != '\0') {
    out_rule.id = rule.id;
  } else {
    // Synthesize a stable id from priority. The format mirrors the
    // x14:cfRule guid-like string Excel emits, but uses a priority
    // suffix so add-then-list is deterministic.
    out_rule.id = "{cf-" + std::to_string(out_rule.priority) + "}";
  }
  if (rule.formula1 != nullptr) {
    out_rule.formula1 = std::string(rule.formula1);
  }
  if (rule.formula2 != nullptr) {
    out_rule.formula2 = std::string(rule.formula2);
  }
  if (rule.op_engaged != 0) {
    out_rule.op = static_cast<formulon::cf::CellIsOperator>(rule.op);
  }
  if (rule.rank_engaged != 0) {
    out_rule.rank = rule.rank;
  }
  out_rule.percent = (rule.percent != 0);
  out_rule.bottom = (rule.bottom != 0);
  out_rule.above_average = (rule.above_average != 0);
  out_rule.equal_average = (rule.equal_average != 0);
  if (rule.std_dev_engaged != 0) {
    out_rule.std_dev = rule.std_dev;
  }
  if (rule.text != nullptr) {
    out_rule.text = std::string(rule.text);
  }
  if (rule.time_period_engaged != 0) {
    out_rule.time_period = static_cast<formulon::cf::TimePeriod>(rule.time_period);
  }

  new_block.rules.push_back(std::move(out_rule));
  wb->workbook().sheet(sheet_index).mutable_conditional_formats().push_back(std::move(new_block));
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_remove_at(fm_workbook_t* wb, std::size_t sheet_index, std::size_t idx) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_remove_at: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_remove_at: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  auto& blocks = wb->workbook().sheet(sheet_index).mutable_conditional_formats();
  std::size_t b = 0;
  std::size_t r = 0;
  if (!resolve_flat_index(blocks, idx, &b, &r)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_cf_remove_at: idx out of range",
                             "idx=" + std::to_string(idx));
  }
  blocks[b].rules.erase(blocks[b].rules.begin() + static_cast<std::ptrdiff_t>(r));
  if (blocks[b].rules.empty()) {
    blocks.erase(blocks.begin() + static_cast<std::ptrdiff_t>(b));
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_clear(fm_workbook_t* wb, std::size_t sheet_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_clear: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_clear: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  wb->workbook().sheet(sheet_index).mutable_conditional_formats().clear();
  return 0;
}

// ---------------------------------------------------------------------------
// Trace precedents / dependents
// ---------------------------------------------------------------------------
//
// Bridges `RecalcEngine::dep_graph()` over the stable C ABI. The opaque
// `fm_cell_nodes_t` owns the BFS-expanded result so the index accessors
// can return data without re-walking the graph. Depth is capped at 32
// to keep cyclic graphs from blowing up the queue.

struct fm_cell_nodes {
  std::vector<formulon::eval::CellNodeId> nodes;
};

namespace {

constexpr std::uint32_t kMaxTraceDepth = 32U;

// Effective depth: 0 / 1 → 1-step (direct neighbors only); larger
// values are capped at `kMaxTraceDepth` to keep BFS bounded.
std::uint32_t effective_depth(std::uint32_t depth) {
  if (depth <= 1) {
    return 1U;
  }
  return depth > kMaxTraceDepth ? kMaxTraceDepth : depth;
}

// BFS from `seed` using `next_neighbors(node) -> vector<CellNodeId>`
// up to `depth` hops. The seed itself is excluded from the result. The
// result preserves first-encountered order, which matches the
// dependency-graph adjacency list order at depth 1 and stays
// deterministic for deeper traversals.
template <typename NextFn>
std::vector<formulon::eval::CellNodeId> bfs_collect(formulon::eval::CellNodeId seed, std::uint32_t depth,
                                                    NextFn&& next_neighbors) {
  std::vector<formulon::eval::CellNodeId> ordered;
  std::unordered_set<formulon::eval::CellNodeId, formulon::eval::CellNodeIdHash> seen;
  std::vector<formulon::eval::CellNodeId> frontier;
  std::vector<formulon::eval::CellNodeId> next_frontier;
  seen.insert(seed);
  frontier.push_back(seed);
  for (std::uint32_t hop = 0; hop < depth; ++hop) {
    next_frontier.clear();
    for (formulon::eval::CellNodeId node : frontier) {
      for (formulon::eval::CellNodeId nbr : next_neighbors(node)) {
        if (seen.insert(nbr).second) {
          ordered.push_back(nbr);
          next_frontier.push_back(nbr);
        }
      }
    }
    if (next_frontier.empty()) {
      break;
    }
    frontier.swap(next_frontier);
  }
  return ordered;
}

template <bool kPrecedents>
fm_status_t trace_impl(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row, std::uint32_t col,
                       std::uint32_t depth, fm_cell_nodes_t** out) {
  clear_last_error();
  constexpr const char* fn_name = kPrecedents ? "fm_workbook_precedents" : "fm_workbook_dependents";
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(
        formulon::FormulonErrorCode::kBindingNullPointer,
        kPrecedents ? "fm_workbook_precedents: NULL argument" : "fm_workbook_dependents: NULL argument");
  }
  if (sheet >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument,
        kPrecedents ? "fm_workbook_precedents: sheet out of range" : "fm_workbook_dependents: sheet out of range",
        "sheet=" + std::to_string(sheet));
  }
  (void)fn_name;
  const auto& graph = wb->workbook().recalc_engine().dep_graph();
  formulon::eval::CellNodeId seed{static_cast<std::uint16_t>(sheet), row, col};
  auto neighbors = [&graph](formulon::eval::CellNodeId node) {
    if constexpr (kPrecedents) {
      return graph.dependencies_of(node);
    } else {
      return graph.dependents_of(node);
    }
  };
  auto nodes = bfs_collect(seed, effective_depth(depth), neighbors);

  auto handle = std::unique_ptr<fm_cell_nodes_t>(new fm_cell_nodes_t{});
  handle->nodes = std::move(nodes);
  *out = handle.release();
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_precedents(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                              std::uint32_t col, std::uint32_t depth, fm_cell_nodes_t** out) {
  return trace_impl<true>(wb, sheet, row, col, depth, out);
}

extern "C" fm_status_t fm_workbook_dependents(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                              std::uint32_t col, std::uint32_t depth, fm_cell_nodes_t** out) {
  return trace_impl<false>(wb, sheet, row, col, depth, out);
}

extern "C" void fm_cell_nodes_destroy(fm_cell_nodes_t* nodes) {
  delete nodes;
}

extern "C" size_t fm_cell_nodes_count(const fm_cell_nodes_t* nodes) {
  if (nodes == nullptr) {
    return 0;
  }
  return nodes->nodes.size();
}

extern "C" fm_status_t fm_cell_nodes_at(const fm_cell_nodes_t* nodes, std::size_t idx, fm_cell_node_t* out) {
  clear_last_error();
  if (nodes == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cell_nodes_at: NULL argument");
  }
  if (idx >= nodes->nodes.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_cell_nodes_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(nodes->nodes.size()));
  }
  out->sheet = static_cast<std::uint32_t>(nodes->nodes[idx].sheet_id);
  out->row = nodes->nodes[idx].row;
  out->col = nodes->nodes[idx].col;
  return 0;
}

// ---------------------------------------------------------------------------
// Dynamic-array spill payload
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Function catalog metadata
// ---------------------------------------------------------------------------
//
// Surfaces the registry's canonical-name + arity data through the C ABI
// so JS / Python autocomplete UIs can enumerate Formulon functions
// without reaching into the engine internals. `description` /
// `signature_template` are reserved for the locale metadata table
// (data/function_metadata_<locale>.cpp) and stay `NULL` until that
// table is wired up; the surface itself is shippable today.
//
// The sorted name list is computed on first use and cached for the
// process lifetime — order matters for UIs that expect deterministic
// enumeration, and the registry's `for_each_name` does not promise
// any order.

namespace {

const std::vector<std::string>& sorted_function_names() {
  static const std::vector<std::string> names = []() {
    std::vector<std::string> out;
    formulon::eval::default_registry().for_each_name(
        [](std::string_view name, void* ctx) { static_cast<std::vector<std::string>*>(ctx)->emplace_back(name); },
        &out);
    std::sort(out.begin(), out.end());
    return out;
  }();
  return names;
}

}  // namespace

extern "C" fm_status_t fm_function_metadata(const char* name, fm_locale_t /*locale*/, fm_function_metadata_t* out) {
  clear_last_error();
  if (name == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_function_metadata: NULL argument");
  }
  const auto& reg = formulon::eval::default_registry();
  const auto* def = reg.lookup(std::string_view(name));
  if (def == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_metadata: unknown function",
                             std::string("name=") + name);
  }
  *out = fm_function_metadata_t{};
  // `canonical_name` is held in static storage by the registry; surface
  // a borrowed pointer so callers can hold the result indefinitely.
  out->canonical_name = def->canonical_name.data();
  out->min_arity = def->min_arity;
  out->max_arity = def->max_arity;
  // Locale metadata table (description / signature_template) is not
  // yet populated — the public ABI documents these as nullable.
  out->signature_template = nullptr;
  out->description = nullptr;
  return 0;
}

extern "C" std::size_t fm_function_count(void) {
  return sorted_function_names().size();
}

extern "C" fm_status_t fm_function_name_at(std::size_t idx, const char** out_name) {
  clear_last_error();
  if (out_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_function_name_at: out_name is NULL");
  }
  const auto& names = sorted_function_names();
  if (idx >= names.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_name_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(names.size()));
  }
  *out_name = names[idx].c_str();
  return 0;
}

extern "C" fm_status_t fm_function_localize(const char* canonical_name, fm_locale_t /*locale*/,
                                            const char** out_localized) {
  clear_last_error();
  if (canonical_name == nullptr || out_localized == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_function_localize: NULL argument");
  }
  const auto* def = formulon::eval::default_registry().lookup(std::string_view(canonical_name));
  if (def == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_localize: unknown function",
                             std::string("canonical_name=") + canonical_name);
  }
  // Locale alias table not yet populated — fall through to the canonical
  // name. Once `data/function_names_<locale>.csv` lands, this lookup
  // will branch on `locale` and consult the alias table first.
  *out_localized = def->canonical_name.data();
  return 0;
}

extern "C" fm_status_t fm_function_canonicalize(const char* localized_name, fm_locale_t /*locale*/,
                                                const char** out_canonical) {
  clear_last_error();
  if (localized_name == nullptr || out_canonical == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_function_canonicalize: NULL argument");
  }
  // Alias table not yet populated — fall through to a case-insensitive
  // canonical-name match.
  const auto* def = formulon::eval::default_registry().lookup(std::string_view(localized_name));
  if (def == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_function_canonicalize: unknown function",
                             std::string("localized_name=") + localized_name);
  }
  *out_canonical = def->canonical_name.data();
  return 0;
}

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
// Styles
// ---------------------------------------------------------------------------

namespace {

/// Validates the `(handle, sheet_index, row, col)` quad and resolves
/// the cell's `xf_index`. On failure populates the thread-local
/// diagnostic and returns the status. On success writes
/// `*out_xf_index` and returns `kOk`. The `xf_index` defaults to `0`
/// (the default xf) when no cell exists at the address.
fm_status_t resolve_xf(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row, std::uint32_t col,
                       std::uint32_t* out_xf_index, const char* fn) {
  if (wb == nullptr || out_xf_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  if (sheet >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, fn,
        "sheet_index=" + std::to_string(sheet) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  const formulon::Cell* cell = wb->workbook().sheet(sheet).cell_at(row, col);
  *out_xf_index = (cell != nullptr) ? cell->xf_index : 0U;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_cell_get_xf_index(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                            uint32_t* out_xf_index) {
  clear_last_error();
  return resolve_xf(wb, sheet, row, col, out_xf_index, "fm_cell_get_xf_index");
}

extern "C" fm_status_t fm_cell_set_xf_index(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                            uint32_t xf_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cell_set_xf_index: wb is NULL");
  }
  auto r = wb->workbook().set_cell_xf_index(sheet, row, col, xf_index);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_xf(fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_cell_xf: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (xf_index >= styles.cell_xfs.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_xf: xf_index out of range",
        "xf_index=" + std::to_string(xf_index) + " cell_xfs_count=" + std::to_string(styles.cell_xfs.size()));
  }
  const formulon::io::CellXf& xf = styles.cell_xfs[xf_index];
  out->font_index = xf.font_index;
  out->fill_index = xf.fill_index;
  out->border_index = xf.border_index;
  out->num_fmt_id = xf.num_fmt_id;
  out->horizontal_align = xf.horizontal_align;
  out->vertical_align = xf.vertical_align;
  out->wrap_text = xf.wrap_text ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_styles_get_font(fm_workbook_t* wb, uint32_t font_index, fm_font_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_font: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (font_index >= styles.fonts.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_font: font_index out of range",
        "font_index=" + std::to_string(font_index) + " fonts_count=" + std::to_string(styles.fonts.size()));
  }
  const formulon::io::FontRecord& f = styles.fonts[font_index];
  out->name = f.name.c_str();
  out->size = f.size;
  out->color_argb = f.color_argb;
  out->bold = f.bold ? 1 : 0;
  out->italic = f.italic ? 1 : 0;
  out->strike = f.strike ? 1 : 0;
  out->underline = f.underline;
  return 0;
}

extern "C" fm_status_t fm_styles_get_num_fmt_string(fm_workbook_t* wb, uint16_t num_fmt_id, const char** out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_num_fmt_string: NULL argument");
  }
  // Built-in ids (0..163) resolve through the writer's `.rodata` table.
  if (num_fmt_id < 164U) {
    const char* s = formulon::io::builtin_num_fmt(num_fmt_id);
    if (s != nullptr && s[0] != '\0') {
      *out = s;
      return 0;
    }
    // Reserved-but-undocumented built-in slot. Fall through to the
    // custom search to honour any caller-defined override; if neither
    // is present, surface kInvalidArgument so callers know the id is
    // not resolvable.
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
    if (n.id == num_fmt_id && n.format_string_index < styles.num_fmt_strings.size()) {
      *out = styles.num_fmt_strings[n.format_string_index].c_str();
      return 0;
    }
  }
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_num_fmt_string: id not found",
                           "num_fmt_id=" + std::to_string(num_fmt_id));
}

extern "C" fm_status_t fm_styles_get_fill(fm_workbook_t* wb, uint32_t fill_index, fm_fill_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_fill: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (fill_index >= styles.fills.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_fill: fill_index out of range",
        "fill_index=" + std::to_string(fill_index) + " fills_count=" + std::to_string(styles.fills.size()));
  }
  const formulon::io::FillRecord& f = styles.fills[fill_index];
  out->pattern = f.pattern;
  out->fg_argb = f.fg_argb;
  out->bg_argb = f.bg_argb;
  return 0;
}

extern "C" fm_status_t fm_styles_get_border(fm_workbook_t* wb, uint32_t border_index, fm_border_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_border: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (border_index >= styles.borders.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_border: border_index out of range",
        "border_index=" + std::to_string(border_index) + " borders_count=" + std::to_string(styles.borders.size()));
  }
  const formulon::io::BorderRecord& b = styles.borders[border_index];
  auto fill_side = [](const formulon::io::BorderSide& src, fm_border_side& dst) noexcept {
    dst.style = src.style;
    dst.color_argb = src.color_argb;
  };
  fill_side(b.left, out->left);
  fill_side(b.right, out->right);
  fill_side(b.top, out->top);
  fill_side(b.bottom, out->bottom);
  fill_side(b.diagonal, out->diagonal);
  out->diagonal_up = b.diagonal_up ? 1 : 0;
  out->diagonal_down = b.diagonal_down ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_styles_get_font_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_font_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().fonts.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_fill_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_fill_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().fills.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_border_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_border_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().borders.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_xf_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_xf_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().cell_xfs.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_style_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().cell_styles.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_style(fm_workbook_t* wb, uint32_t index, fm_cell_style_record_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (index >= styles.cell_styles.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_style: index out of range",
        "index=" + std::to_string(index) + " cell_styles_count=" + std::to_string(styles.cell_styles.size()));
  }
  const formulon::io::CellStyleRecord& cs = styles.cell_styles[index];
  out->name = cs.name.c_str();
  out->xf_id = cs.xf_id;
  out->builtin_id = cs.builtin_id;
  out->i_level = cs.i_level;
  out->hidden = cs.hidden ? 1 : 0;
  out->custom_builtin = cs.custom_builtin ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_style_xf_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_xf_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().cell_style_xfs.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_style_xf(fm_workbook_t* wb, uint32_t index, fm_cell_xf* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_xf: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (index >= styles.cell_style_xfs.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_style_xf: index out of range",
        "index=" + std::to_string(index) + " cell_style_xfs_count=" + std::to_string(styles.cell_style_xfs.size()));
  }
  const formulon::io::CellXf& xf = styles.cell_style_xfs[index];
  out->font_index = xf.font_index;
  out->fill_index = xf.fill_index;
  out->border_index = xf.border_index;
  out->num_fmt_id = xf.num_fmt_id;
  out->horizontal_align = xf.horizontal_align;
  out->vertical_align = xf.vertical_align;
  out->wrap_text = xf.wrap_text ? 1 : 0;
  return 0;
}

namespace {

/// Field-for-field equality for the C++ side `FontRecord`. Excluded
/// from `<algorithm>` because the struct has no `operator==`; defining
/// one here keeps the dedup-comparator confined to the C ABI's needs.
bool font_records_equal(const formulon::io::FontRecord& a, const formulon::io::FontRecord& b) {
  return a.name == b.name && a.size == b.size && a.bold == b.bold && a.italic == b.italic && a.strike == b.strike &&
         a.underline == b.underline && a.color_argb == b.color_argb;
}

bool fill_records_equal(const formulon::io::FillRecord& a, const formulon::io::FillRecord& b) noexcept {
  return a.pattern == b.pattern && a.fg_argb == b.fg_argb && a.bg_argb == b.bg_argb;
}

bool border_sides_equal(const formulon::io::BorderSide& a, const formulon::io::BorderSide& b) noexcept {
  return a.style == b.style && a.color_argb == b.color_argb;
}

bool border_records_equal(const formulon::io::BorderRecord& a, const formulon::io::BorderRecord& b) noexcept {
  return border_sides_equal(a.left, b.left) && border_sides_equal(a.right, b.right) &&
         border_sides_equal(a.top, b.top) && border_sides_equal(a.bottom, b.bottom) &&
         border_sides_equal(a.diagonal, b.diagonal) && a.diagonal_up == b.diagonal_up &&
         a.diagonal_down == b.diagonal_down;
}

bool cell_xfs_equal(const formulon::io::CellXf& a, const formulon::io::CellXf& b) noexcept {
  return a.font_index == b.font_index && a.fill_index == b.fill_index && a.border_index == b.border_index &&
         a.num_fmt_id == b.num_fmt_id && a.horizontal_align == b.horizontal_align &&
         a.vertical_align == b.vertical_align && a.wrap_text == b.wrap_text;
}

}  // namespace

extern "C" fm_status_t fm_styles_add_font(fm_workbook_t* wb, fm_font_record record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_font: NULL argument");
  }
  formulon::io::FontRecord candidate;
  candidate.name = (record.name != nullptr) ? std::string(record.name) : std::string();
  candidate.size = record.size;
  candidate.bold = record.bold != 0;
  candidate.italic = record.italic != 0;
  candidate.strike = record.strike != 0;
  candidate.underline = record.underline;
  candidate.color_argb = record.color_argb;

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.fonts.size(); ++i) {
    if (font_records_equal(styles.fonts[i], candidate)) {
      *out_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.fonts.push_back(std::move(candidate));
  *out_index = static_cast<uint32_t>(styles.fonts.size() - 1);
  return 0;
}

extern "C" fm_status_t fm_styles_add_fill(fm_workbook_t* wb, fm_fill_record record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_fill: NULL argument");
  }
  formulon::io::FillRecord candidate;
  candidate.pattern = record.pattern;
  candidate.fg_argb = record.fg_argb;
  candidate.bg_argb = record.bg_argb;

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.fills.size(); ++i) {
    if (fill_records_equal(styles.fills[i], candidate)) {
      *out_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.fills.push_back(candidate);
  *out_index = static_cast<uint32_t>(styles.fills.size() - 1);
  return 0;
}

extern "C" fm_status_t fm_styles_add_border(fm_workbook_t* wb, fm_border_record record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_border: NULL argument");
  }
  auto pull_side = [](const fm_border_side& src) noexcept {
    formulon::io::BorderSide dst;
    dst.style = src.style;
    dst.color_argb = src.color_argb;
    return dst;
  };
  formulon::io::BorderRecord candidate;
  candidate.left = pull_side(record.left);
  candidate.right = pull_side(record.right);
  candidate.top = pull_side(record.top);
  candidate.bottom = pull_side(record.bottom);
  candidate.diagonal = pull_side(record.diagonal);
  candidate.diagonal_up = record.diagonal_up != 0;
  candidate.diagonal_down = record.diagonal_down != 0;

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.borders.size(); ++i) {
    if (border_records_equal(styles.borders[i], candidate)) {
      *out_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.borders.push_back(candidate);
  *out_index = static_cast<uint32_t>(styles.borders.size() - 1);
  return 0;
}

extern "C" fm_status_t fm_styles_add_num_fmt(fm_workbook_t* wb, const char* format_code, uint16_t* out_num_fmt_id) {
  clear_last_error();
  if (wb == nullptr || out_num_fmt_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_num_fmt: NULL argument");
  }
  const std::string code = (format_code != nullptr) ? std::string(format_code) : std::string();

  // Step 1: built-in match.
  for (uint16_t id = 0; id < 164U; ++id) {
    const char* s = formulon::io::builtin_num_fmt(id);
    if (s != nullptr && s[0] != '\0' && code == s) {
      *out_num_fmt_id = id;
      return 0;
    }
  }

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();

  // Step 2: existing custom entry.
  for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
    if (n.format_string_index < styles.num_fmt_strings.size() &&
        styles.num_fmt_strings[n.format_string_index] == code) {
      *out_num_fmt_id = n.id;
      return 0;
    }
  }

  // Step 3: append a new custom entry. New id is one past the largest
  // existing custom id, with `163` as the lower bound (the last
  // built-in slot).
  uint16_t next_id = 163U;
  for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
    if (n.id > next_id) {
      next_id = n.id;
    }
  }
  ++next_id;

  styles.num_fmt_strings.push_back(code);
  formulon::io::NumFmtRecord rec;
  rec.id = next_id;
  rec.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size() - 1);
  styles.num_fmts.push_back(rec);
  *out_num_fmt_id = next_id;
  return 0;
}

extern "C" fm_status_t fm_styles_add_cell_xf(fm_workbook_t* wb, fm_cell_xf record, uint32_t* out_xf_index) {
  clear_last_error();
  if (wb == nullptr || out_xf_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_cell_xf: NULL argument");
  }
  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();

  // Validate referenced indices. Reject out-of-range references rather
  // than auto-growing the parallel tables; callers must register fonts /
  // fills / borders before referencing them.
  if (record.font_index >= styles.fonts.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_add_cell_xf: font_index out of range",
        "font_index=" + std::to_string(record.font_index) + " fonts_count=" + std::to_string(styles.fonts.size()));
  }
  if (record.fill_index >= styles.fills.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_add_cell_xf: fill_index out of range",
        "fill_index=" + std::to_string(record.fill_index) + " fills_count=" + std::to_string(styles.fills.size()));
  }
  if (record.border_index >= styles.borders.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_styles_add_cell_xf: border_index out of range",
                             "border_index=" + std::to_string(record.border_index) +
                                 " borders_count=" + std::to_string(styles.borders.size()));
  }
  // num_fmt_id must be a documented built-in or a registered custom id.
  bool num_fmt_ok = false;
  if (record.num_fmt_id < 164U) {
    const char* s = formulon::io::builtin_num_fmt(record.num_fmt_id);
    num_fmt_ok = (s != nullptr && s[0] != '\0');
  }
  if (!num_fmt_ok) {
    for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
      if (n.id == record.num_fmt_id) {
        num_fmt_ok = true;
        break;
      }
    }
  }
  if (!num_fmt_ok) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_styles_add_cell_xf: num_fmt_id not registered",
                             "num_fmt_id=" + std::to_string(record.num_fmt_id));
  }

  // Confirm the read-side `fm_cell_xf` struct mirrors the engine
  // `formulon::io::CellXf` field-for-field; the write side reuses the
  // same shape so a layout drift would silently corrupt records.
  static_assert(sizeof(record.font_index) == sizeof(formulon::io::CellXf::font_index),
                "fm_cell_xf::font_index width must match formulon::io::CellXf");
  static_assert(sizeof(record.fill_index) == sizeof(formulon::io::CellXf::fill_index),
                "fm_cell_xf::fill_index width must match formulon::io::CellXf");
  static_assert(sizeof(record.border_index) == sizeof(formulon::io::CellXf::border_index),
                "fm_cell_xf::border_index width must match formulon::io::CellXf");
  static_assert(sizeof(record.num_fmt_id) == sizeof(formulon::io::CellXf::num_fmt_id),
                "fm_cell_xf::num_fmt_id width must match formulon::io::CellXf");

  formulon::io::CellXf candidate;
  candidate.font_index = record.font_index;
  candidate.fill_index = record.fill_index;
  candidate.border_index = record.border_index;
  candidate.num_fmt_id = record.num_fmt_id;
  candidate.horizontal_align = record.horizontal_align;
  candidate.vertical_align = record.vertical_align;
  candidate.wrap_text = record.wrap_text != 0;

  for (std::size_t i = 0; i < styles.cell_xfs.size(); ++i) {
    if (cell_xfs_equal(styles.cell_xfs[i], candidate)) {
      *out_xf_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.cell_xfs.push_back(candidate);
  *out_xf_index = static_cast<uint32_t>(styles.cell_xfs.size() - 1);
  return 0;
}

// ---------------------------------------------------------------------------
// External links
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_external_link_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_external_link_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().external_links().size());
  return 0;
}

extern "C" fm_status_t fm_workbook_external_link_at(fm_workbook_t* wb, uint32_t index, fm_external_link_record_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_external_link_at: NULL argument");
  }
  const auto& links = wb->workbook().external_links();
  if (index >= links.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_external_link_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(links.size()));
  }
  const formulon::io::ExternalLinkRecord& rec = links[index];
  out->index = rec.index;
  out->rel_id = rec.rel_id.c_str();
  out->part_path = rec.part_path.c_str();
  out->target = rec.target.c_str();
  out->target_external = rec.target_external ? 1 : 0;
  switch (rec.kind) {
    case formulon::io::ExternalLinkRecord::Kind::kExternalBook:
      out->kind = FM_EXTERNAL_LINK_KIND_EXTERNAL_BOOK;
      break;
    case formulon::io::ExternalLinkRecord::Kind::kOleLink:
      out->kind = FM_EXTERNAL_LINK_KIND_OLE;
      break;
    case formulon::io::ExternalLinkRecord::Kind::kDdeLink:
      out->kind = FM_EXTERNAL_LINK_KIND_DDE;
      break;
    case formulon::io::ExternalLinkRecord::Kind::kUnknown:
    default:
      out->kind = FM_EXTERNAL_LINK_KIND_UNKNOWN;
      break;
  }
  return 0;
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
