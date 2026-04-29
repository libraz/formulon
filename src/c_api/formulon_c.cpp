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

#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
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
