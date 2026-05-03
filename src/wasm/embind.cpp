// Copyright 2026 libraz. Licensed under the MIT License.
//
// Emscripten / embind bindings for the Formulon WASM build.
//
// This translation unit is compiled only when `FM_BUILD_WASM=ON` (i.e.
// the build is being driven by `emcmake`). It is a thin C++ wrapper
// around the stable C ABI declared in `c_api/formulon_c.h`; the
// JavaScript / TypeScript surface NEVER touches `formulon::Workbook`,
// `formulon::Value`, or any other internal symbol directly. This keeps
// the WASM binding decoupled from engine internals exactly the way the
// CLI, Python ctypes wheel, and any future binding are.
//
// ## Design notes
//
//   * The whole engine is built with `-fno-exceptions -fno-rtti` and we
//     do NOT lift that constraint here. embind itself can be compiled
//     without exception catching as long as user code never throws —
//     the binding surface uses a result-object pattern (`{ ok, status,
//     message, ... }`) instead of throwing across the JS boundary.
//
//   * Every entry point clears and re-fetches the thread-local last
//     error so JS callers can read `lastErrorMessage()` / `lastErrorContext()`
//     after a failed call. That mirrors the contract of the C ABI.
//
//   * `JsWorkbook` owns an `fm_workbook_t*` via RAII (destructor calls
//     `fm_workbook_destroy`). It is move-only on the C++ side; embind
//     surfaces it as a JS class with explicit `delete()` semantics
//     (Emscripten convention for native handles).
//
//   * `loadBytes` accepts `emscripten::val` because that's the simplest
//     way to receive a `Uint8Array` without committing to a typed-array
//     binding helper. We copy into a `std::vector<uint8_t>` and forward
//     to `fm_workbook_load`. `save` mirrors the inverse direction:
//     allocate the C-side buffer, copy into a fresh `Uint8Array` on the
//     JS heap, free the native buffer.
//
//   * `evalFormula(formula)` is a convenience that mirrors
//     `formulon_cli eval`: create empty workbook, place the formula at
//     `Sheet1!A1`, recalc, return the cached value.

#include <emscripten/bind.h>
#include <emscripten/val.h>

#include <cstdint>
#include <cstring>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"

namespace {

/// Mirror of `fm_value_t` for the JS boundary.
///
/// embind's `value_object<>` cannot register a C union directly, so we
/// flatten the variant payload into individual scalar fields and let JS
/// code dispatch on `kind`. Only the field selected by `kind` is
/// meaningful; the others carry default-zero values.
struct JsValue {
  int32_t kind = 0;       ///< `fm_value_kind_t` ordinal (0..7).
  double number = 0.0;    ///< Active when kind == FM_VAL_NUMBER.
  int32_t boolean = 0;    ///< Active when kind == FM_VAL_BOOL (0/1).
  std::string text;       ///< Active when kind == FM_VAL_TEXT.
  int32_t errorCode = 0;  ///< Active when kind == FM_VAL_ERROR.
};

/// Generic { ok, status, message, context } envelope returned by every
/// fallible binding entry point. Modelled as a value-object so JS sees
/// `{ ok: true, status: 0, message: '', context: '' }` on success, or
/// `{ ok: false, status: <int>, message: <str>, context: <str> }` on
/// failure. This is the substitute for throwing across the JS boundary,
/// which we cannot do under `-fno-exceptions`.
struct JsStatus {
  bool ok = true;
  int32_t status = 0;
  std::string message;
  std::string context;
};

/// Result envelope for `evalFormula` — bundles a `JsStatus` and a
/// `JsValue` payload so JS only inspects one return shape.
struct JsEvalResult {
  JsStatus status;
  JsValue value;
};

/// Result envelope for `getValue` calls on a `Workbook`.
struct JsCellResult {
  JsStatus status;
  JsValue value;
};

/// Result envelope for `save()` returning the bytes as a JS Uint8Array.
struct JsSaveResult {
  JsStatus status;
  emscripten::val bytes = emscripten::val::null();
};

/// Result envelope for `sheetName` (string-payload variant of JsStatus).
struct JsStringResult {
  JsStatus status;
  std::string value;
};

/// JS-side mirror of `fm_cf_color_t`. Channels are 0-255 (sRGB); embind
/// tends to surface signed `int32_t` more cleanly than `uint8_t`, so
/// the wire types here are intentionally widened.
struct JsCfColor {
  int32_t r = 0;
  int32_t g = 0;
  int32_t b = 0;
  int32_t a = 0;
};

/// JS-side mirror of `fm_cf_match_t`. The active fields depend on
/// `kind`; the others carry default-zero values.
struct JsCfMatch {
  int32_t kind = 0;
  int32_t priority = 0;
  int32_t dxfIdEngaged = 0;
  uint32_t dxfId = 0;
  JsCfColor color{};
  double barLengthPct = 0.0;
  double barAxisPositionPct = 0.0;
  int32_t barIsNegative = 0;
  JsCfColor barFill{};
  int32_t barBorderEngaged = 0;
  JsCfColor barBorder{};
  int32_t barGradient = 0;
  int32_t iconSetName = 0;
  int32_t iconIndex = 0;
};

/// One cell's CF result: row / col plus the priority-ascending match
/// list. Mirrors `cf::CFRangeCellMatches` flattened across the C ABI.
struct JsCfCellResult {
  uint32_t row = 0;
  uint32_t col = 0;
  std::vector<JsCfMatch> matches;
};

/// Return envelope for `Workbook.evaluateCfRange(...)`. `cells` is
/// sparse: only cells that produced at least one match appear.
struct JsCfRangeResult {
  JsStatus status;
  std::vector<JsCfCellResult> cells;
};

/// JS-side mirror of `fm_sheet_view_t`.
struct JsSheetView {
  uint32_t zoomScale = 100U;
  uint32_t freezeRows = 0U;
  uint32_t freezeCols = 0U;
  int32_t tabHidden = 0;
};

/// Return envelope for `Workbook.getSheetView(...)`.
struct JsSheetViewResult {
  JsStatus status;
  JsSheetView view{};
};

/// JS-side mirror of `fm_column_layout_t`. Channel widths are normal
/// JS numbers; the embind-side `int32_t` stand-in for booleans matches
/// the rest of this binding surface.
struct JsColumnLayout {
  uint32_t first = 0U;
  uint32_t last = 0U;
  double width = 0.0;
  int32_t hidden = 0;
  int32_t outlineLevel = 0;
};

/// JS-side mirror of `fm_row_layout_t`.
struct JsRowLayout {
  uint32_t row = 0U;
  double height = 0.0;
  int32_t hidden = 0;
  int32_t outlineLevel = 0;
};

/// Return envelope for `Workbook.getSheetColumns(...)`.
struct JsColumnsResult {
  JsStatus status;
  std::vector<JsColumnLayout> columns;
};

/// Return envelope for `Workbook.getSheetRowOverrides(...)`.
struct JsRowsResult {
  JsStatus status;
  std::vector<JsRowLayout> rows;
};

/// Builds an `ok` envelope with empty diagnostic strings.
JsStatus ok_status() {
  return JsStatus{true, 0, std::string(), std::string()};
}

/// Builds a failure envelope, copying out the thread-local diagnostics
/// surfaced by the most recent C-ABI call.
JsStatus error_status(int32_t code) {
  JsStatus s;
  s.ok = false;
  s.status = code;
  const char* msg = fm_last_error_message();
  const char* ctx = fm_last_error_context();
  s.message = msg != nullptr ? msg : "";
  s.context = ctx != nullptr ? ctx : "";
  return s;
}

/// Translates a `fm_value_t` into the embind-friendly `JsValue`.
///
/// The text variant is copied (embind owns the string on the JS side).
/// The other variants project the scalar payload directly.
JsValue translate_value(const fm_value_t& v) {
  JsValue out;
  out.kind = static_cast<int32_t>(v.kind);
  switch (v.kind) {
    case FM_VAL_NUMBER:
      out.number = v.u.number;
      break;
    case FM_VAL_BOOL:
      out.boolean = v.u.boolean;
      break;
    case FM_VAL_TEXT:
      out.text = (v.u.text != nullptr) ? std::string(v.u.text) : std::string();
      break;
    case FM_VAL_ERROR:
      out.errorCode = v.u.error_code;
      break;
    case FM_VAL_BLANK:
    case FM_VAL_ARRAY:
    case FM_VAL_REF:
    case FM_VAL_LAMBDA:
    default:
      // Other variants carry no scalar payload across this boundary.
      break;
  }
  return out;
}

/// Copies the contents of a JS Uint8Array (passed via `emscripten::val`)
/// into a `std::vector<uint8_t>`. Returns an empty vector when the
/// argument is not a typed array.
std::vector<uint8_t> val_to_bytes(const emscripten::val& v) {
  if (v.isNull() || v.isUndefined()) {
    return {};
  }
  // `length` works for both Uint8Array and plain arrays; `byteLength`
  // is the canonical Uint8Array property. Prefer `length`.
  const auto len_val = v["length"];
  if (len_val.isUndefined()) {
    return {};
  }
  const std::size_t len = len_val.as<std::size_t>();
  std::vector<uint8_t> out(len);
  if (len == 0) {
    return out;
  }
  // emscripten::val's vecFromJSArray copies element-by-element via
  // `as<T>`. For a typed Uint8Array this is correct and avoids relying
  // on heap-pointer aliasing tricks that depend on internal layout.
  for (std::size_t i = 0; i < len; ++i) {
    out[i] = v[i].as<uint8_t>();
  }
  return out;
}

/// Materialises a fresh JS `Uint8Array` from a contiguous byte range.
///
/// We construct an empty Uint8Array of the requested length on the JS
/// side and write each byte through `set(i, val)`. This is O(N) but
/// avoids any reliance on internal heap layout, and the smoke-test
/// payloads (a few KB at most) keep this comfortably under a
/// millisecond.
emscripten::val bytes_to_val(const uint8_t* data, std::size_t len) {
  emscripten::val u8 = emscripten::val::global("Uint8Array").new_(len);
  for (std::size_t i = 0; i < len; ++i) {
    u8.set(i, data[i]);
  }
  return u8;
}

/// Translates an `fm_cf_color_t` into the embind value-object mirror.
JsCfColor translate_cf_color(const fm_cf_color_t& c) {
  JsCfColor out;
  out.r = static_cast<int32_t>(c.r);
  out.g = static_cast<int32_t>(c.g);
  out.b = static_cast<int32_t>(c.b);
  out.a = static_cast<int32_t>(c.a);
  return out;
}

/// Translates an `fm_cf_match_t` POD into the embind value-object mirror.
JsCfMatch translate_cf_match(const fm_cf_match_t& m) {
  JsCfMatch out;
  out.kind = static_cast<int32_t>(m.kind);
  out.priority = m.priority;
  out.dxfIdEngaged = m.dxf_id_engaged;
  out.dxfId = m.dxf_id;
  out.color = translate_cf_color(m.color);
  out.barLengthPct = m.bar_length_pct;
  out.barAxisPositionPct = m.bar_axis_position_pct;
  out.barIsNegative = m.bar_is_negative;
  out.barFill = translate_cf_color(m.bar_fill);
  out.barBorderEngaged = m.bar_border_engaged;
  out.barBorder = translate_cf_color(m.bar_border);
  out.barGradient = m.bar_gradient;
  out.iconSetName = m.icon_set_name;
  out.iconIndex = static_cast<int32_t>(m.icon_index);
  return out;
}

/// RAII wrapper around `fm_workbook_t*`. Move-only; embind surfaces it
/// as a JS class with `.delete()` for explicit lifetime management.
class JsWorkbook {
 public:
  JsWorkbook() = default;

  ~JsWorkbook() {
    if (handle_ != nullptr) {
      fm_workbook_destroy(handle_);
      handle_ = nullptr;
    }
  }

  // Move-only.
  JsWorkbook(const JsWorkbook&) = delete;
  JsWorkbook& operator=(const JsWorkbook&) = delete;

  JsWorkbook(JsWorkbook&& other) noexcept : handle_(other.handle_) { other.handle_ = nullptr; }

  JsWorkbook& operator=(JsWorkbook&& other) noexcept {
    if (this != &other) {
      if (handle_ != nullptr) {
        fm_workbook_destroy(handle_);
      }
      handle_ = other.handle_;
      other.handle_ = nullptr;
    }
    return *this;
  }

  /// Static factory: empty workbook with the default `Sheet1`.
  static JsWorkbook* createDefault() {
    auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
    fm_status_t rc = fm_workbook_create(&wb->handle_);
    if (rc != 0) {
      // The handle remains null; the caller can read lastErrorMessage().
      // We still return the object so JS can inspect status via a
      // subsequent isValid() check.
    }
    return wb.release();
  }

  /// Static factory: empty workbook with no sheets.
  static JsWorkbook* createEmpty() {
    auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
    (void)fm_workbook_create_empty(&wb->handle_);
    return wb.release();
  }

  /// Static factory: load from an in-memory `.xlsx` byte buffer.
  ///
  /// On success returns a newly-allocated JsWorkbook with a populated
  /// handle. On failure the returned wrapper has a null handle and the
  /// caller should consult `lastErrorMessage()` for diagnostics. JS
  /// callers should always test `isValid()` before using the workbook.
  static JsWorkbook* loadBytes(emscripten::val bytes) {
    auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
    const std::vector<uint8_t> buf = val_to_bytes(bytes);
    if (buf.empty()) {
      return wb.release();
    }
    (void)fm_workbook_load(buf.data(), buf.size(), &wb->handle_);
    return wb.release();
  }

  /// True when the wrapper holds a live handle.
  bool isValid() const { return handle_ != nullptr; }

  /// Returns `{ ok, status, message, context, bytes }`. `bytes` is a
  /// freshly-allocated Uint8Array on the JS heap; `null` on failure.
  JsSaveResult save() const {
    JsSaveResult r;
    if (handle_ == nullptr) {
      r.status = error_status(/*kBindingNullPointer=*/7000);
      return r;
    }
    uint8_t* out = nullptr;
    std::size_t len = 0;
    fm_status_t rc = fm_workbook_save(handle_, &out, &len);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.bytes = bytes_to_val(out, len);
    fm_buffer_free(out);
    r.status = ok_status();
    return r;
  }

  /// Appends a new sheet with the given UTF-8 display name.
  JsStatus addSheet(const std::string& name) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_add_sheet(handle_, name.c_str());
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Moves the sheet at `fromIdx` to `toIdx`. `toIdx` is interpreted
  /// in the post-removal sheet list (Excel UI semantics).
  JsStatus moveSheet(uint32_t fromIdx, uint32_t toIdx) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_move_sheet(handle_, fromIdx, toIdx);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Removes the sheet at `index`.
  JsStatus removeSheet(uint32_t index) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_remove_sheet(handle_, index);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Renames the sheet at `index` to `newName`.
  JsStatus renameSheet(uint32_t index, const std::string& newName) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_rename_sheet(handle_, index, newName.c_str());
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets / appends / removes a workbook-scoped defined name. An empty
  /// `formula` removes the entry.
  JsStatus setDefinedName(const std::string& name, const std::string& formula) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_set_defined_name(handle_, name.c_str(), formula.c_str());
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Inserts `count` rows at `row` on `sheet`. References across the
  /// workbook are rewritten to follow the shift.
  JsStatus insertRows(uint32_t sheet, uint32_t row, uint32_t count) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_insert_rows(handle_, sheet, row, count);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Deletes `count` rows starting at `row` on `sheet`. References that
  /// fall inside the deleted interval collapse to `#REF!`.
  JsStatus deleteRows(uint32_t sheet, uint32_t row, uint32_t count) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_delete_rows(handle_, sheet, row, count);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Inserts `count` columns at `col` on `sheet`.
  JsStatus insertCols(uint32_t sheet, uint32_t col, uint32_t count) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_insert_cols(handle_, sheet, col, count);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Deletes `count` columns starting at `col` on `sheet`.
  JsStatus deleteCols(uint32_t sheet, uint32_t col, uint32_t count) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_delete_cols(handle_, sheet, col, count);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Returns the number of sheets (0 when handle is invalid).
  uint32_t sheetCount() const {
    if (handle_ == nullptr) {
      return 0;
    }
    return static_cast<uint32_t>(fm_workbook_sheet_count(handle_));
  }

  /// Returns the display name of sheet `idx`.
  JsStringResult sheetName(uint32_t idx) const {
    JsStringResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    const char* name = nullptr;
    fm_status_t rc = fm_workbook_sheet_name(handle_, idx, &name);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.value = (name != nullptr) ? std::string(name) : std::string();
    r.status = ok_status();
    return r;
  }

  /// Sets a numeric literal at `(sheet, row, col)`.
  JsStatus setNumber(uint32_t sheet, uint32_t row, uint32_t col, double value) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_set_number(handle_, sheet, row, col, value);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets a boolean literal at `(sheet, row, col)`.
  JsStatus setBool(uint32_t sheet, uint32_t row, uint32_t col, bool value) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_set_bool(handle_, sheet, row, col, value ? 1 : 0);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets a UTF-8 text literal at `(sheet, row, col)`.
  JsStatus setText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& text) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_set_text(handle_, sheet, row, col, text.c_str());
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Stores a Blank literal (clearing the cell).
  JsStatus setBlank(uint32_t sheet, uint32_t row, uint32_t col) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_set_blank(handle_, sheet, row, col);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Stores a formula at `(sheet, row, col)`.
  JsStatus setFormula(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_set_formula(handle_, sheet, row, col, formula.c_str());
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Reads the cached cell value at `(sheet, row, col)`.
  JsCellResult getValue(uint32_t sheet, uint32_t row, uint32_t col) const {
    JsCellResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    fm_value_t v{};
    fm_status_t rc = fm_workbook_get_value(handle_, sheet, row, col, &v);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.value = translate_value(v);
    r.status = ok_status();
    return r;
  }

  /// Drives a full incremental recalc.
  JsStatus recalc() {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_recalc(handle_);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Configures iterative-calculation knobs.
  JsStatus setIterative(bool enabled, uint32_t max_iterations, double max_change) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc =
        fm_workbook_set_iterative(handle_, enabled ? 1 : 0, static_cast<int32_t>(max_iterations), max_change);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Drives a viewport-bounded incremental recalc.
  ///
  /// `viewport` is a JS object of the shape
  /// `{ sheet, firstRow, lastRow, firstCol, lastCol }`. The returned
  /// envelope has the standard `{ ok, status, message, context }` plus
  /// a `recomputed` field reporting the number of cells the engine
  /// actually evaluated during the call.
  emscripten::val partialRecalc(emscripten::val viewport) {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      o.set("recomputed", static_cast<uint32_t>(0));
      return o;
    }
    fm_viewport vp{};
    vp.sheet = viewport["sheet"].as<uint32_t>();
    vp.first_row = viewport["firstRow"].as<uint32_t>();
    vp.last_row = viewport["lastRow"].as<uint32_t>();
    vp.first_col = viewport["firstCol"].as<uint32_t>();
    vp.last_col = viewport["lastCol"].as<uint32_t>();
    uint32_t recomputed = 0;
    fm_status_t rc = fm_workbook_partial_recalc(handle_, &vp, &recomputed);
    if (rc != 0) {
      o.set("status", error_status(rc));
      o.set("recomputed", static_cast<uint32_t>(0));
      return o;
    }
    o.set("status", ok_status());
    o.set("recomputed", recomputed);
    return o;
  }

  /// Installs an iterative-solver progress callback.
  ///
  /// The JS callback receives `(iteration, maxResidual, maxIterations)`
  /// after each Gauss-Seidel sweep and must return a truthy value to
  /// continue or a falsy value to abort. Pass `null` (or any value
  /// whose `isNull() / isUndefined()` test holds) to clear the
  /// callback.
  ///
  /// The wrapper holds the JS callable in a static `emscripten::val`
  /// slot for the lifetime of this workbook handle. We do not pass
  /// `user_data` through the C ABI: the JS layer does not need it
  /// because the closure captures whatever state the JS caller wants.
  /// As a consequence, only ONE JS progress callback can be active at
  /// a time across all workbook handles in this WASM instance —
  /// installing a new one displaces any previous registration. This
  /// matches the typical UI workflow (a single "calculation in
  /// progress" dialog) without requiring per-handle thread-local
  /// storage.
  JsStatus setIterativeProgress(emscripten::val cb) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    if (cb.isNull() || cb.isUndefined()) {
      js_progress_callback() = emscripten::val::null();
      fm_status_t rc = fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
      return rc == 0 ? ok_status() : error_status(rc);
    }
    js_progress_callback() = cb;
    fm_status_t rc = fm_workbook_set_iterative_progress(handle_, &JsWorkbook::iterativeProgressTrampoline, nullptr);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  // ---- Iteration / metadata accessors --------------------------------------
  // These mirror the C-ABI iteration / metadata entry points. Each one
  // returns either a status-bearing envelope or, for plain count
  // accessors, an unsigned integer (0 when the handle is invalid).

  uint32_t cellCount(uint32_t sheet) const {
    if (handle_ == nullptr) {
      return 0;
    }
    std::size_t count = 0;
    if (fm_workbook_cell_count(handle_, sheet, &count) != 0) {
      return 0;
    }
    return static_cast<uint32_t>(count);
  }

  /// Returns `{ status, row, col, formula, value }` for the `idx`-th
  /// stored cell on `sheet`.
  emscripten::val cellAt(uint32_t sheet, uint32_t idx) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    uint32_t row = 0;
    uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t v{};
    fm_status_t rc = fm_workbook_cell_at(handle_, sheet, idx, &row, &col, &formula, &v);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("row", row);
    o.set("col", col);
    o.set("formula", formula != nullptr ? emscripten::val(std::string(formula)) : emscripten::val::null());
    o.set("value", translate_value(v));
    return o;
  }

  uint32_t definedNameCount() const {
    if (handle_ == nullptr) {
      return 0;
    }
    return static_cast<uint32_t>(fm_workbook_defined_name_count(handle_));
  }

  emscripten::val definedNameAt(uint32_t idx) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    const char* name = nullptr;
    const char* formula = nullptr;
    fm_status_t rc = fm_workbook_defined_name_at(handle_, idx, &name, &formula);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("name", name != nullptr ? std::string(name) : std::string());
    o.set("formula", formula != nullptr ? std::string(formula) : std::string());
    return o;
  }

  uint32_t tableCount() const {
    if (handle_ == nullptr) {
      return 0;
    }
    return static_cast<uint32_t>(fm_workbook_table_count(handle_));
  }

  emscripten::val tableAt(uint32_t idx) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    const char* name = nullptr;
    const char* display = nullptr;
    const char* ref = nullptr;
    std::size_t sheet_index = 0;
    fm_status_t rc = fm_workbook_table_at(handle_, idx, &name, &display, &ref, &sheet_index);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("name", name != nullptr ? std::string(name) : std::string());
    o.set("displayName", display != nullptr ? std::string(display) : std::string());
    o.set("ref", ref != nullptr ? std::string(ref) : std::string());
    o.set("sheetIndex", static_cast<uint32_t>(sheet_index));
    return o;
  }

  uint32_t passthroughCount() const {
    if (handle_ == nullptr) {
      return 0;
    }
    return static_cast<uint32_t>(fm_workbook_passthrough_count(handle_));
  }

  emscripten::val passthroughAt(uint32_t idx) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    const char* path = nullptr;
    fm_status_t rc = fm_workbook_passthrough_at(handle_, idx, &path);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("path", path != nullptr ? std::string(path) : std::string());
    return o;
  }

  /// Evaluates every CF block on `sheet` against the inclusive cell
  /// range `[(firstRow, firstCol), (lastRow, lastCol)]`. The result is
  /// sparse: only cells that produced at least one match appear.
  ///
  /// `todaySerial` pins the date reference for `TimePeriod` rules.
  /// Pass `NaN` to disable.
  JsCfRangeResult evaluateCfRange(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow,
                                  uint32_t lastCol, double todaySerial) const {
    JsCfRangeResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    fm_cf_results_t* results = nullptr;
    fm_status_t rc =
        fm_workbook_cf_evaluate_range(handle_, sheet, firstRow, firstCol, lastRow, lastCol, todaySerial, &results);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    const std::size_t cell_count = fm_cf_results_cell_count(results);
    r.cells.reserve(cell_count);
    for (std::size_t i = 0; i < cell_count; ++i) {
      JsCfCellResult cell;
      uint32_t row = 0;
      uint32_t col = 0;
      std::size_t match_count = 0;
      if (fm_cf_results_cell_at(results, i, &row, &col, &match_count) != 0) {
        // Skip entries the C ABI declines to materialise; this is purely
        // defensive — the index is always valid by construction.
        continue;
      }
      cell.row = row;
      cell.col = col;
      cell.matches.reserve(match_count);
      for (std::size_t j = 0; j < match_count; ++j) {
        fm_cf_match_t m{};
        if (fm_cf_results_match_at(results, i, j, &m) != 0) {
          continue;
        }
        cell.matches.push_back(translate_cf_match(m));
      }
      r.cells.push_back(std::move(cell));
    }
    fm_cf_results_destroy(results);
    r.status = ok_status();
    return r;
  }

  // ---- Sheet view / layout ------------------------------------------------

  /// Reads the per-sheet view (zoom, frozen panes, tab visibility).
  JsSheetViewResult getSheetView(uint32_t sheet) const {
    JsSheetViewResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    fm_sheet_view_t v{};
    fm_status_t rc = fm_sheet_get_view(handle_, sheet, &v);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.view.zoomScale = v.zoom_scale;
    r.view.freezeRows = v.freeze_rows;
    r.view.freezeCols = v.freeze_cols;
    r.view.tabHidden = v.tab_hidden;
    r.status = ok_status();
    return r;
  }

  /// Sets the sheet's zoom percentage. Out-of-range values are clamped
  /// to `[10, 400]`.
  JsStatus setSheetZoom(uint32_t sheet, uint32_t zoomScale) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_set_zoom(handle_, sheet, zoomScale);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets the sheet's frozen pane in `(rows, cols)`.
  JsStatus setSheetFreeze(uint32_t sheet, uint32_t freezeRows, uint32_t freezeCols) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_set_freeze(handle_, sheet, freezeRows, freezeCols);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets the sheet tab's hidden flag.
  JsStatus setSheetTabHidden(uint32_t sheet, bool hidden) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_set_tab_hidden(handle_, sheet, hidden ? 1 : 0);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Returns the column-layout overrides on `sheet` in storage order.
  JsColumnsResult getSheetColumns(uint32_t sheet) const {
    JsColumnsResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    std::size_t count = 0;
    fm_status_t rc = fm_sheet_get_column_count(handle_, sheet, &count);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.columns.reserve(count);
    for (std::size_t i = 0; i < count; ++i) {
      fm_column_layout_t entry{};
      if (fm_sheet_get_column(handle_, sheet, i, &entry) != 0) {
        continue;
      }
      JsColumnLayout out;
      out.first = entry.first;
      out.last = entry.last;
      out.width = entry.width;
      out.hidden = entry.hidden;
      out.outlineLevel = static_cast<int32_t>(entry.outline_level);
      r.columns.push_back(out);
    }
    r.status = ok_status();
    return r;
  }

  /// Sets / replaces the column width override on `[first, last]`.
  JsStatus setColumnWidth(uint32_t sheet, uint32_t first, uint32_t last, double width) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_set_column_width(handle_, sheet, first, last, width);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets / replaces the column hidden flag on `[first, last]`.
  JsStatus setColumnHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_set_column_hidden(handle_, sheet, first, last, hidden ? 1 : 0);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets / replaces the column outline level on `[first, last]`.
  JsStatus setColumnOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    if (level > 255U) {
      level = 255U;
    }
    fm_status_t rc = fm_sheet_set_column_outline(handle_, sheet, first, last, static_cast<uint8_t>(level));
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Returns the row-layout overrides on `sheet` in storage order.
  JsRowsResult getSheetRowOverrides(uint32_t sheet) const {
    JsRowsResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    std::size_t count = 0;
    fm_status_t rc = fm_sheet_get_row_override_count(handle_, sheet, &count);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.rows.reserve(count);
    for (std::size_t i = 0; i < count; ++i) {
      fm_row_layout_t entry{};
      if (fm_sheet_get_row_override(handle_, sheet, i, &entry) != 0) {
        continue;
      }
      JsRowLayout out;
      out.row = entry.row;
      out.height = entry.height;
      out.hidden = entry.hidden;
      out.outlineLevel = static_cast<int32_t>(entry.outline_level);
      r.rows.push_back(out);
    }
    r.status = ok_status();
    return r;
  }

  /// Sets / replaces the row height override at `row`.
  JsStatus setRowHeight(uint32_t sheet, uint32_t row, double height) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_set_row_height(handle_, sheet, row, height);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets / replaces the row hidden flag at `row`.
  JsStatus setRowHidden(uint32_t sheet, uint32_t row, bool hidden) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_set_row_hidden(handle_, sheet, row, hidden ? 1 : 0);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Sets / replaces the row outline level at `row`.
  JsStatus setRowOutline(uint32_t sheet, uint32_t row, uint32_t level) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    if (level > 255U) {
      level = 255U;
    }
    fm_status_t rc = fm_sheet_set_row_outline(handle_, sheet, row, static_cast<uint8_t>(level));
    return rc == 0 ? ok_status() : error_status(rc);
  }

  // ---- Styles --------------------------------------------------------------
  // Three thin wrappers around the cell-level xf accessors plus the
  // `fm_styles_get_cell_xf` projector. Font / num-fmt-string getters
  // are intentionally not surfaced here: JS callers can derive those
  // through the C API directly when needed, and skipping them keeps
  // the embind surface small.

  /// Returns `{ status, xfIndex }` for the cell at `(sheet, row, col)`.
  emscripten::val getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    uint32_t xf = 0;
    fm_status_t rc = fm_cell_get_xf_index(handle_, sheet, row, col, &xf);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("xfIndex", xf);
    return o;
  }

  /// Persists `xfIndex` on the cell at `(sheet, row, col)`.
  JsStatus setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_cell_set_xf_index(handle_, sheet, row, col, xf_index);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// Returns `{ status, fontIndex, fillIndex, borderIndex, numFmtId,
  /// horizontalAlign, verticalAlign, wrapText }` for the `xfIndex`-th
  /// xf record.
  emscripten::val getCellXf(uint32_t xf_index) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    fm_cell_xf xf{};
    fm_status_t rc = fm_styles_get_cell_xf(handle_, xf_index, &xf);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("fontIndex", xf.font_index);
    o.set("fillIndex", xf.fill_index);
    o.set("borderIndex", xf.border_index);
    o.set("numFmtId", static_cast<uint32_t>(xf.num_fmt_id));
    o.set("horizontalAlign", static_cast<uint32_t>(xf.horizontal_align));
    o.set("verticalAlign", static_cast<uint32_t>(xf.vertical_align));
    o.set("wrapText", xf.wrap_text != 0);
    return o;
  }

  // ---- Sheet UI features (merges, hyperlinks, comments, validations) ------
  // Each accessor returns a JS-friendly value (Array<...> or null) so JS
  // callers don't need to step through count + getter pairs.

  /// `addMerge(sheetIdx, {firstRow, lastRow, firstCol, lastCol})`.
  JsStatus addMerge(uint32_t sheet, emscripten::val range) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_merge_range m;
    m.first_row = range["firstRow"].as<uint32_t>();
    m.last_row = range["lastRow"].as<uint32_t>();
    m.first_col = range["firstCol"].as<uint32_t>();
    m.last_col = range["lastCol"].as<uint32_t>();
    fm_status_t rc = fm_sheet_add_merge(handle_, sheet, m);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `getComment(sheetIdx, row, col) -> {author, text} | null`.
  emscripten::val getComment(uint32_t sheet, uint32_t row, uint32_t col) const {
    if (handle_ == nullptr) {
      return emscripten::val::null();
    }
    fm_comment c{};
    if (fm_sheet_get_comment_at(handle_, sheet, row, col, &c) != 0) {
      return emscripten::val::null();
    }
    emscripten::val o = emscripten::val::object();
    o.set("author", c.author != nullptr ? std::string(c.author) : std::string());
    o.set("text", c.text != nullptr ? std::string(c.text) : std::string());
    return o;
  }

  /// `getHyperlinks(sheetIdx) -> Array<{row, col, target, display, tooltip}>`.
  emscripten::val getHyperlinks(uint32_t sheet) const {
    emscripten::val arr = emscripten::val::array();
    if (handle_ == nullptr) {
      return arr;
    }
    uint32_t count = 0;
    if (fm_sheet_get_hyperlink_count(handle_, sheet, &count) != 0) {
      return arr;
    }
    for (uint32_t i = 0; i < count; ++i) {
      fm_hyperlink h{};
      if (fm_sheet_get_hyperlink_at(handle_, sheet, i, &h) != 0) {
        continue;
      }
      emscripten::val item = emscripten::val::object();
      item.set("row", h.row);
      item.set("col", h.col);
      item.set("target", h.target != nullptr ? std::string(h.target) : std::string());
      item.set("display", h.display != nullptr ? std::string(h.display) : std::string());
      item.set("tooltip", h.tooltip != nullptr ? std::string(h.tooltip) : std::string());
      arr.set(i, item);
    }
    return arr;
  }

  /// `getMerges(sheetIdx) -> Array<{firstRow, lastRow, firstCol, lastCol}>`.
  emscripten::val getMerges(uint32_t sheet) const {
    emscripten::val arr = emscripten::val::array();
    if (handle_ == nullptr) {
      return arr;
    }
    uint32_t count = 0;
    if (fm_sheet_get_merge_count(handle_, sheet, &count) != 0) {
      return arr;
    }
    for (uint32_t i = 0; i < count; ++i) {
      fm_merge_range m{};
      if (fm_sheet_get_merge_at(handle_, sheet, i, &m) != 0) {
        continue;
      }
      emscripten::val item = emscripten::val::object();
      item.set("firstRow", m.first_row);
      item.set("lastRow", m.last_row);
      item.set("firstCol", m.first_col);
      item.set("lastCol", m.last_col);
      arr.set(i, item);
    }
    return arr;
  }

  /// `getValidations(sheetIdx) -> Array<{ranges, type, op, formula1, formula2, errorMessage}>`.
  /// Read-only: full mutators land in a follow-up.
  emscripten::val getValidations(uint32_t sheet) const {
    emscripten::val arr = emscripten::val::array();
    if (handle_ == nullptr) {
      return arr;
    }
    // The C ABI does not yet surface validation entries (next bundle);
    // fall through to the engine via the binding's existing handle by
    // reaching into the workbook directly is not allowed here, so we
    // currently return an empty array. JS callers can detect "feature
    // unavailable" via length === 0 paired with an absent comments
    // path. The structural contract stays stable for the follow-up
    // bundle.
    (void)sheet;
    return arr;
  }

  /// `setComment(sheetIdx, row, col, author, text)` (empty text removes).
  JsStatus setComment(uint32_t sheet, uint32_t row, uint32_t col, const std::string& author, const std::string& text) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    const char* author_c = author.empty() ? nullptr : author.c_str();
    const char* text_c = text.empty() ? nullptr : text.c_str();
    fm_status_t rc = fm_sheet_set_comment(handle_, sheet, row, col, author_c, text_c);
    return rc == 0 ? ok_status() : error_status(rc);
  }

 private:
  // Holder for the currently-installed JS progress callback. Function
  // local static keeps the slot alive for the WASM module's lifetime
  // without needing a global variable; embind's val type is not safe
  // to default-initialise at static-init time.
  static emscripten::val& js_progress_callback() {
    static emscripten::val cb = emscripten::val::null();
    return cb;
  }

  // C-ABI compatible trampoline that forwards to the held JS callback.
  // Returning `false` from the JS side aborts the iterative solve.
  static bool iterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                          void* /*user_data*/) {
    emscripten::val& cb = js_progress_callback();
    if (cb.isNull() || cb.isUndefined()) {
      return true;
    }
    emscripten::val ret = cb(iteration, max_residual, max_iterations);
    if (ret.isUndefined() || ret.isNull()) {
      return true;
    }
    return ret.as<bool>();
  }

  fm_workbook_t* handle_ = nullptr;
};

/// Convenience: evaluate a single formula in a fresh workbook.
///
/// Mirrors the `formulon_cli eval <formula>` workflow: spin up an empty
/// workbook with the default Sheet1, place the formula at A1, recalc,
/// and return the resulting cell value.
JsEvalResult eval_formula(const std::string& formula) {
  JsEvalResult r;
  fm_workbook_t* wb = nullptr;
  fm_status_t rc = fm_workbook_create(&wb);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  rc = fm_workbook_set_formula(wb, 0, 0, 0, formula.c_str());
  if (rc != 0) {
    r.status = error_status(rc);
    fm_workbook_destroy(wb);
    return r;
  }
  rc = fm_workbook_recalc(wb);
  if (rc != 0) {
    r.status = error_status(rc);
    fm_workbook_destroy(wb);
    return r;
  }
  fm_value_t v{};
  rc = fm_workbook_get_value(wb, 0, 0, 0, &v);
  if (rc != 0) {
    r.status = error_status(rc);
    fm_workbook_destroy(wb);
    return r;
  }
  r.value = translate_value(v);
  r.status = ok_status();
  fm_workbook_destroy(wb);
  return r;
}

/// Returns the Formulon library version (NUL-terminated UTF-8).
std::string version_string() {
  const char* s = fm_version_string();
  return s != nullptr ? std::string(s) : std::string();
}

/// Returns the static C string for `status`.
std::string status_string(int32_t status) {
  const char* s = fm_status_string(static_cast<fm_status_t>(status));
  return s != nullptr ? std::string(s) : std::string();
}

/// Returns the most-recent thread-local error message.
std::string last_error_message() {
  const char* s = fm_last_error_message();
  return s != nullptr ? std::string(s) : std::string();
}

/// Returns the most-recent thread-local error context.
std::string last_error_context() {
  const char* s = fm_last_error_context();
  return s != nullptr ? std::string(s) : std::string();
}

}  // namespace

EMSCRIPTEN_BINDINGS(formulon) {
  using emscripten::allow_raw_pointers;
  using emscripten::class_;
  using emscripten::function;
  using emscripten::register_vector;
  using emscripten::value_object;

  // ---- Value-object surface ------------------------------------------------
  value_object<JsStatus>("Status")
      .field("ok", &JsStatus::ok)
      .field("status", &JsStatus::status)
      .field("message", &JsStatus::message)
      .field("context", &JsStatus::context);

  value_object<JsValue>("Value")
      .field("kind", &JsValue::kind)
      .field("number", &JsValue::number)
      .field("boolean", &JsValue::boolean)
      .field("text", &JsValue::text)
      .field("errorCode", &JsValue::errorCode);

  value_object<JsCellResult>("CellResult").field("status", &JsCellResult::status).field("value", &JsCellResult::value);

  value_object<JsEvalResult>("EvalResult").field("status", &JsEvalResult::status).field("value", &JsEvalResult::value);

  value_object<JsSaveResult>("SaveResult").field("status", &JsSaveResult::status).field("bytes", &JsSaveResult::bytes);

  value_object<JsStringResult>("StringResult")
      .field("status", &JsStringResult::status)
      .field("value", &JsStringResult::value);

  // ---- Conditional-format value-objects ------------------------------------
  // The vector-of-value-object classes below (`CfMatchVector`,
  // `CfCellVector`) surface as iterable handles in JS with `.size()` and
  // `.get(i)` accessors. They mirror how embind exposes
  // `register_vector<T>` for any value-object payload.
  value_object<JsCfColor>("CfColor")
      .field("r", &JsCfColor::r)
      .field("g", &JsCfColor::g)
      .field("b", &JsCfColor::b)
      .field("a", &JsCfColor::a);

  value_object<JsCfMatch>("CfMatch")
      .field("kind", &JsCfMatch::kind)
      .field("priority", &JsCfMatch::priority)
      .field("dxfIdEngaged", &JsCfMatch::dxfIdEngaged)
      .field("dxfId", &JsCfMatch::dxfId)
      .field("color", &JsCfMatch::color)
      .field("barLengthPct", &JsCfMatch::barLengthPct)
      .field("barAxisPositionPct", &JsCfMatch::barAxisPositionPct)
      .field("barIsNegative", &JsCfMatch::barIsNegative)
      .field("barFill", &JsCfMatch::barFill)
      .field("barBorderEngaged", &JsCfMatch::barBorderEngaged)
      .field("barBorder", &JsCfMatch::barBorder)
      .field("barGradient", &JsCfMatch::barGradient)
      .field("iconSetName", &JsCfMatch::iconSetName)
      .field("iconIndex", &JsCfMatch::iconIndex);

  register_vector<JsCfMatch>("CfMatchVector");

  value_object<JsCfCellResult>("CfCellResult")
      .field("row", &JsCfCellResult::row)
      .field("col", &JsCfCellResult::col)
      .field("matches", &JsCfCellResult::matches);

  register_vector<JsCfCellResult>("CfCellVector");

  value_object<JsCfRangeResult>("CfRangeResult")
      .field("status", &JsCfRangeResult::status)
      .field("cells", &JsCfRangeResult::cells);

  // ---- Sheet view / layout value-objects -----------------------------------
  value_object<JsSheetView>("SheetView")
      .field("zoomScale", &JsSheetView::zoomScale)
      .field("freezeRows", &JsSheetView::freezeRows)
      .field("freezeCols", &JsSheetView::freezeCols)
      .field("tabHidden", &JsSheetView::tabHidden);

  value_object<JsSheetViewResult>("SheetViewResult")
      .field("status", &JsSheetViewResult::status)
      .field("view", &JsSheetViewResult::view);

  value_object<JsColumnLayout>("ColumnLayout")
      .field("first", &JsColumnLayout::first)
      .field("last", &JsColumnLayout::last)
      .field("width", &JsColumnLayout::width)
      .field("hidden", &JsColumnLayout::hidden)
      .field("outlineLevel", &JsColumnLayout::outlineLevel);

  register_vector<JsColumnLayout>("ColumnLayoutVector");

  value_object<JsColumnsResult>("ColumnsResult")
      .field("status", &JsColumnsResult::status)
      .field("columns", &JsColumnsResult::columns);

  value_object<JsRowLayout>("RowLayout")
      .field("row", &JsRowLayout::row)
      .field("height", &JsRowLayout::height)
      .field("hidden", &JsRowLayout::hidden)
      .field("outlineLevel", &JsRowLayout::outlineLevel);

  register_vector<JsRowLayout>("RowLayoutVector");

  value_object<JsRowsResult>("RowsResult").field("status", &JsRowsResult::status).field("rows", &JsRowsResult::rows);

  // ---- Workbook class ------------------------------------------------------
  class_<JsWorkbook>("Workbook")
      .class_function("createDefault", &JsWorkbook::createDefault, allow_raw_pointers())
      .class_function("createEmpty", &JsWorkbook::createEmpty, allow_raw_pointers())
      .class_function("loadBytes", &JsWorkbook::loadBytes, allow_raw_pointers())
      .function("addMerge", &JsWorkbook::addMerge)
      .function("addSheet", &JsWorkbook::addSheet)
      .function("cellAt", &JsWorkbook::cellAt)
      .function("cellCount", &JsWorkbook::cellCount)
      .function("definedNameAt", &JsWorkbook::definedNameAt)
      .function("definedNameCount", &JsWorkbook::definedNameCount)
      .function("deleteCols", &JsWorkbook::deleteCols)
      .function("deleteRows", &JsWorkbook::deleteRows)
      .function("evaluateCfRange", &JsWorkbook::evaluateCfRange)
      .function("getCellXf", &JsWorkbook::getCellXf)
      .function("getCellXfIndex", &JsWorkbook::getCellXfIndex)
      .function("getComment", &JsWorkbook::getComment)
      .function("getHyperlinks", &JsWorkbook::getHyperlinks)
      .function("getMerges", &JsWorkbook::getMerges)
      .function("getSheetColumns", &JsWorkbook::getSheetColumns)
      .function("getSheetRowOverrides", &JsWorkbook::getSheetRowOverrides)
      .function("getSheetView", &JsWorkbook::getSheetView)
      .function("getValidations", &JsWorkbook::getValidations)
      .function("getValue", &JsWorkbook::getValue)
      .function("insertCols", &JsWorkbook::insertCols)
      .function("insertRows", &JsWorkbook::insertRows)
      .function("moveSheet", &JsWorkbook::moveSheet)
      .function("partialRecalc", &JsWorkbook::partialRecalc)
      .function("passthroughAt", &JsWorkbook::passthroughAt)
      .function("passthroughCount", &JsWorkbook::passthroughCount)
      .function("recalc", &JsWorkbook::recalc)
      .function("removeSheet", &JsWorkbook::removeSheet)
      .function("renameSheet", &JsWorkbook::renameSheet)
      .function("setBlank", &JsWorkbook::setBlank)
      .function("setBool", &JsWorkbook::setBool)
      .function("setCellXfIndex", &JsWorkbook::setCellXfIndex)
      .function("setColumnHidden", &JsWorkbook::setColumnHidden)
      .function("setColumnOutline", &JsWorkbook::setColumnOutline)
      .function("setColumnWidth", &JsWorkbook::setColumnWidth)
      .function("setComment", &JsWorkbook::setComment)
      .function("setDefinedName", &JsWorkbook::setDefinedName)
      .function("setFormula", &JsWorkbook::setFormula)
      .function("setIterative", &JsWorkbook::setIterative)
      .function("setIterativeProgress", &JsWorkbook::setIterativeProgress)
      .function("setNumber", &JsWorkbook::setNumber)
      .function("setRowHeight", &JsWorkbook::setRowHeight)
      .function("setRowHidden", &JsWorkbook::setRowHidden)
      .function("setRowOutline", &JsWorkbook::setRowOutline)
      .function("setSheetFreeze", &JsWorkbook::setSheetFreeze)
      .function("setSheetTabHidden", &JsWorkbook::setSheetTabHidden)
      .function("setSheetZoom", &JsWorkbook::setSheetZoom)
      .function("setText", &JsWorkbook::setText)
      .function("sheetCount", &JsWorkbook::sheetCount)
      .function("sheetName", &JsWorkbook::sheetName)
      .function("tableAt", &JsWorkbook::tableAt)
      .function("tableCount", &JsWorkbook::tableCount);

  // ---- Free functions ------------------------------------------------------
  function("evalFormula", &eval_formula);
  function("versionString", &version_string);
  function("statusString", &status_string);
  function("lastErrorMessage", &last_error_message);
  function("lastErrorContext", &last_error_context);
}
