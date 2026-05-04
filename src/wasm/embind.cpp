// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

/// JS-side mirror of `fm_sheet_protection_t`. Strings are passed as
/// `std::string` so embind handles the lifetime; the pointers from
/// `fm_sheet_get_protection` are copied into these fields before the
/// value crosses the WASM boundary, so JS code can hold them
/// indefinitely without aliasing the workbook's storage.
struct JsSheetProtection {
  int32_t enabled = 0;
  std::string algorithmName;
  std::string hashValue;
  std::string saltValue;
  uint32_t spinCount = 0U;
  std::string legacyPassword;
  int32_t sheet = 0;
  int32_t objects = 0;
  int32_t scenarios = 0;
  int32_t formatCells = 0;
  int32_t formatColumns = 0;
  int32_t formatRows = 0;
  int32_t insertColumns = 0;
  int32_t insertRows = 0;
  int32_t insertHyperlinks = 0;
  int32_t deleteColumns = 0;
  int32_t deleteRows = 0;
  int32_t selectLockedCells = 0;
  int32_t selectUnlockedCells = 0;
  int32_t sort = 0;
  int32_t autoFilter = 0;
  int32_t pivotTables = 0;
};

/// Return envelope for `Workbook.getSheetProtection(...)`.
struct JsSheetProtectionResult {
  JsStatus status;
  JsSheetProtection protection{};
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

/// Return envelope for `Workbook.addFont` / `addFill` / `addBorder` /
/// `addXf`. The `index` field carries the resolved (existing or new)
/// table index.
struct JsAddStyleResult {
  JsStatus status;
  uint32_t index = 0U;
};

/// Return envelope for `Workbook.addNumFmt`. `numFmtId` is either the
/// matched built-in id (`0..163`) or the freshly-assigned custom id
/// (`>= 164`).
struct JsAddNumFmtResult {
  JsStatus status;
  uint32_t numFmtId = 0U;
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

// ---- JS field-extraction helpers ----------------------------------------
//
// These small helpers replace the repetitive
// `record["x"].isUndefined() ? dflt : record["x"].as<T>()` pattern that
// previously appeared inside addFont / addFill / addBorder / addXf /
// addValidation. Centralising them lets the compiler emit one copy of the
// embind glue (val::operator[], val::isUndefined, val::as<T>) per field
// type instead of one per call site, which is a measurable WASM size win
// because every embind operation pulls in non-trivial JS-bridge stubs.
//
// They are intentionally `inline` (not `noinline`): empirically, letting
// `-Oz` keep them inlined yields a smaller `.wasm.br` than forcing a
// call. The dedup win comes from the helpers themselves being shared
// between sites, not from inhibiting inlining.

/// Returns the `uint32_t` value of `v[key]`, or `dflt` when the field is
/// missing / undefined / null.
inline uint32_t js_pull_u32(const emscripten::val& v, const char* key, uint32_t dflt) {
  emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return dflt;
  }
  return f.as<uint32_t>();
}

/// Returns the low-byte `uint8_t` value of `v[key]`, or `dflt` when missing.
inline uint8_t js_pull_u8(const emscripten::val& v, const char* key, uint8_t dflt) {
  return static_cast<uint8_t>(js_pull_u32(v, key, dflt) & 0xFFU);
}

/// Returns the `uint16_t` value of `v[key]`, or `dflt` when missing.
inline uint16_t js_pull_u16(const emscripten::val& v, const char* key, uint16_t dflt) {
  return static_cast<uint16_t>(js_pull_u32(v, key, dflt) & 0xFFFFU);
}

/// Returns the `double` value of `v[key]`, or `dflt` when missing.
inline double js_pull_double(const emscripten::val& v, const char* key, double dflt) {
  emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return dflt;
  }
  return f.as<double>();
}

/// Returns the `bool` value of `v[key]`, or `dflt` when missing.
inline bool js_pull_bool(const emscripten::val& v, const char* key, bool dflt) {
  emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return dflt;
  }
  return f.as<bool>();
}

/// Returns the string value of `v[key]`, or an empty string when missing.
inline std::string js_pull_string(const emscripten::val& v, const char* key) {
  emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return std::string();
  }
  return f.as<std::string>();
}

/// Pulls a `{style, colorArgb}` border-side record out of `v`, defaulting
/// every absent field to zero.
inline fm_border_side js_pull_border_side(const emscripten::val& v) {
  fm_border_side s{};
  if (v.isUndefined() || v.isNull()) {
    return s;
  }
  s.style = js_pull_u8(v, "style", 0);
  s.color_argb = js_pull_u32(v, "colorArgb", 0U);
  return s;
}

/// Builds a `{firstRow, lastRow, firstCol, lastCol}` JS object from a
/// merge / validation range record.
inline emscripten::val merge_range_to_val(const fm_merge_range& m) {
  emscripten::val item = emscripten::val::object();
  item.set("firstRow", m.first_row);
  item.set("lastRow", m.last_row);
  item.set("firstCol", m.first_col);
  item.set("lastCol", m.last_col);
  return item;
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

  /// Renders the lambda closure stored at `(sheet, row, col)` as the
  /// surface form `LAMBDA(p1,p2,body)` (no leading `=`). Returns
  /// `{ status, text }`; `kInvalidArgument` when the cell is absent or
  /// does not hold a lambda value.
  emscripten::val getLambdaText(uint32_t sheet, uint32_t row, uint32_t col) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      o.set("text", std::string());
      return o;
    }
    const char* text = nullptr;
    fm_status_t rc = fm_workbook_lambda_text_at(handle_, sheet, row, col, &text);
    if (rc != 0) {
      o.set("status", error_status(rc));
      o.set("text", std::string());
      return o;
    }
    o.set("status", ok_status());
    o.set("text", std::string(text != nullptr ? text : ""));
    return o;
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

  /// Returns the workbook's calc mode as the underlying enum value:
  /// `0` = auto, `1` = manual, `2` = autoNoTable. Defaults to `0` for an
  /// invalid handle so the JS side never sees an error envelope here.
  uint32_t calcMode() const {
    if (handle_ == nullptr) {
      return static_cast<uint32_t>(FM_CALC_MODE_AUTO);
    }
    fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
    fm_workbook_calc_mode(handle_, &mode);
    return static_cast<uint32_t>(mode);
  }

  /// Sets the workbook's calc mode. Accepts the same enum codes as
  /// `calcMode()`. Returns `kInvalidArgument` for unknown values.
  JsStatus setCalcMode(uint32_t mode) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_workbook_set_calc_mode(handle_, static_cast<fm_calc_mode_t>(mode));
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

  /// Reads the per-sheet `<sheetProtection>` flags. Strings are
  /// deep-copied so the returned object is independent of the
  /// workbook's storage.
  JsSheetProtectionResult getSheetProtection(uint32_t sheet) const {
    JsSheetProtectionResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    fm_sheet_protection_t p{};
    fm_status_t rc = fm_sheet_get_protection(handle_, sheet, &p);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.protection.enabled = p.enabled;
    r.protection.algorithmName = p.algorithm_name == nullptr ? std::string() : p.algorithm_name;
    r.protection.hashValue = p.hash_value == nullptr ? std::string() : p.hash_value;
    r.protection.saltValue = p.salt_value == nullptr ? std::string() : p.salt_value;
    r.protection.spinCount = p.spin_count;
    r.protection.legacyPassword = p.legacy_password == nullptr ? std::string() : p.legacy_password;
    r.protection.sheet = p.sheet;
    r.protection.objects = p.objects;
    r.protection.scenarios = p.scenarios;
    r.protection.formatCells = p.format_cells;
    r.protection.formatColumns = p.format_columns;
    r.protection.formatRows = p.format_rows;
    r.protection.insertColumns = p.insert_columns;
    r.protection.insertRows = p.insert_rows;
    r.protection.insertHyperlinks = p.insert_hyperlinks;
    r.protection.deleteColumns = p.delete_columns;
    r.protection.deleteRows = p.delete_rows;
    r.protection.selectLockedCells = p.select_locked_cells;
    r.protection.selectUnlockedCells = p.select_unlocked_cells;
    r.protection.sort = p.sort;
    r.protection.autoFilter = p.auto_filter;
    r.protection.pivotTables = p.pivot_tables;
    r.status = ok_status();
    return r;
  }

  /// Replaces the per-sheet `<sheetProtection>` flags wholesale.
  JsStatus setSheetProtection(uint32_t sheet, JsSheetProtection in) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_sheet_protection_t p{};
    p.enabled = in.enabled;
    p.algorithm_name = in.algorithmName.c_str();
    p.hash_value = in.hashValue.c_str();
    p.salt_value = in.saltValue.c_str();
    p.spin_count = in.spinCount;
    p.legacy_password = in.legacyPassword.c_str();
    p.sheet = in.sheet;
    p.objects = in.objects;
    p.scenarios = in.scenarios;
    p.format_cells = in.formatCells;
    p.format_columns = in.formatColumns;
    p.format_rows = in.formatRows;
    p.insert_columns = in.insertColumns;
    p.insert_rows = in.insertRows;
    p.insert_hyperlinks = in.insertHyperlinks;
    p.delete_columns = in.deleteColumns;
    p.delete_rows = in.deleteRows;
    p.select_locked_cells = in.selectLockedCells;
    p.select_unlocked_cells = in.selectUnlockedCells;
    p.sort = in.sort;
    p.auto_filter = in.autoFilter;
    p.pivot_tables = in.pivotTables;
    fm_status_t rc = fm_sheet_set_protection(handle_, sheet, &p);
    return rc == 0 ? ok_status() : error_status(rc);
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

  /// Returns the resolved `<font>` record at `font_index`. Shape mirrors
  /// the read-side `fm_font_record`.
  emscripten::val getFont(uint32_t font_index) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    fm_font_record f{};
    fm_status_t rc = fm_styles_get_font(handle_, font_index, &f);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("name", std::string(f.name != nullptr ? f.name : ""));
    o.set("size", f.size);
    o.set("colorArgb", f.color_argb);
    o.set("bold", f.bold != 0);
    o.set("italic", f.italic != 0);
    o.set("strike", f.strike != 0);
    o.set("underline", static_cast<uint32_t>(f.underline));
    return o;
  }

  /// Returns the resolved `<fill>` record at `fill_index`.
  emscripten::val getFill(uint32_t fill_index) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    fm_fill_record f{};
    fm_status_t rc = fm_styles_get_fill(handle_, fill_index, &f);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("pattern", static_cast<uint32_t>(f.pattern));
    o.set("fgArgb", f.fg_argb);
    o.set("bgArgb", f.bg_argb);
    return o;
  }

  /// Returns the resolved `<border>` record at `border_index`. Each
  /// side appears as a `{ style, colorArgb }` object.
  emscripten::val getBorder(uint32_t border_index) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    fm_border_record b{};
    fm_status_t rc = fm_styles_get_border(handle_, border_index, &b);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    auto side_obj = [](const fm_border_side& s) {
      emscripten::val v = emscripten::val::object();
      v.set("style", static_cast<uint32_t>(s.style));
      v.set("colorArgb", s.color_argb);
      return v;
    };
    o.set("status", ok_status());
    o.set("left", side_obj(b.left));
    o.set("right", side_obj(b.right));
    o.set("top", side_obj(b.top));
    o.set("bottom", side_obj(b.bottom));
    o.set("diagonal", side_obj(b.diagonal));
    o.set("diagonalUp", b.diagonal_up != 0);
    o.set("diagonalDown", b.diagonal_down != 0);
    return o;
  }

  /// Returns the format string for a built-in or custom `num_fmt_id`.
  emscripten::val getNumFmt(uint32_t num_fmt_id) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    const char* s = nullptr;
    fm_status_t rc = fm_styles_get_num_fmt_string(handle_, static_cast<uint16_t>(num_fmt_id), &s);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("numFmtId", num_fmt_id);
    o.set("formatCode", std::string(s != nullptr ? s : ""));
    return o;
  }

  /// Adds (or dedups against an existing entry) a font record. JS-side
  /// shape: `{ name, size, bold, italic, strike, underline, colorArgb }`.
  JsAddStyleResult addFont(emscripten::val record) {
    JsAddStyleResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    const std::string name = js_pull_string(record, "name");
    fm_font_record fr{};
    fr.name = name.c_str();
    fr.size = js_pull_double(record, "size", 11.0);
    fr.bold = js_pull_bool(record, "bold", false) ? 1 : 0;
    fr.italic = js_pull_bool(record, "italic", false) ? 1 : 0;
    fr.strike = js_pull_bool(record, "strike", false) ? 1 : 0;
    fr.underline = js_pull_u8(record, "underline", 0U);
    fr.color_argb = js_pull_u32(record, "colorArgb", 0xFF000000U);
    uint32_t idx = 0;
    fm_status_t rc = fm_styles_add_font(handle_, fr, &idx);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.status = ok_status();
    r.index = idx;
    return r;
  }

  /// Adds (or dedups against an existing entry) a fill record. JS-side
  /// shape: `{ pattern, fgArgb, bgArgb }`.
  JsAddStyleResult addFill(emscripten::val record) {
    JsAddStyleResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    fm_fill_record fr{};
    fr.pattern = js_pull_u8(record, "pattern", 0U);
    fr.fg_argb = js_pull_u32(record, "fgArgb", 0U);
    fr.bg_argb = js_pull_u32(record, "bgArgb", 0U);
    uint32_t idx = 0;
    fm_status_t rc = fm_styles_add_fill(handle_, fr, &idx);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.status = ok_status();
    r.index = idx;
    return r;
  }

  /// Adds (or dedups against an existing entry) a border record. JS-side
  /// shape: `{ left:{style,colorArgb}, right, top, bottom, diagonal,
  /// diagonalUp, diagonalDown }`. Missing sides default to `{0,0}`.
  JsAddStyleResult addBorder(emscripten::val record) {
    JsAddStyleResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    fm_border_record br{};
    br.left = js_pull_border_side(record["left"]);
    br.right = js_pull_border_side(record["right"]);
    br.top = js_pull_border_side(record["top"]);
    br.bottom = js_pull_border_side(record["bottom"]);
    br.diagonal = js_pull_border_side(record["diagonal"]);
    br.diagonal_up = js_pull_bool(record, "diagonalUp", false) ? 1 : 0;
    br.diagonal_down = js_pull_bool(record, "diagonalDown", false) ? 1 : 0;
    uint32_t idx = 0;
    fm_status_t rc = fm_styles_add_border(handle_, br, &idx);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.status = ok_status();
    r.index = idx;
    return r;
  }

  /// Adds (or dedups against an existing entry) a number-format entry.
  /// `formatCode` matching a built-in id returns the built-in id; a
  /// custom code is appended at `max(existing_custom_id, 163) + 1`.
  JsAddNumFmtResult addNumFmt(const std::string& format_code) {
    JsAddNumFmtResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    uint16_t id = 0;
    fm_status_t rc = fm_styles_add_num_fmt(handle_, format_code.c_str(), &id);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.status = ok_status();
    r.numFmtId = id;
    return r;
  }

  /// Adds (or dedups against an existing entry) an `<xf>` record.
  /// JS-side shape: `{ fontIndex, fillIndex, borderIndex, numFmtId,
  /// horizontalAlign, verticalAlign, wrapText }`.
  JsAddStyleResult addXf(emscripten::val record) {
    JsAddStyleResult r;
    if (handle_ == nullptr) {
      r.status = error_status(7000);
      return r;
    }
    fm_cell_xf xf{};
    xf.font_index = js_pull_u32(record, "fontIndex", 0U);
    xf.fill_index = js_pull_u32(record, "fillIndex", 0U);
    xf.border_index = js_pull_u32(record, "borderIndex", 0U);
    xf.num_fmt_id = js_pull_u16(record, "numFmtId", 0U);
    xf.horizontal_align = js_pull_u8(record, "horizontalAlign", 0U);
    xf.vertical_align = js_pull_u8(record, "verticalAlign", 0U);
    xf.wrap_text = js_pull_bool(record, "wrapText", false) ? 1 : 0;
    uint32_t idx = 0;
    fm_status_t rc = fm_styles_add_cell_xf(handle_, xf, &idx);
    if (rc != 0) {
      r.status = error_status(rc);
      return r;
    }
    r.status = ok_status();
    r.index = idx;
    return r;
  }

  /// Returns the number of font records currently registered.
  uint32_t fontCount() const {
    if (handle_ == nullptr) {
      return 0U;
    }
    uint32_t n = 0;
    if (fm_styles_get_font_count(handle_, &n) != 0) {
      return 0U;
    }
    return n;
  }

  /// Returns the number of fill records currently registered.
  uint32_t fillCount() const {
    if (handle_ == nullptr) {
      return 0U;
    }
    uint32_t n = 0;
    if (fm_styles_get_fill_count(handle_, &n) != 0) {
      return 0U;
    }
    return n;
  }

  /// Returns the number of border records currently registered.
  uint32_t borderCount() const {
    if (handle_ == nullptr) {
      return 0U;
    }
    uint32_t n = 0;
    if (fm_styles_get_border_count(handle_, &n) != 0) {
      return 0U;
    }
    return n;
  }

  /// Returns the number of `<xf>` records currently registered.
  uint32_t xfCount() const {
    if (handle_ == nullptr) {
      return 0U;
    }
    uint32_t n = 0;
    if (fm_styles_get_cell_xf_count(handle_, &n) != 0) {
      return 0U;
    }
    return n;
  }

  /// Returns the number of named cell styles registered.
  uint32_t cellStyleCount() const {
    if (handle_ == nullptr) {
      return 0U;
    }
    uint32_t n = 0;
    if (fm_styles_get_cell_style_count(handle_, &n) != 0) {
      return 0U;
    }
    return n;
  }

  /// Returns the number of `<cellStyleXfs>` records (named-style xf table).
  uint32_t cellStyleXfCount() const {
    if (handle_ == nullptr) {
      return 0U;
    }
    uint32_t n = 0;
    if (fm_styles_get_cell_style_xf_count(handle_, &n) != 0) {
      return 0U;
    }
    return n;
  }

  /// Returns the named cell style at `index`. Shape:
  /// `{ status, name, xfId, builtinId, iLevel, hidden, customBuiltin }`.
  /// `builtinId` is `0xFFFFFFFF` (`FM_CELL_STYLE_BUILTIN_ID_NONE`) for
  /// custom (non-built-in) entries.
  emscripten::val getCellStyle(uint32_t index) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    fm_cell_style_record_t cs{};
    fm_status_t rc = fm_styles_get_cell_style(handle_, index, &cs);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("name", std::string(cs.name != nullptr ? cs.name : ""));
    o.set("xfId", cs.xf_id);
    o.set("builtinId", cs.builtin_id);
    o.set("iLevel", cs.i_level);
    o.set("hidden", cs.hidden != 0);
    o.set("customBuiltin", cs.custom_builtin != 0);
    return o;
  }

  /// Returns the `<cellStyleXfs>` record at `index`. Shape mirrors
  /// `getCellXf`.
  emscripten::val getCellStyleXf(uint32_t index) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    fm_cell_xf xf{};
    fm_status_t rc = fm_styles_get_cell_style_xf(handle_, index, &xf);
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

  /// Returns the number of external-link records carried by the
  /// workbook. Always 0 for fresh workbooks.
  uint32_t externalLinkCount() const {
    if (handle_ == nullptr) {
      return 0U;
    }
    uint32_t n = 0;
    if (fm_workbook_external_link_count(handle_, &n) != 0) {
      return 0U;
    }
    return n;
  }

  /// Returns the `index`-th external-link record. The result object
  /// follows the same `{ status, ... }` envelope as the styles getters.
  emscripten::val getExternalLink(uint32_t index) const {
    emscripten::val o = emscripten::val::object();
    if (handle_ == nullptr) {
      o.set("status", error_status(7000));
      return o;
    }
    fm_external_link_record_t rec{};
    fm_status_t rc = fm_workbook_external_link_at(handle_, index, &rec);
    if (rc != 0) {
      o.set("status", error_status(rc));
      return o;
    }
    o.set("status", ok_status());
    o.set("index", rec.index);
    o.set("relId", std::string(rec.rel_id != nullptr ? rec.rel_id : ""));
    o.set("partPath", std::string(rec.part_path != nullptr ? rec.part_path : ""));
    o.set("target", std::string(rec.target != nullptr ? rec.target : ""));
    o.set("targetExternal", rec.target_external != 0);
    o.set("kind", rec.kind);
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

  /// `removeMerge(sheetIdx, {firstRow, lastRow, firstCol, lastCol})`.
  JsStatus removeMerge(uint32_t sheet, emscripten::val range) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_merge_range m;
    m.first_row = range["firstRow"].as<uint32_t>();
    m.last_row = range["lastRow"].as<uint32_t>();
    m.first_col = range["firstCol"].as<uint32_t>();
    m.last_col = range["lastCol"].as<uint32_t>();
    fm_status_t rc = fm_sheet_remove_merge(handle_, sheet, m);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `removeMergeAt(sheetIdx, index)`.
  JsStatus removeMergeAt(uint32_t sheet, uint32_t index) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_remove_merge_at(handle_, sheet, index);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `clearMerges(sheetIdx)`.
  JsStatus clearMerges(uint32_t sheet) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_clear_merges(handle_, sheet);
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

  /// `addHyperlink(sheetIdx, row, col, target, display, tooltip)`.
  ///
  /// Append a hyperlink entry to `sheet`. The frontend surface omits the
  /// `location` field; an empty string is forwarded to the C ABI so the
  /// writer mints a fresh `rId` on save and treats the location as
  /// absent. Pass empty strings for `display` / `tooltip` to mean "use
  /// the default" / "no tooltip".
  JsStatus addHyperlink(uint32_t sheet, uint32_t row, uint32_t col, const std::string& target,
                        const std::string& display, const std::string& tooltip) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_hyperlink hl{};
    hl.row = row;
    hl.col = col;
    hl.target = target.empty() ? nullptr : target.c_str();
    hl.location = nullptr;
    hl.display = display.empty() ? nullptr : display.c_str();
    hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
    fm_status_t rc = fm_sheet_add_hyperlink(handle_, sheet, hl);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `removeHyperlink(sheetIdx, row, col)`.
  JsStatus removeHyperlink(uint32_t sheet, uint32_t row, uint32_t col) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_remove_hyperlink(handle_, sheet, row, col);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `removeHyperlinkAt(sheetIdx, index)`.
  JsStatus removeHyperlinkAt(uint32_t sheet, uint32_t index) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_remove_hyperlink_at(handle_, sheet, index);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `clearHyperlinks(sheetIdx)`.
  JsStatus clearHyperlinks(uint32_t sheet) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_clear_hyperlinks(handle_, sheet);
    return rc == 0 ? ok_status() : error_status(rc);
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
      arr.set(i, merge_range_to_val(m));
    }
    return arr;
  }

  /// `getValidations(sheetIdx) -> Array<{ranges, type, op, errorStyle,
  ///                                      allowBlank, showInputMessage,
  ///                                      showErrorMessage, formula1,
  ///                                      formula2, errorTitle,
  ///                                      errorMessage, promptTitle,
  ///                                      promptMessage}>`.
  ///
  /// Each entry's `ranges` is an Array of `{firstRow, lastRow, firstCol,
  /// lastCol}`. Boolean fields are surfaced as JS booleans (not 0/1).
  emscripten::val getValidations(uint32_t sheet) const {
    emscripten::val arr = emscripten::val::array();
    if (handle_ == nullptr) {
      return arr;
    }
    uint32_t count = 0;
    if (fm_sheet_get_validation_count(handle_, sheet, &count) != 0) {
      return arr;
    }
    for (uint32_t i = 0; i < count; ++i) {
      fm_data_validation v{};
      if (fm_sheet_get_validation_at(handle_, sheet, i, &v) != 0) {
        continue;
      }
      emscripten::val item = emscripten::val::object();
      emscripten::val ranges = emscripten::val::array();
      for (uint32_t r = 0; r < v.range_count; ++r) {
        ranges.set(r, merge_range_to_val(v.ranges[r]));
      }
      item.set("ranges", ranges);
      item.set("type", v.type);
      item.set("op", v.op);
      item.set("errorStyle", v.error_style);
      item.set("allowBlank", v.allow_blank != 0);
      item.set("showInputMessage", v.show_input_message != 0);
      item.set("showErrorMessage", v.show_error_message != 0);
      item.set("formula1", v.formula1 != nullptr ? std::string(v.formula1) : std::string());
      item.set("formula2", v.formula2 != nullptr ? std::string(v.formula2) : std::string());
      item.set("errorTitle", v.error_title != nullptr ? std::string(v.error_title) : std::string());
      item.set("errorMessage", v.error_message != nullptr ? std::string(v.error_message) : std::string());
      item.set("promptTitle", v.prompt_title != nullptr ? std::string(v.prompt_title) : std::string());
      item.set("promptMessage", v.prompt_message != nullptr ? std::string(v.prompt_message) : std::string());
      arr.set(i, item);
    }
    return arr;
  }

  /// `addValidation(sheetIdx, {ranges, type, op?, errorStyle?, allowBlank?,
  ///                            showInputMessage?, showErrorMessage?,
  ///                            formula1?, formula2?, errorTitle?,
  ///                            errorMessage?, promptTitle?,
  ///                            promptMessage?})`.
  ///
  /// Optional fields default to `0` for the small enum-shaped integers,
  /// `false` for booleans, and `""` for strings. The string buffers
  /// (`formula1`, etc.) are kept alive on the C++ stack frame for the
  /// duration of the C ABI call so the borrowed `const char*` pointers
  /// stay valid; the C ABI deep-copies them into the workbook's storage.
  JsStatus addValidation(uint32_t sheet, emscripten::val v) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    // Pull every JS field into local storage first; the C ABI receives
    // borrowed `const char*` views that must stay valid until
    // `fm_sheet_add_validation` returns.
    std::vector<fm_merge_range> ranges_buf;
    if (v.hasOwnProperty("ranges")) {
      emscripten::val ranges_js = v["ranges"];
      if (ranges_js.isArray()) {
        const uint32_t n = ranges_js["length"].as<uint32_t>();
        ranges_buf.reserve(n);
        for (uint32_t i = 0; i < n; ++i) {
          emscripten::val rng = ranges_js[i];
          fm_merge_range m{};
          m.first_row = rng["firstRow"].as<uint32_t>();
          m.last_row = rng["lastRow"].as<uint32_t>();
          m.first_col = rng["firstCol"].as<uint32_t>();
          m.last_col = rng["lastCol"].as<uint32_t>();
          ranges_buf.push_back(m);
        }
      }
    }
    const std::string formula1 = js_pull_string(v, "formula1");
    const std::string formula2 = js_pull_string(v, "formula2");
    const std::string error_title = js_pull_string(v, "errorTitle");
    const std::string error_message = js_pull_string(v, "errorMessage");
    const std::string prompt_title = js_pull_string(v, "promptTitle");
    const std::string prompt_message = js_pull_string(v, "promptMessage");

    fm_data_validation dv{};
    dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
    dv.range_count = static_cast<uint32_t>(ranges_buf.size());
    dv.type = js_pull_u8(v, "type", 0U);
    dv.op = js_pull_u8(v, "op", 0U);
    dv.error_style = js_pull_u8(v, "errorStyle", 0U);
    dv.allow_blank = js_pull_bool(v, "allowBlank", true) ? 1 : 0;
    dv.show_input_message = js_pull_bool(v, "showInputMessage", false) ? 1 : 0;
    dv.show_error_message = js_pull_bool(v, "showErrorMessage", false) ? 1 : 0;
    dv.formula1 = formula1.empty() ? nullptr : formula1.c_str();
    dv.formula2 = formula2.empty() ? nullptr : formula2.c_str();
    dv.error_title = error_title.empty() ? nullptr : error_title.c_str();
    dv.error_message = error_message.empty() ? nullptr : error_message.c_str();
    dv.prompt_title = prompt_title.empty() ? nullptr : prompt_title.c_str();
    dv.prompt_message = prompt_message.empty() ? nullptr : prompt_message.c_str();
    fm_status_t rc = fm_sheet_add_validation(handle_, sheet, dv);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `removeValidationAt(sheetIdx, index)`.
  JsStatus removeValidationAt(uint32_t sheet, uint32_t index) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_remove_validation_at(handle_, sheet, index);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `clearValidations(sheetIdx)`.
  JsStatus clearValidations(uint32_t sheet) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_clear_validations(handle_, sheet);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `getConditionalFormats(sheetIdx) -> Array<{id, type, priority, stopIfTrue,
  ///   sqref:[{firstRow,firstCol,lastRow,lastCol}], dxfId?, formula1?, formula2?,
  ///   op?, rank?, percent?, bottom?, aboveAverage?, equalAverage?, stdDev?,
  ///   text?, timePeriod?}>`.
  ///
  /// Visual rule kinds (`ColorScale`=2, `DataBar`=3, `IconSet`=4) appear
  /// with `type` populated but their visual sub-spec fields omitted: the
  /// XML round-trip preserves them verbatim, but the embind read surface
  /// surfaces only the fields it can also accept on `addConditionalFormat`.
  emscripten::val getConditionalFormats(uint32_t sheet) const {
    emscripten::val arr = emscripten::val::array();
    if (handle_ == nullptr) {
      return arr;
    }
    std::size_t count = 0;
    if (fm_sheet_cf_count(handle_, sheet, &count) != 0) {
      return arr;
    }
    for (std::size_t i = 0; i < count; ++i) {
      fm_cf_rule_t rule{};
      if (fm_sheet_cf_get_at(handle_, sheet, i, &rule) != 0) {
        continue;
      }
      emscripten::val item = emscripten::val::object();
      item.set("id", rule.id != nullptr ? std::string(rule.id) : std::string());
      item.set("type", static_cast<uint32_t>(rule.type));
      item.set("priority", rule.priority);
      item.set("stopIfTrue", rule.stop_if_true != 0);
      emscripten::val sqref = emscripten::val::array();
      for (uint32_t r = 0; r < rule.sqref_count; ++r) {
        emscripten::val rng = emscripten::val::object();
        rng.set("firstRow", rule.sqref[r].first_row);
        rng.set("firstCol", rule.sqref[r].first_col);
        rng.set("lastRow", rule.sqref[r].last_row);
        rng.set("lastCol", rule.sqref[r].last_col);
        sqref.set(r, rng);
      }
      item.set("sqref", sqref);
      if (rule.dxf_id_engaged != 0) {
        item.set("dxfId", rule.dxf_id);
      }
      if (rule.formula1 != nullptr) {
        item.set("formula1", std::string(rule.formula1));
      }
      if (rule.formula2 != nullptr) {
        item.set("formula2", std::string(rule.formula2));
      }
      if (rule.op_engaged != 0) {
        item.set("op", static_cast<uint32_t>(rule.op));
      }
      if (rule.rank_engaged != 0) {
        item.set("rank", rule.rank);
        item.set("percent", rule.percent != 0);
        item.set("bottom", rule.bottom != 0);
      }
      // aboveAverage flags are always present (default-engineered), but
      // we only surface them for the AboveAverage rule type to avoid
      // confusing FE consumers.
      if (rule.type == 6 /* AboveAverage */) {
        item.set("aboveAverage", rule.above_average != 0);
        item.set("equalAverage", rule.equal_average != 0);
        if (rule.std_dev_engaged != 0) {
          item.set("stdDev", rule.std_dev);
        }
      }
      if (rule.text != nullptr) {
        item.set("text", std::string(rule.text));
      }
      if (rule.time_period_engaged != 0) {
        item.set("timePeriod", static_cast<uint32_t>(rule.time_period));
      }
      arr.set(static_cast<uint32_t>(i), item);
    }
    return arr;
  }

  /// `addConditionalFormat(sheetIdx, {sqref:[{firstRow,firstCol,lastRow,
  ///   lastCol}], type, priority?, stopIfTrue?, id?, dxfId?, formula1?,
  ///   formula2?, op?, rank?, percent?, bottom?, aboveAverage?,
  ///   equalAverage?, stdDev?, text?, timePeriod?})`.
  ///
  /// Visual rule kinds (`ColorScale`=2, `DataBar`=3, `IconSet`=4) are
  /// rejected — those payloads still round-trip through OOXML reader /
  /// writer but cannot be created from this API yet.
  JsStatus addConditionalFormat(uint32_t sheet, emscripten::val v) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    // Pull every JS field into local storage; the C ABI receives
    // borrowed `const char*` views that must stay valid for the
    // duration of the call.
    std::vector<fm_cf_cell_range_t> ranges_buf;
    if (v.hasOwnProperty("sqref")) {
      emscripten::val sqref_js = v["sqref"];
      if (sqref_js.isArray()) {
        const uint32_t n = sqref_js["length"].as<uint32_t>();
        ranges_buf.reserve(n);
        for (uint32_t i = 0; i < n; ++i) {
          emscripten::val rng = sqref_js[i];
          fm_cf_cell_range_t r{};
          r.first_row = rng["firstRow"].as<uint32_t>();
          r.first_col = rng["firstCol"].as<uint32_t>();
          r.last_row = rng["lastRow"].as<uint32_t>();
          r.last_col = rng["lastCol"].as<uint32_t>();
          ranges_buf.push_back(r);
        }
      }
    }
    const std::string id = js_pull_string(v, "id");
    const std::string formula1 = js_pull_string(v, "formula1");
    const std::string formula2 = js_pull_string(v, "formula2");
    const std::string text = js_pull_string(v, "text");

    fm_cf_rule_t rule{};
    rule.id = id.empty() ? nullptr : id.c_str();
    rule.type = js_pull_u8(v, "type", 0U);
    rule.priority = js_pull_u32(v, "priority", 0U) != 0 ? static_cast<int32_t>(js_pull_u32(v, "priority", 0U)) : 0;
    rule.stop_if_true = js_pull_bool(v, "stopIfTrue", false) ? 1 : 0;
    if (!v["dxfId"].isUndefined() && !v["dxfId"].isNull()) {
      rule.dxf_id_engaged = 1;
      rule.dxf_id = v["dxfId"].as<uint32_t>();
    }
    rule.sqref = ranges_buf.empty() ? nullptr : ranges_buf.data();
    rule.sqref_count = static_cast<uint32_t>(ranges_buf.size());
    rule.formula1 = formula1.empty() ? nullptr : formula1.c_str();
    rule.formula2 = formula2.empty() ? nullptr : formula2.c_str();
    if (!v["op"].isUndefined() && !v["op"].isNull()) {
      rule.op_engaged = 1;
      rule.op = js_pull_u8(v, "op", 0U);
    }
    if (!v["rank"].isUndefined() && !v["rank"].isNull()) {
      rule.rank_engaged = 1;
      rule.rank = static_cast<int32_t>(v["rank"].as<int32_t>());
    }
    rule.percent = js_pull_bool(v, "percent", false) ? 1 : 0;
    rule.bottom = js_pull_bool(v, "bottom", false) ? 1 : 0;
    rule.above_average = js_pull_bool(v, "aboveAverage", true) ? 1 : 0;
    rule.equal_average = js_pull_bool(v, "equalAverage", false) ? 1 : 0;
    if (!v["stdDev"].isUndefined() && !v["stdDev"].isNull()) {
      rule.std_dev_engaged = 1;
      rule.std_dev = v["stdDev"].as<double>();
    }
    rule.text = text.empty() ? nullptr : text.c_str();
    if (!v["timePeriod"].isUndefined() && !v["timePeriod"].isNull()) {
      rule.time_period_engaged = 1;
      rule.time_period = js_pull_u8(v, "timePeriod", 0U);
    }
    fm_status_t rc = fm_sheet_cf_add_rule(handle_, sheet, rule);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `removeConditionalFormatAt(sheetIdx, index)`.
  JsStatus removeConditionalFormatAt(uint32_t sheet, uint32_t index) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_cf_remove_at(handle_, sheet, index);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `clearConditionalFormats(sheetIdx)`.
  JsStatus clearConditionalFormats(uint32_t sheet) {
    if (handle_ == nullptr) {
      return error_status(7000);
    }
    fm_status_t rc = fm_sheet_cf_clear(handle_, sheet);
    return rc == 0 ? ok_status() : error_status(rc);
  }

  /// `precedents(sheet, row, col, depth?) -> Array<{sheet, row, col}>`.
  /// `depth` defaults to 1 (direct precedents); larger values BFS up to
  /// the engine cap (32).
  emscripten::val precedents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
    return trace_to_val(fm_workbook_precedents, sheet, row, col, depth);
  }

  /// `dependents(sheet, row, col, depth?) -> Array<{sheet, row, col}>`.
  emscripten::val dependents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
    return trace_to_val(fm_workbook_dependents, sheet, row, col, depth);
  }

  /// `functionMetadata(name, locale?) -> {name, minArity, maxArity,
  ///                                        signatureTemplate?, description?}`.
  /// `locale` is `0` for `en-US` (default) or `1` for `ja-JP`. The
  /// signature / description fields are absent until the locale
  /// metadata table is populated; consumers must default-handle.
  emscripten::val functionMetadata(const std::string& name, uint32_t locale) const {
    emscripten::val o = emscripten::val::object();
    fm_function_metadata_t md{};
    fm_status_t rc = fm_function_metadata(name.c_str(), static_cast<fm_locale_t>(locale), &md);
    if (rc != 0) {
      o.set("ok", false);
      return o;
    }
    o.set("ok", true);
    o.set("name", md.canonical_name != nullptr ? std::string(md.canonical_name) : std::string());
    o.set("minArity", md.min_arity);
    o.set("maxArity", md.max_arity);
    if (md.signature_template != nullptr) {
      o.set("signatureTemplate", std::string(md.signature_template));
    }
    if (md.description != nullptr) {
      o.set("description", std::string(md.description));
    }
    return o;
  }

  /// `functionNames() -> string[]` — every Formulon function's
  /// canonical name in ascending sort order.
  emscripten::val functionNames() const {
    emscripten::val arr = emscripten::val::array();
    const std::size_t n = fm_function_count();
    for (std::size_t i = 0; i < n; ++i) {
      const char* name = nullptr;
      if (fm_function_name_at(i, &name) != 0 || name == nullptr) {
        continue;
      }
      arr.set(static_cast<uint32_t>(i), std::string(name));
    }
    return arr;
  }

  /// `localizeFunctionName(canonicalName, locale) -> string`. Returns
  /// the localized display name; falls through to the canonical name
  /// when the locale's alias table is empty (currently always for
  /// non-`en-US` locales). Returns the empty string when the canonical
  /// name does not match any registered function.
  std::string localizeFunctionName(const std::string& canonical_name, uint32_t locale) const {
    const char* out = nullptr;
    if (fm_function_localize(canonical_name.c_str(), static_cast<fm_locale_t>(locale), &out) != 0 || out == nullptr) {
      return std::string();
    }
    return std::string(out);
  }

  /// `canonicalizeFunctionName(localizedName, locale) -> string`.
  std::string canonicalizeFunctionName(const std::string& localized_name, uint32_t locale) const {
    const char* out = nullptr;
    if (fm_function_canonicalize(localized_name.c_str(), static_cast<fm_locale_t>(locale), &out) != 0 ||
        out == nullptr) {
      return std::string();
    }
    return std::string(out);
  }

  /// `spillInfo(sheet, row, col) -> {engaged, anchorRow, anchorCol, rows, cols}`.
  /// `engaged` is `false` when the cell is not part of any spill region;
  /// the other fields are zero in that case.
  emscripten::val spillInfo(uint32_t sheet, uint32_t row, uint32_t col) const {
    emscripten::val item = emscripten::val::object();
    item.set("engaged", false);
    item.set("anchorRow", 0U);
    item.set("anchorCol", 0U);
    item.set("rows", 0U);
    item.set("cols", 0U);
    if (handle_ == nullptr) {
      return item;
    }
    fm_spill_info_t info{};
    if (fm_workbook_spill_info(handle_, sheet, row, col, &info) != 0) {
      return item;
    }
    item.set("engaged", info.engaged != 0);
    item.set("anchorRow", info.anchor_row);
    item.set("anchorCol", info.anchor_col);
    item.set("rows", info.rows);
    item.set("cols", info.cols);
    return item;
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
  // Shared bridge for `precedents` / `dependents`: invokes the C ABI
  // entry point, copies the result into a JS array of {sheet, row, col}
  // value-objects, and frees the C-owned handle. Returns an empty array
  // on any error so the JS side does not need a separate failure path.
  using TraceFn = fm_status_t (*)(const fm_workbook_t*, uint32_t, uint32_t, uint32_t, uint32_t, fm_cell_nodes_t**);
  emscripten::val trace_to_val(TraceFn fn, uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
    emscripten::val arr = emscripten::val::array();
    if (handle_ == nullptr) {
      return arr;
    }
    fm_cell_nodes_t* nodes = nullptr;
    if (fn(handle_, sheet, row, col, depth, &nodes) != 0) {
      return arr;
    }
    const std::size_t count = fm_cell_nodes_count(nodes);
    for (std::size_t i = 0; i < count; ++i) {
      fm_cell_node_t n{};
      if (fm_cell_nodes_at(nodes, i, &n) != 0) {
        continue;
      }
      emscripten::val item = emscripten::val::object();
      item.set("sheet", n.sheet);
      item.set("row", n.row);
      item.set("col", n.col);
      arr.set(static_cast<uint32_t>(i), item);
    }
    fm_cell_nodes_destroy(nodes);
    return arr;
  }

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

  // @size-budget: 14 KB
  // Covers the JsSheetProtection value_object, JsSheetProtectionResult,
  // and the two bridge methods (getSheetProtection / setSheetProtection).
  value_object<JsSheetProtection>("SheetProtection")
      .field("enabled", &JsSheetProtection::enabled)
      .field("algorithmName", &JsSheetProtection::algorithmName)
      .field("hashValue", &JsSheetProtection::hashValue)
      .field("saltValue", &JsSheetProtection::saltValue)
      .field("spinCount", &JsSheetProtection::spinCount)
      .field("legacyPassword", &JsSheetProtection::legacyPassword)
      .field("sheet", &JsSheetProtection::sheet)
      .field("objects", &JsSheetProtection::objects)
      .field("scenarios", &JsSheetProtection::scenarios)
      .field("formatCells", &JsSheetProtection::formatCells)
      .field("formatColumns", &JsSheetProtection::formatColumns)
      .field("formatRows", &JsSheetProtection::formatRows)
      .field("insertColumns", &JsSheetProtection::insertColumns)
      .field("insertRows", &JsSheetProtection::insertRows)
      .field("insertHyperlinks", &JsSheetProtection::insertHyperlinks)
      .field("deleteColumns", &JsSheetProtection::deleteColumns)
      .field("deleteRows", &JsSheetProtection::deleteRows)
      .field("selectLockedCells", &JsSheetProtection::selectLockedCells)
      .field("selectUnlockedCells", &JsSheetProtection::selectUnlockedCells)
      .field("sort", &JsSheetProtection::sort)
      .field("autoFilter", &JsSheetProtection::autoFilter)
      .field("pivotTables", &JsSheetProtection::pivotTables);

  value_object<JsSheetProtectionResult>("SheetProtectionResult")
      .field("status", &JsSheetProtectionResult::status)
      .field("protection", &JsSheetProtectionResult::protection);

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

  value_object<JsAddStyleResult>("AddStyleResult")
      .field("status", &JsAddStyleResult::status)
      .field("index", &JsAddStyleResult::index);

  value_object<JsAddNumFmtResult>("AddNumFmtResult")
      .field("status", &JsAddNumFmtResult::status)
      .field("numFmtId", &JsAddNumFmtResult::numFmtId);

  // ---- Workbook class ------------------------------------------------------
  class_<JsWorkbook>("Workbook")
      .class_function("createDefault", &JsWorkbook::createDefault, allow_raw_pointers())
      .class_function("createEmpty", &JsWorkbook::createEmpty, allow_raw_pointers())
      .class_function("loadBytes", &JsWorkbook::loadBytes, allow_raw_pointers())
      .function("addBorder", &JsWorkbook::addBorder)
      .function("addConditionalFormat", &JsWorkbook::addConditionalFormat)
      .function("addFill", &JsWorkbook::addFill)
      .function("addFont", &JsWorkbook::addFont)
      .function("addHyperlink", &JsWorkbook::addHyperlink)
      .function("addMerge", &JsWorkbook::addMerge)
      .function("addNumFmt", &JsWorkbook::addNumFmt)
      .function("addSheet", &JsWorkbook::addSheet)
      .function("addValidation", &JsWorkbook::addValidation)
      .function("addXf", &JsWorkbook::addXf)
      .function("borderCount", &JsWorkbook::borderCount)
      .function("calcMode", &JsWorkbook::calcMode)
      .function("canonicalizeFunctionName", &JsWorkbook::canonicalizeFunctionName)
      .function("cellAt", &JsWorkbook::cellAt)
      .function("cellCount", &JsWorkbook::cellCount)
      .function("cellStyleCount", &JsWorkbook::cellStyleCount)
      .function("cellStyleXfCount", &JsWorkbook::cellStyleXfCount)
      .function("clearConditionalFormats", &JsWorkbook::clearConditionalFormats)
      .function("clearHyperlinks", &JsWorkbook::clearHyperlinks)
      .function("clearMerges", &JsWorkbook::clearMerges)
      .function("clearValidations", &JsWorkbook::clearValidations)
      .function("definedNameAt", &JsWorkbook::definedNameAt)
      .function("definedNameCount", &JsWorkbook::definedNameCount)
      .function("deleteCols", &JsWorkbook::deleteCols)
      .function("deleteRows", &JsWorkbook::deleteRows)
      .function("dependents", &JsWorkbook::dependents)
      .function("evaluateCfRange", &JsWorkbook::evaluateCfRange)
      .function("externalLinkCount", &JsWorkbook::externalLinkCount)
      .function("fillCount", &JsWorkbook::fillCount)
      .function("fontCount", &JsWorkbook::fontCount)
      .function("functionMetadata", &JsWorkbook::functionMetadata)
      .function("functionNames", &JsWorkbook::functionNames)
      .function("getBorder", &JsWorkbook::getBorder)
      .function("getCellStyle", &JsWorkbook::getCellStyle)
      .function("getCellStyleXf", &JsWorkbook::getCellStyleXf)
      .function("getCellXf", &JsWorkbook::getCellXf)
      .function("getCellXfIndex", &JsWorkbook::getCellXfIndex)
      .function("getComment", &JsWorkbook::getComment)
      .function("getConditionalFormats", &JsWorkbook::getConditionalFormats)
      .function("getExternalLink", &JsWorkbook::getExternalLink)
      .function("getFill", &JsWorkbook::getFill)
      .function("getFont", &JsWorkbook::getFont)
      .function("getHyperlinks", &JsWorkbook::getHyperlinks)
      .function("getLambdaText", &JsWorkbook::getLambdaText)
      .function("getMerges", &JsWorkbook::getMerges)
      .function("getNumFmt", &JsWorkbook::getNumFmt)
      .function("getSheetColumns", &JsWorkbook::getSheetColumns)
      .function("getSheetProtection", &JsWorkbook::getSheetProtection)
      .function("getSheetRowOverrides", &JsWorkbook::getSheetRowOverrides)
      .function("getSheetView", &JsWorkbook::getSheetView)
      .function("getValidations", &JsWorkbook::getValidations)
      .function("getValue", &JsWorkbook::getValue)
      .function("insertCols", &JsWorkbook::insertCols)
      .function("insertRows", &JsWorkbook::insertRows)
      .function("isValid", &JsWorkbook::isValid)
      .function("localizeFunctionName", &JsWorkbook::localizeFunctionName)
      .function("moveSheet", &JsWorkbook::moveSheet)
      .function("partialRecalc", &JsWorkbook::partialRecalc)
      .function("passthroughAt", &JsWorkbook::passthroughAt)
      .function("passthroughCount", &JsWorkbook::passthroughCount)
      .function("precedents", &JsWorkbook::precedents)
      .function("recalc", &JsWorkbook::recalc)
      .function("removeConditionalFormatAt", &JsWorkbook::removeConditionalFormatAt)
      .function("removeHyperlink", &JsWorkbook::removeHyperlink)
      .function("removeHyperlinkAt", &JsWorkbook::removeHyperlinkAt)
      .function("removeMerge", &JsWorkbook::removeMerge)
      .function("removeMergeAt", &JsWorkbook::removeMergeAt)
      .function("removeSheet", &JsWorkbook::removeSheet)
      .function("removeValidationAt", &JsWorkbook::removeValidationAt)
      .function("renameSheet", &JsWorkbook::renameSheet)
      .function("save", &JsWorkbook::save)
      .function("setBlank", &JsWorkbook::setBlank)
      .function("setBool", &JsWorkbook::setBool)
      .function("setCalcMode", &JsWorkbook::setCalcMode)
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
      .function("setSheetProtection", &JsWorkbook::setSheetProtection)
      .function("setSheetTabHidden", &JsWorkbook::setSheetTabHidden)
      .function("setSheetZoom", &JsWorkbook::setSheetZoom)
      .function("setText", &JsWorkbook::setText)
      .function("sheetCount", &JsWorkbook::sheetCount)
      .function("sheetName", &JsWorkbook::sheetName)
      .function("spillInfo", &JsWorkbook::spillInfo)
      .function("tableAt", &JsWorkbook::tableAt)
      .function("tableCount", &JsWorkbook::tableCount)
      .function("xfCount", &JsWorkbook::xfCount);

  // ---- Free functions ------------------------------------------------------
  function("evalFormula", &eval_formula);
  function("versionString", &version_string);
  function("statusString", &status_string);
  function("lastErrorMessage", &last_error_message);
  function("lastErrorContext", &last_error_context);
}
