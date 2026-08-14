//
// Shared types and helpers for the Emscripten binding TUs under
// `src/wasm/parts/`. The whole bundle is compiled only when
// `FM_BUILD_WASM=ON` and exists purely to keep the per-area binding
// surface manageable -- each `parts/*.cpp` declares one slice of the
// `JsWorkbook` class plus its own `EMSCRIPTEN_BINDINGS` block or
// dispatches into the central registrar in `bindings_register.cpp`.
//
// What lives here:
//   * Value-object structs (`JsStatus`, `JsValue`, `JsCellResult`, ...)
//     that the JS-facing API surface returns.
//   * Translation helpers (`translate_value`, `translate_cf_*`,
//     `merge_range_to_val`, `bytes_to_val`, `val_to_bytes`).
//   * Status builders (`ok_status`, `error_status`, `status_from_rc`).
//   * Small `js_pull_*` field-extractor helpers used by the value-object
//     adders.
//
// The translation helpers are deliberately not inline: they sit on the
// cold side of the binding surface (one call per failure / read), and
// keeping a single emission of them in `embind_common.cpp` lets the
// linker dedupe across the ~9 part TUs. The `js_pull_*` helpers are
// inline because they fold into the call sites that pull
// `emscripten::val` fields apart, and inlining yields a measurably
// smaller `.wasm.br` under `-O3 + wasm-opt -Oz`.

#ifndef FORMULON_WASM_PARTS_EMBIND_COMMON_H_
#define FORMULON_WASM_PARTS_EMBIND_COMMON_H_

#include <emscripten/bind.h>
#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"

namespace formulon {
namespace wasm {
namespace parts {

/// A JS wrapper was used after its workbook handle was destroyed. This is
/// the binding-layer `kBindingInvalidHandle` code, not the C ABI's 7001
/// `kBindingNullPointer` argument-validation error.
constexpr int32_t kBindingInvalidHandle = 7000;

// ---- Value-object mirrors -----------------------------------------------
//
// Every cross-boundary record is reflected through a POD struct so embind
// can serialise it via `value_object<>`. The fields use signed `int32_t`
// for boolean-shaped enums because embind tends to surface those more
// cleanly than `uint8_t`; the few unsigned fields are kept `uint32_t` to
// match the upstream `fm_*` ABI shape.

/// Mirror of `fm_value_t`. embind cannot project a C union, so the
/// variant payload is flattened across `number / boolean / text /
/// errorCode` and JS dispatches on `kind`.
struct JsValue {
  int32_t kind = 0;       ///< `fm_value_kind_t` ordinal (0..7).
  double number = 0.0;    ///< Active when kind == FM_VAL_NUMBER.
  int32_t boolean = 0;    ///< Active when kind == FM_VAL_BOOL (0/1).
  std::string text;       ///< Active when kind == FM_VAL_TEXT.
  int32_t errorCode = 0;  ///< Active when kind == FM_VAL_ERROR.
};

/// `{ ok, status, message, context }` envelope returned by every
/// fallible binding entry point. Replaces throwing across the JS
/// boundary, which we cannot do under `-fno-exceptions`.
struct JsStatus {
  bool ok = true;
  int32_t status = 0;
  std::string message;
  std::string context;
};

/// Per-call counters returned by `Workbook.recalcParallel`. The five
/// scheduler counters originate as `uint64_t` in the C ABI but are widened
/// to `double` here deliberately: embind then exposes them as JavaScript
/// `number` values (exact for every value through `Number.MAX_SAFE_INTEGER`)
/// instead of depending on WASM BigInt support.
struct JsParallelRecalcStats {
  double cellsEvaluated = 0.0;
  double sccsProcessed = 0.0;
  double parallelSteps = 0.0;
  double serialFallbackSteps = 0.0;
  double cycleRecoveries = 0.0;
  uint32_t workerThreadsStarted = 0U;
  uint32_t workerThreadsUsed = 0U;
};

/// `{ status, stats }` envelope returned by `Workbook.recalcParallel`.
struct JsParallelRecalcResult {
  JsStatus status;
  JsParallelRecalcStats stats;
};

/// Bundles `JsStatus` and a `JsValue` payload for `evalFormula`.
struct JsEvalResult {
  JsStatus status;
  JsValue value;
};

/// Result envelope for `getValue`-style calls.
struct JsCellResult {
  JsStatus status;
  JsValue value;
};

/// Result envelope for `save()` returning the bytes as a JS Uint8Array.
struct JsSaveResult {
  JsStatus status;
  emscripten::val bytes = emscripten::val::null();
};

/// Result envelope for an explicit-format save with the loss counters the
/// write produced. Counters a given container cannot produce stay zero.
struct JsSaveDiagnosticsResult {
  JsStatus status;
  emscripten::val bytes = emscripten::val::null();
  uint32_t downgradedFormulaCount = 0;
  uint32_t deferredFeatureCount = 0;
  uint32_t droppedPartCount = 0;
  uint32_t droppedRelationshipCount = 0;
  uint32_t renumberedPartCount = 0;
};

/// Recovery / passthrough counters captured while loading a workbook, for
/// either container format.
struct JsReadDiagnosticsResult {
  JsStatus status;
  uint32_t undecodedFormulaCount = 0;
  uint32_t undecodedDefinedNameCount = 0;
  uint32_t undecodedPartCount = 0;
  uint32_t skippedFeatureCount = 0;
  uint32_t unknownContentTypeCount = 0;
};

/// Result envelope for the string-payload accessors (`sheetName`,
/// `pivotCacheFieldName`, ...).
struct JsStringResult {
  JsStatus status;
  std::string value;
};

/// JS-side mirror of `fm_cf_color_t`. Channels are 0-255 (sRGB); widened
/// to signed int32 because embind serialises that more cleanly.
struct JsCfColor {
  int32_t r = 0;
  int32_t g = 0;
  int32_t b = 0;
  int32_t a = 0;
};

/// JS-side mirror of `fm_cf_match_t`. The active sub-fields depend on
/// `kind`; the rest carry default-zero values.
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

/// JS-side mirror of `fm_sheet_view_t`.
struct JsSheetView {
  uint32_t zoomScale = 100U;
  uint32_t freezeRows = 0U;
  uint32_t freezeCols = 0U;
  int32_t tabHidden = 0;
  int32_t showGridLines = 1;
  int32_t showRowColHeaders = 1;
  int32_t showZeros = 1;
  int32_t rightToLeft = 0;
  int32_t tabSelected = 0;
  std::string viewMode;
};

/// Return envelope for `Workbook.getSheetView(...)`.
struct JsSheetViewResult {
  JsStatus status;
  JsSheetView view{};
};

/// JS-side mirror of `fm_sheet_protection_t`. The deeply-copied string
/// fields decouple the returned record from the workbook's storage.
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

/// Return envelope for `Workbook.addFont` / `addFill` / `addBorder` /
/// `addXf` and most "add and return the index" surfaces.
struct JsAddStyleResult {
  JsStatus status;
  uint32_t index = 0U;
};

/// Return envelope for `Workbook.addNumFmt`. `numFmtId` may reuse any
/// existing effective mapping (built-in or custom, including a custom record
/// overriding a built-in slot), or be the freshly-assigned custom id
/// (`>= 164`). Effective mappings use the first valid custom record for an id
/// in document order, then a non-empty built-in code.
struct JsAddNumFmtResult {
  JsStatus status;
  uint32_t numFmtId = 0U;
};

// ---- Status builders ----------------------------------------------------
//
// Out-of-line so the linker emits exactly one copy across all part TUs;
// they capture the thread-local diagnostic surface that follows every
// C-ABI call.

/// Builds an `ok` envelope with empty diagnostic strings.
JsStatus ok_status();

/// Builds a failure envelope from `code`, copying out the thread-local
/// diagnostics surfaced by the most recent C-ABI call.
JsStatus error_status(int32_t code);

/// Bridges a `fm_status_t` into a `JsStatus` envelope.
JsStatus status_from_rc(fm_status_t rc);

// ---- Translation helpers -----------------------------------------------

/// Translates a `fm_value_t` into the embind-friendly `JsValue`. The
/// text variant is deep-copied; the other variants project the scalar
/// payload directly.
JsValue translate_value(const fm_value_t& v);

/// Translates an `fm_cf_color_t` into the embind value-object mirror.
JsCfColor translate_cf_color(const fm_cf_color_t& c);

/// Translates an `fm_cf_match_t` POD into the embind value-object
/// mirror.
JsCfMatch translate_cf_match(const fm_cf_match_t& m);

/// Builds a `{firstRow, lastRow, firstCol, lastCol}` JS object from a
/// merge / validation range record.
emscripten::val merge_range_to_val(const fm_merge_range& m);

/// Builds the empty `{status, top:0, left:0, rows:0, cols:0, cells:[]}`
/// envelope shape used by `pivotLayout()` on failure paths.
emscripten::val empty_pivot_layout_result(JsStatus status);

/// Translates one `fm_pivot_cell_t` into a JS object suitable for
/// inclusion in the pivotLayout `cells` array.
emscripten::val pivot_cell_to_val(const fm_pivot_cell_t& cell);

/// Copies the contents of a JS Uint8Array (passed via `emscripten::val`)
/// into a `std::vector<uint8_t>`. Returns an empty vector when the
/// argument is not a typed array.
std::vector<uint8_t> val_to_bytes(const emscripten::val& v);

/// Materialises a fresh JS `Uint8Array` from a contiguous byte range.
/// Element-by-element copy via `set(i, val)` for layout independence.
emscripten::val bytes_to_val(const uint8_t* data, std::size_t len);

// ---- JS field-extraction helpers ---------------------------------------
//
// These replace the repetitive
// `record["x"].isUndefined() ? dflt : record["x"].as<T>()` pattern that
// previously appeared inside `addFont` / `addFill` / `addBorder` /
// `addXf` / `addValidation`. Centralising them lets the compiler emit
// one copy of the embind glue (`val::operator[]`, `val::isUndefined`,
// `val::as<T>`) per field type instead of per call site, which is a
// measurable WASM size win because every embind operation pulls in
// non-trivial JS-bridge stubs.
//
// They are deliberately `inline`: empirically `-Oz` keeps the inlined
// forms smaller after `wasm-opt --converge` than the call-shaped forms.

/// Returns the `uint32_t` value of `v[key]`, or `dflt` when the field
/// is missing / undefined / null.
inline uint32_t js_pull_u32(const emscripten::val& v, const char* key, uint32_t dflt) {
  emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return dflt;
  }
  return f.as<uint32_t>();
}

/// Returns the low-byte `uint8_t` value of `v[key]`, or `dflt` when
/// missing.
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

/// Returns the string value of `v[key]`, or an empty string when
/// missing.
inline std::string js_pull_string(const emscripten::val& v, const char* key) {
  emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return std::string();
  }
  return f.as<std::string>();
}

/// Pulls the `{kind, rgb, theme, tint, indexed}` colour specification out
/// of `v[key]`. An absent object leaves `kind` at `kFmColorNone`, which
/// makes the writer emit the sibling `*Argb` as literal `rgb`. A supplied
/// selector is authoritative; the binding does not resolve theme/indexed /
/// auto colours.
inline fm_color_spec js_pull_color_spec(const emscripten::val& v, const char* key) {
  fm_color_spec spec{};
  emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return spec;
  }
  spec.kind = js_pull_u8(f, "kind", 0);
  spec.rgb = js_pull_u32(f, "rgb", 0U);
  spec.theme = js_pull_u32(f, "theme", 0U);
  spec.tint = js_pull_double(f, "tint", 0.0);
  spec.indexed = js_pull_u32(f, "indexed", 0U);
  return spec;
}

/// Builds the JS mirror of a colour specification. Reads emit it
/// unconditionally so a get / edit / add cycle passes the theme or
/// indexed colour straight back and keeps hitting style-table dedup.
inline emscripten::val js_color_spec(const fm_color_spec& spec) {
  emscripten::val o = emscripten::val::object();
  o.set("kind", static_cast<uint32_t>(spec.kind));
  o.set("rgb", spec.rgb);
  o.set("theme", spec.theme);
  o.set("tint", spec.tint);
  o.set("indexed", spec.indexed);
  return o;
}

/// Pulls a `{style, colorArgb, color}` border-side record out of `v`,
/// defaulting every absent field to zero.
inline fm_border_side js_pull_border_side(const emscripten::val& v) {
  fm_border_side s{};
  if (v.isUndefined() || v.isNull()) {
    return s;
  }
  s.style = js_pull_u8(v, "style", 0);
  s.color_argb = js_pull_u32(v, "colorArgb", 0U);
  s.color = js_pull_color_spec(v, "color");
  return s;
}

/// Builds the JS mirror of a border side.
inline emscripten::val js_border_side(const fm_border_side& s) {
  emscripten::val o = emscripten::val::object();
  o.set("style", static_cast<uint32_t>(s.style));
  o.set("colorArgb", s.color_argb);
  o.set("color", js_color_spec(s.color));
  return o;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon

#endif  // FORMULON_WASM_PARTS_EMBIND_COMMON_H_
