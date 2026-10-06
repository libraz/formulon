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
//     `merge_range_to_val`, `bytes_to_val`, `val_to_bytes_checked`).
//   * Status builders (`ok_status`, `error_status`, `status_from_rc`).
//   * checked nested-field readers and `js_set_*` record builders.
//
// The helpers are deliberately out of line: they sit on the cold side of
// the binding surface, and a single emission in `embind_common.cpp` lets
// the ~9 part TUs share one copy of each embind bridge sequence.

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

/// A JS callable handed to the binding threw. Reported on the envelope of
/// the call that ran it; see `call_js_callback` below.
constexpr int32_t kBindingCallbackException = 7003;

/// `formulon::FormulonErrorCode::kInvalidArgument` (src/utils/error.h),
/// mirrored here the same way as the codes above: a value out of a
/// binding-enforced range, used when the binding itself rejects the
/// argument before any C call.
constexpr int32_t kInvalidArgument = 2;

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

/// Result envelope for the numeric accessors (every `*Count`, `calcMode`).
/// `value` stays zero on failure.
struct JsNumberResult {
  JsStatus status;
  double value = 0.0;
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
  int32_t barDirection = 0;
  int32_t iconSetName = 0;
  int32_t iconIndex = 0;
};

/// JS-side mirror of `fm_sheet_view_t`.
struct JsSheetView {
  uint32_t zoomScale = 100U;
  uint32_t freezeRows = 0U;
  uint32_t freezeCols = 0U;
  int32_t tabHidden = 0;
  /// `fm_sheet_visibility_t` ordinal. `tabHidden` is its two-state view
  /// and is non-zero for both hidden states.
  int32_t visibility = 0;
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
/// diagnostics surfaced by the most recent C-ABI call. `code ==
/// kBindingInvalidHandle` is special-cased: that failure is raised before
/// any C-ABI call, so the envelope carries a synthesised message rather
/// than the residue of an earlier, unrelated call.
JsStatus error_status(int32_t code);

/// Builds a failure envelope for a rejection the binding raised itself --
/// a missing or wrongly-shaped JS argument -- carrying `message` and an
/// empty context. Never reads the thread-local diagnostics, which at that
/// point still belong to a previous call.
JsStatus binding_error_status(int32_t code, const char* message);

/// Bridges a `fm_status_t` into a `JsStatus` envelope.
JsStatus status_from_rc(fm_status_t rc);

/// Per-call validator for fields in nested binding records. Binding parsers
/// share this reader to reject lossy conversions before the C ABI mutates a
/// workbook. Every access
/// to caller-owned JavaScript values goes through a JS try/catch trampoline so
/// proxy/getter exceptions become the reader's first error under
/// `-fno-exceptions`.
/// Missing, `null`, and `undefined` retain the supplied default.
class JsNarrowNumericReader {
 public:
  explicit JsNarrowNumericReader(const char* operation);

  uint32_t u32(const emscripten::val& owner, const char* key, uint32_t dflt, const char* field = nullptr);
  int32_t i32(const emscripten::val& owner, const char* key, int32_t dflt, const char* field = nullptr);
  uint8_t u8(const emscripten::val& owner, const char* key, uint8_t dflt, const char* field = nullptr);
  uint16_t u16(const emscripten::val& owner, const char* key, uint16_t dflt, const char* field = nullptr);
  uint8_t u8_value(const emscripten::val& value, uint8_t dflt, const char* field);
  uint16_t u16_value(const emscripten::val& value, uint16_t dflt, const char* field);
  uint32_t u32_value(const emscripten::val& value, uint32_t dflt, const char* field);
  int32_t i32_value(const emscripten::val& value, int32_t dflt, const char* field);
  int64_t i64(const emscripten::val& owner, const char* key, int64_t dflt, const char* field = nullptr);
  double number(const emscripten::val& owner, const char* key, double dflt, const char* field = nullptr);
  double number_value(const emscripten::val& value, double dflt, const char* field);
  bool boolean(const emscripten::val& owner, const char* key, bool dflt, const char* field = nullptr);
  bool boolean_value(const emscripten::val& value, bool dflt, const char* field);
  std::string string(const emscripten::val& owner, const char* key, const char* field = nullptr);
  std::string string_value(const emscripten::val& value, const char* field);
  const char* optional_string(const emscripten::val& owner, const char* key, std::string& storage,
                              const char* field = nullptr);
  const char* optional_string_value(const emscripten::val& value, std::string& storage, const char* field);
  emscripten::val value(const emscripten::val& owner, const char* key, const char* field = nullptr);
  bool has_own(const emscripten::val& owner, const char* key, const char* field = nullptr);
  bool is_array(const emscripten::val& value, const char* field = nullptr);
  uint32_t length(const emscripten::val& array, const char* field = nullptr);
  emscripten::val array_element(const emscripten::val& array, uint32_t index, const char* field = nullptr);

  bool ok() const;
  const std::string& message() const;

 private:
  bool read_integer(const emscripten::val& value, const char* key, const char* field, double lower, double upper,
                    const char* range, double* number);
  bool read_number(const emscripten::val& value, const char* key, const char* field, double* number);
  emscripten::val safe_operation(const emscripten::val& owner, const emscripten::val& key, int32_t operation,
                                 bool* succeeded);
  emscripten::val safe_get(const emscripten::val& owner, const emscripten::val& key, const char* field,
                           bool* succeeded);
  void reject_access(const char* key, const char* field);
  void reject(const char* key, const char* field, const char* range);
  void reject_type(const char* field, const char* expected);

  std::string operation_;
  std::string message_;
};

/// Builds a `JsNumberResult` from `rc`, carrying `value` only on success.
JsNumberResult number_result(fm_status_t rc, double value);

/// Builds a `JsStringResult` from `rc`, carrying `value` only on success.
/// A NULL `value` becomes the empty string.
JsStringResult string_result(fm_status_t rc, const char* value);

// ---- Guarded JS callback invocation -------------------------------------
//
// Every path that calls back into caller-supplied JavaScript from inside
// an engine frame MUST go through `call_js_callback`, and no part TU may
// invoke an `emscripten::val` callable directly. An exception thrown out
// of a JS callback unwinds the intervening WASM frames without running a
// single C++ destructor: the binary is linked `-fno-exceptions` /
// `-sDISABLE_EXCEPTION_CATCHING=1`, so there is no landing pad to restore
// RAII state such as the recalc re-entry flag or the engine's write
// mutex, and the damage outlives the call. The helper moves the `try` /
// `catch` to the JS side, where it is expressible, and hands the throw
// back as a value the C++ caller can act on while unwinding normally.
//
// Its definition lives in `workbook_core.cpp` rather than
// `embind_common.cpp`: it is backed by an `EM_JS` shim, and `EM_JS` emits
// a non-static descriptor symbol that may exist in exactly one
// translation unit.

/// Outcome of a guarded JS callback invocation. The truthy / falsy split
/// mirrors JavaScript's own coercion, so a callback returning an
/// unexpected type is classified rather than cast (and cannot throw a
/// second time on the way out).
enum class JsCallbackOutcome : int32_t {
  kThrew = -1,   ///< Threw, or was not callable at all.
  kFalsy = 0,    ///< Returned a value JavaScript coerces to `false`.
  kTruthy = 1,   ///< Returned a value JavaScript coerces to `true`.
  kNoValue = 2,  ///< Returned `undefined` or `null`.
};

/// Invokes `fn` with the JS array `args`, catching anything it throws.
/// Never propagates a JS exception into the calling C++ frame.
JsCallbackOutcome call_js_callback(const emscripten::val& fn, const emscripten::val& args);

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

/// Result of a guarded byte-array read. `ok == false` means the input was not
/// the accepted Uint8Array shape or a JS proxy/typed-array operation threw;
/// `message` is suitable for a binding-level invalid-argument envelope.
struct JsBytesReadResult {
  bool ok = true;
  std::vector<uint8_t> bytes;
  std::string message;
};

/// Reads a caller-supplied Uint8Array through the JS try/catch boundary. The
/// result distinguishes an invalid input from a valid empty array so factory
/// callers can preserve their existing invalid-workbook contract.
JsBytesReadResult val_to_bytes_checked(const emscripten::val& v);

/// Materialises a fresh JS `Uint8Array` from a contiguous byte range.
/// Element-by-element copy via `set(i, val)` for layout independence.
emscripten::val bytes_to_val(const uint8_t* data, std::size_t len);

// ---- JS field readers and record builders -----------------------------
//
// Every read of a caller-supplied `emscripten::val` record goes through
// these helpers rather than an inline `v[key].isUndefined() ? ... :
// v[key].as<T>()` chain. Each embind operation expands to a JS-bridge
// stub, so one out-of-line copy per field type is markedly smaller than
// the same glue repeated at every call site.
//
// "Missing" means `undefined` or `null` throughout. Required fields keep
// reading through `val::as<T>` directly so a missing one converts exactly
// as before.

/// Reads a range record with the explicit numeric reader. Missing / nullish
/// fields retain the all-zero defaults.
fm_merge_range js_pull_range(const emscripten::val& v, JsNarrowNumericReader* reader);

/// Reads a range list with the explicit numeric reader. Other WASM families
/// must use the same per-call reader before invoking a C mutator.
std::vector<fm_merge_range> js_pull_ranges(const emscripten::val& v, const char* key, JsNarrowNumericReader* reader);

/// Pulls the `{kind, rgb, theme, tint, indexed}` colour specification out
/// of `v[key]`. An absent object leaves `kind` at `kFmColorNone`, which
/// makes the writer emit the sibling `*Argb` as literal `rgb`. A supplied
/// selector is authoritative; the binding does not resolve theme/indexed /
/// auto colours.
fm_color_spec js_pull_color_spec(const emscripten::val& v, const char* key, JsNarrowNumericReader* reader);

/// Pulls a `{style, colorArgb, color}` border-side record out of `v`,
/// defaulting every absent field to zero.
fm_border_side js_pull_border_side(const emscripten::val& v, JsNarrowNumericReader* reader);

/// One string field of a record built by `js_set_cstr_fields`.
struct JsStrField {
  const char* key;
  const char* value;  ///< NULL is emitted as the empty string.
};

/// Sets `o[key]` to `s`, or to the empty string when `s` is NULL.
void js_set_cstr(emscripten::val& o, const char* key, const char* s);

/// Sets every field of `fields` on `o` as `js_set_cstr` would.
void js_set_cstr_fields(emscripten::val& o, const JsStrField* fields, std::size_t n);

template <std::size_t N>
inline void js_set_cstr_fields(emscripten::val& o, const JsStrField (&fields)[N]) {
  js_set_cstr_fields(o, fields, N);
}

/// Builds the JS mirror of a colour specification. Reads emit it
/// unconditionally so a get / edit / add cycle passes the theme or
/// indexed colour straight back and keeps hitting style-table dedup.
emscripten::val js_color_spec(const fm_color_spec& spec);

/// Builds the JS mirror of a border side.
emscripten::val js_border_side(const fm_border_side& s);

}  // namespace parts
}  // namespace wasm
}  // namespace formulon

#endif  // FORMULON_WASM_PARTS_EMBIND_COMMON_H_
