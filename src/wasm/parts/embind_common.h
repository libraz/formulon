//
// Shared types and helpers for the Emscripten binding TUs under
// `src/wasm/parts/`. The whole bundle is compiled only when
// `FM_BUILD_WASM=ON` and exists purely to keep the per-area binding
// surface manageable -- each `parts/*.cpp` declares one slice of the
// `JsWorkbook` class plus its own `EMSCRIPTEN_BINDINGS` block or
// dispatches into the central registrar in `bindings_register.cpp`.
//
// Every binding method returns its payload directly. A failure is recorded
// through `record_failure` (or one of the builders over it) and the method
// returns the type's default value; the generated pre-js wrapper throws the
// recorded failure as a `FormulonError` once the native frame has returned,
// so no JS exception ever unwinds a C++ frame.
//
// What lives here:
//   * Value-object structs (`JsSheetView`, ...) for the
//     records the JS-facing API returns.
//   * Failure builders (`record_failure`, `succeeded`, `failed`,
//     `optional_or_null`) and the `Value` converters (`value_to_val`,
//     `val_to_value`).
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
#include <string_view>
#include <vector>

#include "c_api/binding_status.h"
#include "c_api/formulon_c.h"

namespace formulon {
namespace wasm {
namespace parts {

using formulon::c_api::kBindingCallbackException;
using formulon::c_api::kBindingInvalidHandle;

/// Largest integer a JS number holds exactly (`Number.MAX_SAFE_INTEGER`).
constexpr double kMaxSafeInteger = 9007199254740991.0;

/// Largest `uint32_t`, as the double bound handed to the numeric readers.
constexpr double kMaxU32 = 4294967295.0;

// ---- Value-object mirrors -----------------------------------------------
//
// Every cross-boundary record is reflected through a POD struct so embind
// can serialise it via `value_object<>`. The fields use signed `int32_t`
// for boolean-shaped enums because embind tends to surface those more
// cleanly than `uint8_t`; the few unsigned fields are kept `uint32_t` to
// match the upstream `fm_*` ABI shape.

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

/// The bytes of an explicit-format save with the loss counters the write
/// produced. Counters a given container cannot produce stay zero.
struct JsSaveDiagnostics {
  emscripten::val bytes = emscripten::val::null();
  uint32_t downgradedFormulaCount = 0;
  uint32_t deferredFeatureCount = 0;
  uint32_t droppedPartCount = 0;
  uint32_t droppedRelationshipCount = 0;
  uint32_t renumberedPartCount = 0;
};

/// Recovery / passthrough counters captured while loading a workbook, for
/// either container format.
struct JsReadDiagnostics {
  uint32_t undecodedFormulaCount = 0;
  uint32_t undecodedDefinedNameCount = 0;
  uint32_t undecodedPartCount = 0;
  uint32_t skippedFeatureCount = 0;
  uint32_t unknownContentTypeCount = 0;
};

/// JS-side mirror of `fm_sheet_view_info_t`.
struct JsSheetView {
  uint32_t zoomPct = 100U;
  uint32_t freezeRows = 0U;
  uint32_t freezeCols = 0U;
  bool showGridLines = true;
  bool showRowColumnHeaders = true;
  bool showZeros = true;
  bool rightToLeft = false;
  bool tabSelected = false;
  /// `fm_sheet_view_mode_t` ordinal.
  int32_t viewMode = 0;
  uint32_t topLeftRow = 0U;
  uint32_t topLeftCol = 0U;
  uint32_t paneTopLeftRow = 0U;
  uint32_t paneTopLeftCol = 0U;
  double splitXTwips = 0.0;
  double splitYTwips = 0.0;
  bool showFormulas = false;
  bool showOutlineSymbols = true;
};

// ---- Failure builders ---------------------------------------------------
//
// A failure is recorded into the pre-js wrapper's per-call slot and thrown
// there as a `FormulonError` after the native method returns. An argument
// rejection recorded in the same call takes precedence. Outside a wrapped
// call there is no slot, and the record is dropped.

/// Records `rc` as the call's failure, carrying the message and context of
/// the most recent C-ABI call and, after a formula syntax error, its
/// `parseError` position. `kBindingInvalidHandle` records as
/// `record_invalid_handle` does, since no C-ABI call produced it.
void record_failure(fm_status_t rc);

/// Records a `BindingInvalidHandle` failure for a call on a disposed
/// workbook, with the message "the workbook is disposed".
void record_invalid_handle();

/// Records a failure the binding raised itself, carrying `message` and a
/// (possibly empty) context. Never reads the thread-local diagnostics, which at that
/// point still belong to a previous call.
void record_binding_failure(int32_t code, const char* message, const char* context = "");

/// Records a `TypeError` the wrapper throws once the method returns: an
/// argument of the wrong shape (not an array, not a function, missing).
void record_type_error(const char* message);

/// Records a `RangeError` the wrapper throws once the method returns: a
/// number outside the integer range the parameter accepts.
void record_range_error(const char* message);

/// True when `rc` is 0; otherwise records it through `record_failure` and
/// returns false.
bool succeeded(fm_status_t rc);

/// Records `rc` through `record_failure` and returns `T`'s default value,
/// which the wrapper's throw keeps from reaching the caller.
template <class T>
T failed(fm_status_t rc) {
  record_failure(rc);
  return T{};
}

/// Result of a lookup by key: `item` when `rc` is 0 and the key was
/// `present`, `null` when it is 0 and absent, and `undefined` after
/// recording the failure otherwise.
emscripten::val optional_or_null(fm_status_t rc, int32_t present, emscripten::val item);

/// Returns true when a JS value is neither `undefined` nor `null`.
bool js_value_present(const emscripten::val& value);

/// Returns `value`, or a fresh empty object when it is `undefined` or `null`.
emscripten::val js_object_or_empty(const emscripten::val& value);

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
  /// Reads a required uint32 field: absent or non-number is a type error, a
  /// non-integer or too-large value a range error, a negative one InvalidArgument.
  uint32_t required_u32(const emscripten::val& owner, const char* key, const char* field);
  int64_t i64(const emscripten::val& owner, const char* key, int64_t dflt, const char* field = nullptr);
  double number(const emscripten::val& owner, const char* key, double dflt, const char* field = nullptr);
  double number_value(const emscripten::val& value, double dflt, const char* field);
  bool boolean(const emscripten::val& owner, const char* key, bool dflt, const char* field = nullptr);
  /// Accepts `true` / `false` only; absent yields `dflt`, anything else is a type error. A null `field` names
  /// the field read last.
  bool boolean_value(const emscripten::val& value, bool dflt, const char* field = nullptr);
  std::string string(const emscripten::val& owner, const char* key, const char* field = nullptr);
  std::string string_value(const emscripten::val& value, const char* field);
  const char* optional_string(const emscripten::val& owner, const char* key, std::string& storage,
                              const char* field = nullptr);
  const char* optional_string_value(const emscripten::val& value, std::string& storage, const char* field);
  emscripten::val value(const emscripten::val& owner, const char* key, const char* field = nullptr);
  bool is_array(const emscripten::val& value, const char* field = nullptr);
  /// `is_array`, failing the reader with a type error when `value` is not an array.
  bool require_array(const emscripten::val& value, const char* field);
  /// Fails the reader with a type error unless `value` is a non-null object.
  bool require_object(const emscripten::val& value, const char* field);
  /// Reads a required string field: absent, `null` or a non-string is a type error.
  std::string required_string(const emscripten::val& owner, const char* key, const char* field);
  /// Records a type error for `field` (it "must be `expected`") unless one is recorded already.
  void fail_type(const char* field, const char* expected);
  /// Records InvalidArgument for `field` (it "must be in" the `[lo, hi]` span `range`).
  void fail_range(const char* field, const char* range);
  /// Records a type error for `field` (it "must not be null").
  void fail_null(const char* field);
  /// Records a type error for `field` (it "is required").
  void fail_required(const char* field);
  uint32_t length(const emscripten::val& array, const char* field = nullptr);
  emscripten::val array_element(const emscripten::val& array, uint32_t index, const char* field = nullptr);

  /// True until the reader has recorded its first rejection.
  bool ok() const;
  const std::string& message() const;

 private:
  /// Reads `owner[key]`; `*fetched` is false when the reader already failed or the access threw.
  emscripten::val fetch(const emscripten::val& owner, const char* key, const char* field, bool* fetched);
  /// Reads `value` as an integer in [lower, upper]: absent yields `dflt`, a rejected value 0.
  double integer_value(const emscripten::val& value, double dflt, const char* field, double lower, double upper);
  bool safe_predicate(const emscripten::val& owner, const emscripten::val& key, int32_t operation, const char* key_name,
                      const char* field);
  /// True when `value` is present and a JS Number; a present non-Number fails the reader.
  bool present_number(const emscripten::val& value, const char* key, const char* field);
  bool read_integer(const emscripten::val& value, const char* key, const char* field, double lower, double upper,
                    double* number);
  bool read_number(const emscripten::val& value, const char* key, const char* field, double* number);
  emscripten::val safe_operation(const emscripten::val& owner, const emscripten::val& key, int32_t operation,
                                 bool* succeeded);
  emscripten::val safe_get(const emscripten::val& owner, const emscripten::val& key, const char* field,
                           bool* succeeded);
  void reject_access(const char* key, const char* field);
  /// Classifies a rejected integer: non-integral, negative for an unsigned field, or too large.
  void reject_integer(const char* field, double candidate, double lower);
  void reject_negative(const char* field, double candidate);
  /// Records `<operation> <path> <constraint>`; `kind` is the argument-error class the wrapper throws.
  void reject_with(const char* field, const char* constraint, int32_t kind);
  /// The record path of `field` under the operation's record parameter.
  std::string path_of(const char* field) const;
  void reject_type(const char* field, const char* expected);

  std::string operation_;
  std::string message_;
  /// The field most recently read through `value()` / `fetch()`, named by a field-less boolean read.
  std::string last_field_;
};

/// The state of one key of a partial-update (patch) object.
enum class JsPatchKey : int32_t {
  kKeep = 0,   ///< `undefined` or omitted: the stored value stays.
  kClear = 1,  ///< `null` on a nullable field: the stored value is removed.
  kSet = 2,    ///< Any other value.
};

/// Classifies `value`, read from a patch, as `JsPatchKey`. When `nullable`
/// is false a `null` records a type error on `reader` and reads as `kKeep`.
JsPatchKey js_patch_key(JsNarrowNumericReader& reader, const emscripten::val& value, const char* field, bool nullable);

/// Reads the boolean `owner[key]` of a patch into an `_op` / value pair: an omitted key leaves `KEEP`, a
/// boolean writes `SET`; `null` records a type error (a flag cannot be cleared).
void js_patch_flag(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field,
                   fm_patch_op_t* op, int32_t* value);

/// Reads the string `owner[key]` of a patch into an `_op` / pointer pair: an omitted key leaves `KEEP`, a
/// string writes `SET` (the pointer views `storage`); `null` writes `CLEAR` when `nullable`, else records a type error.
void js_patch_string(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field,
                     bool nullable, std::string& storage, fm_patch_op_t* op, const char** value);

/// As `js_patch_number` for an unsigned 32-bit key.
void js_patch_u32(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field,
                  bool nullable, fm_patch_op_t* op, uint32_t* value);

/// As `js_patch_u32` for an 8-bit key that cannot be cleared, stored widened to `int32_t`.
void js_patch_u8(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field,
                 fm_patch_op_t* op, int32_t* value);

/// Reads the number `owner[key]` of a patch into an `_op` / value pair: an omitted key leaves `KEEP`, a number writes
/// `SET`; `null` writes `CLEAR` when `nullable`, else records a type error.
void js_patch_number(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field,
                     bool nullable, fm_patch_op_t* op, double* value);

/// As `js_patch_number` for an integer-valued key.
void js_patch_int(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field,
                  bool nullable, fm_patch_op_t* op, int32_t* value);

/// Reads `owner[key]` of a patch into `*v`: an omitted key leaves `op` and returns false, `null` writes `CLEAR` when
/// `nullable` (else records a type error) and returns false, any other value writes `SET` and returns true.
bool js_patch_begin(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field,
                    bool nullable, fm_patch_op_t* op, emscripten::val* v);

/// Reads `owner[key]` into an `_op` / value pair for any value type: an omitted key keeps, `null` is `CLEAR` when
/// `nullable` (else a type error), anything else is `SET` with the value `read` extracts.
template <class T, class Read>
void js_patch_value(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field,
                    bool nullable, fm_patch_op_t* op, T* value, Read read) {
  emscripten::val v;
  if (js_patch_begin(reader, owner, key, field, nullable, op, &v)) {
    *value = read(v);
  }
}

/// Reads the optional boolean `owner[key]` as an `fm_tristate_t`: an omitted or `null` value is
/// `FM_TRISTATE_DEFAULT`, `true` / `false` are `FM_TRISTATE_TRUE` / `FM_TRISTATE_FALSE`.
int32_t js_tristate(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key, const char* field);

/// Count getter of a count-then-`at` collection scoped to a sheet.
using JsListCount = fm_status_t (*)(const fm_workbook_t*, size_t, size_t*);

/// Builds entry `index` of the collection scoped to `sheet` into `*out`.
using JsListItem = fm_status_t (*)(const fm_workbook_t*, size_t, size_t, emscripten::val*);

/// Appends every entry of a count-then-`at` collection to the JS array `list`,
/// stopping at the first failing call. Returns that status, or
/// `kBindingInvalidHandle` for a NULL `handle`.
fm_status_t js_collect_list(const fm_workbook_t* handle, size_t sheet, JsListCount count, JsListItem item,
                            emscripten::val& list);

/// Reads the element count of a count-then-`at` collection scoped to
/// `sheet`, recording a failure (`kBindingInvalidHandle` for a NULL
/// `handle`) and returning 0 when the read fails.
double js_count(const fm_workbook_t* handle, size_t sheet, JsListCount count);

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
/// Never propagates a JS exception into the calling C++ frame; a thrown
/// exception is recorded so the enclosing wrapped call rethrows that same
/// object once it returns.
JsCallbackOutcome call_js_callback(const emscripten::val& fn, const emscripten::val& args);

// ---- Translation helpers -----------------------------------------------

/// Builds the JS `Value` for `v`: `{kind}` plus only the field that kind
/// carries (`number`, a JS boolean `boolean`, a deep-copied `text`, or
/// `errorCode`).
emscripten::val value_to_val(const fm_value_t& v);

/// Reads a caller-supplied `Value` into `*out`. `kind` is required, and
/// only the field that kind carries is read: `boolean` must be a JS
/// boolean, `number` a Number, `text` a string, `errorCode` an int32.
/// `text` owns the storage a `FM_VALUE_KIND_TEXT` payload points at, so it must
/// outlive the C call. Records a TypeError and returns false on any
/// missing or wrongly-typed field.
bool val_to_value(const emscripten::val& in, fm_value_t* out, std::string& text);

/// Builds `{rows, cols, cells}` from `array`, `cells` a rows x cols nested
/// array of `Value` objects. Does not take ownership of `array`.
emscripten::val value_array_to_val(const fm_value_array_t* array);

/// Translates an `fm_cf_match_info_t` POD into the embind value-object
/// mirror.
emscripten::val translate_cf_match(const fm_cf_match_info_t& m);

/// Builds a `CellRect` `{firstRow, firstCol, lastRow, lastCol}` JS object.
emscripten::val cell_rect_to_val(const fm_cell_rect_t& m);

/// The whole grid, for a list read whose range was omitted.
constexpr fm_cell_rect_t kWholeSheetRect{0U, 0U, 1048575U, 16383U};

/// Builds a `PrintTitles` `{ rows, columns }` object, each axis a
/// `{ first, last }` span or `null`; `null` when neither axis repeats.
emscripten::val print_titles_to_val(const fm_print_titles_t& titles);

/// Builds a JS array of `CellRect` objects from `count` rectangles.
emscripten::val cell_rects_to_val(const fm_cell_rect_t* rects, std::size_t count);

/// Translates one `fm_pivot_cell_info_t` into a JS object suitable for
/// inclusion in the pivotLayout `cells` array.
emscripten::val pivot_cell_to_val(const fm_pivot_cell_info_t& cell);

/// Copies a caller-supplied Uint8Array into `*out` through the JS try/catch
/// boundary. A non-Uint8Array records a TypeError, and a throwing proxy or
/// typed-array hook records its exception; both return false. A valid empty
/// array returns true with `*out` empty.
bool val_to_bytes_checked(const emscripten::val& v, std::vector<uint8_t>* out);

/// Materialises a fresh JS `Uint8Array` from a contiguous byte range
/// (a single bulk `set` from a transient WASM memory view).
emscripten::val bytes_to_val(const uint8_t* data, std::size_t len);

/// Returns a snapshot call's bytes as a fresh `Uint8Array`, or `undefined`
/// after recording the failure when `rc` is non-zero; releases `data`
/// either way.
emscripten::val snapshot_bytes(fm_status_t rc, uint8_t* data, std::size_t len);

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

/// Reads a `CellRect` record with the explicit numeric reader; every field is
/// required, and `field` (the record's path under the call's record parameter)
/// names it in a rejection. A missing / nullish record reads as all zeros.
fm_cell_rect_t js_pull_range(const emscripten::val& v, JsNarrowNumericReader* reader, std::string_view field = "range");

/// Reads the `CellRect` array `array` itself (a positional argument); `field`
/// names it in a rejection. A missing array reads as empty.
std::vector<fm_cell_rect_t> js_pull_range_array(const emscripten::val& array, const char* field,
                                                JsNarrowNumericReader* reader);

/// Pulls the `{kind, argb, theme, tint, indexed}` colour specification out
/// of `v[key]`. An absent object leaves `kind` at `FM_COLOR_KIND_NONE` (no
/// colour). A supplied selector is authoritative; the binding does not
/// resolve theme/indexed / auto colours. `parent` is the path of `v` under the
/// call's record parameter.
fm_color_t js_pull_color_spec(const emscripten::val& v, const char* key, JsNarrowNumericReader* reader,
                              std::string_view parent = {});

/// `js_pull_color_spec` for a value already in hand; `field` is its path, which
/// a rejection of it or of one of its members names. A missing value reads as
/// `FM_COLOR_KIND_NONE`.
fm_color_t js_pull_color_spec_value(const emscripten::val& value, const char* field, JsNarrowNumericReader* reader);

/// Pulls a `{style, color}` border-side record out of `v`, defaulting every
/// absent field to zero; `field` is its path.
fm_border_side_t js_pull_border_side(const emscripten::val& v, std::string_view field, JsNarrowNumericReader* reader);

/// One string field of a record built by `js_set_cstr_fields`.
struct JsStrField {
  const char* key;
  const char* value;  ///< NULL is emitted as the empty string.
};

/// Copies `s`, or returns the empty string when `s` is NULL.
std::string string_from_cstr(const char* s);

/// Sets `o[key]` to `s`, or to the empty string when `s` is NULL.
void js_set_cstr(emscripten::val& o, const char* key, const char* s);

/// Sets every field of `fields` on `o` as `js_set_cstr` would.
void js_set_cstr_fields(emscripten::val& o, const JsStrField* fields, std::size_t n);

template <std::size_t N>
inline void js_set_cstr_fields(emscripten::val& o, const JsStrField (&fields)[N]) {
  js_set_cstr_fields(o, fields, N);
}

/// `s` as a JS string, or `null` when `s` is NULL or empty: the C ABI spells
/// an absent string attribute as "".
emscripten::val cstr_or_null(const char* s);

/// Sets `o[key]` to `cstr_or_null(s)`.
void js_set_cstr_or_null(emscripten::val& o, const char* key, const char* s);

/// Sets every field of `fields` on `o` as `js_set_cstr_or_null` would.
void js_set_cstr_or_null_fields(emscripten::val& o, const JsStrField* fields, std::size_t n);

template <std::size_t N>
inline void js_set_cstr_or_null_fields(emscripten::val& o, const JsStrField (&fields)[N]) {
  js_set_cstr_or_null_fields(o, fields, N);
}

/// Sets `o[key]` to `value` when `present` is non-zero, else to `null`.
template <class T>
inline void js_set_present_or_null(emscripten::val& o, const char* key, int32_t present, T value) {
  if (present != 0) {
    o.set(key, value);
  } else {
    o.set(key, emscripten::val::null());
  }
}

/// Builds the JS mirror of a colour specification. Reads emit it
/// unconditionally so a get / edit / add cycle passes the theme or
/// indexed colour straight back and keeps hitting style-table dedup.
emscripten::val js_color_spec(const fm_color_t& spec);

/// `js_color_spec(spec)`, or `null` when `spec.kind` is `FM_COLOR_KIND_NONE`.
emscripten::val js_color_spec_or_null(const fm_color_t& spec);

/// Builds the JS mirror of a border side.
emscripten::val js_border_side(const fm_border_side_t& s);

}  // namespace parts
}  // namespace wasm
}  // namespace formulon

#endif  // FORMULON_WASM_PARTS_EMBIND_COMMON_H_
