// Shared translation helpers and module-global state for the Node.js
// N-API addon. Per-area implementation TUs (`parts/*.cc`) include this
// header to obtain the result builders (`ThrowStatus`, `ValueOrThrow`,
// `OptionalOrNull`, ...), the `Value` translation pair and the JS-shape
// builders that the bindings collectively rely on.
//
// A method returns its payload directly on success. A failure throws: a
// C-ABI status becomes the injected `FormulonError`, an argument the binding
// cannot represent a `TypeError` / `RangeError`. All helpers translate the
// C ABI types declared in `c_api/formulon_c.h` into the JS object shapes
// the WASM binding also produces, so JS callers can swap packages.

#ifndef FORMULON_NODE_ADDON_PARTS_ADDON_COMMON_H_
#define FORMULON_NODE_ADDON_PARTS_ADDON_COMMON_H_

// NOTE: `NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS` and
// `NAPI_DISABLE_CPP_EXCEPTIONS` are defined on the command line by
// `cmake/FormulonNodeAddon.cmake` so they apply uniformly to every TU
// that includes `napi.h`. They are required for the addon to compile
// under the project's `-fno-exceptions` policy.
//
// The compiler driver assigns the implicit replacement list `1` to a
// `-D X` flag, but `napi.h` later does an unconditional `#define X`
// (empty replacement list). To keep the build `-Werror`-clean we undef
// the command-line versions first; the napi.h `#ifdef NAPI_DISABLE_*`
// blocks immediately afterwards re-establish the same macros with the
// expected empty replacement list.
#ifdef NAPI_DISABLE_CPP_EXCEPTIONS
#undef NAPI_DISABLE_CPP_EXCEPTIONS
#endif
#ifdef NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS
#undef NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS
#endif
#define NAPI_DISABLE_CPP_EXCEPTIONS
#define NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS

// NOLINTNEXTLINE(misc-include-cleaner): napi.h is the canonical entry point.
#include <napi.h>

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <type_traits>
#include <utility>
#include <vector>

#include "c_api/binding_message.h"
#include "c_api/binding_status.h"
#include "c_api/formulon_c.h"

namespace formulon_node {

/// Copies borrowed bytes onto the JS heap; the source remains owned by the caller.
Napi::Uint8Array CopyBytes(Napi::Env env, const uint8_t* bytes, std::size_t len);

/// Copies an engine-owned buffer onto the JS heap and releases the source.
Napi::Uint8Array TakeSavedBytes(Napi::Env env, uint8_t* bytes, std::size_t len);

using formulon::c_api::kBindingInvalidHandle;

/// Per-environment addon state, held as the environment's instance data.
struct AddonData {
  /// The `Workbook` constructor, created once per environment.
  Napi::FunctionReference workbook_class;
  /// The `FormulonError` constructor injected through `__setErrorClass`.
  Napi::FunctionReference error_class;
};

/// This environment's addon state, created on first use.
AddonData& GetAddonData(Napi::Env env);

// ---------------------------------------------------------------------
// Result builders
// ---------------------------------------------------------------------

/// Throws a `FormulonError` carrying `rc`, the last-error message and
/// context of the C-ABI call that failed and, after a formula syntax error,
/// its `parseError` position; returns `env.Undefined()`.
/// Throws nothing when an exception is already pending (a callback's own
/// exception, or a rejected argument), so that exception reaches JS as-is.
/// Without an injected error class it throws a `TypeError` instead.
Napi::Value ThrowStatus(Napi::Env env, fm_status_t rc);

/// Throws a `FormulonError` carrying `code` and `message` for a failure the
/// binding detects before any C-ABI call; returns `env.Undefined()`.
Napi::Value ThrowBindingStatus(Napi::Env env, fm_status_t code, const std::string& message);

/// Throws the `FormulonError` of a call on a disposed workbook
/// (`BindingInvalidHandle`, "the workbook is disposed") and returns
/// `env.Undefined()`. No C-ABI call ran, so no last-error state is read.
Napi::Value ThrowInvalidHandle(Napi::Env env);

/// Throws the `FormulonError` of a refused re-entrant call
/// (`GraphRecalcReentrant`, carrying `message`) and returns `env.Undefined()`.
/// No C-ABI call ran, so no last-error state is read.
Napi::Value ThrowReentrant(Napi::Env env, const char* message);

/// Returns `ok_value` when `rc` is 0, else `ThrowStatus(env, rc)`.
Napi::Value ValueOrThrow(Napi::Env env, fm_status_t rc, Napi::Value ok_value);

/// Result of a lookup by key: `ThrowStatus` when `rc` is non-zero, `null`
/// when the key was not `present`, else `item`.
Napi::Value OptionalOrNull(Napi::Env env, fm_status_t rc, int32_t present, Napi::Value item);

/// Translates an `fm_value_t` into the JS `Value` union: `kind` plus only
/// the field that kind carries (`number`, `boolean`, `text` or `errorCode`).
Napi::Value ValueToJs(Napi::Env env, const fm_value_t& v);

/// Translates `array` into `{ rows, cols, cells }`, `cells` a rows x cols
/// nested array of `Value` objects. Does not take ownership of `array`.
Napi::Object ValueArrayToJs(Napi::Env env, const fm_value_array_t* array);

/// Reads a JS `Value` into `out`. `kind` is required, and the field its kind
/// carries must be present with its type (`boolean` must be a JS boolean);
/// other fields are ignored. A text payload is copied into `text`, which
/// must outlive the C call reading `out`. Throws a `TypeError` and returns
/// false on a malformed value.
bool JsToValue(Napi::Env env, Napi::Value in, fm_value_t* out, std::string* text);

/// A default record for an out-parameter: `FM_STRUCT_INIT(T)` for a C-ABI
/// record (one leading `struct_size`), value-initialised otherwise.
template <typename T, typename = void>
struct IsCRecord : std::false_type {};
template <typename T>
struct IsCRecord<T, std::void_t<decltype(std::declval<T&>().struct_size)>> : std::true_type {};

template <typename T>
T InitOut() {
  if constexpr (IsCRecord<T>::value) {
    return FM_STRUCT_INIT(T);
  } else {
    return T{};
  }
}

/// Fills `items` by asking `count_fn(size_t*)` for the element count and then
/// `at_fn(size_t, Item*)` for each element, converting it with `to_js(item)`.
/// Stops at the first failing status and returns it; returns 0 otherwise,
/// including when a conversion left an exception pending.
template <typename Item, typename CountFn, typename AtFn, typename ToJs>
fm_status_t CollectList(Napi::Array& items, CountFn count_fn, AtFn at_fn, ToJs to_js) {
  std::size_t count = 0;
  fm_status_t rc = count_fn(&count);
  for (std::size_t i = 0; rc == 0 && i < count; ++i) {
    Item item = InitOut<Item>();
    rc = at_fn(i, &item);
    if (rc != 0) {
      break;
    }
    items.Set(static_cast<uint32_t>(i), to_js(item));
    if (items.Env().IsExceptionPending()) {
      break;
    }
  }
  return rc;
}

/// Fills `items` from a caller-buffer C getter `call(T* out, size_t capacity, size_t* out_count)`: a sizing call
/// with capacity 0, then one call with a buffer of that size. Returns the first failing status.
template <typename T, typename Call>
fm_status_t FillCallerBuffer(std::vector<T>& items, Call&& call) {
  std::size_t count = 0;
  fm_status_t rc = call(static_cast<T*>(nullptr), static_cast<std::size_t>(0), &count);
  if (rc != 0 || count == 0) {
    items.clear();
    return rc;
  }
  items.assign(count, InitOut<T>());
  std::size_t written = 0;
  rc = call(items.data(), items.size(), &written);
  items.resize(rc == 0 ? std::min(written, items.size()) : 0);
  return rc;
}

/// The largest uint32 value as a JS Number, the upper bound of uint32 checks.
constexpr double kUint32MaxNumber = static_cast<double>(std::numeric_limits<std::uint32_t>::max());

/// Translates an `fm_pivot_cell_info_t` into the JS shape.
Napi::Object TranslatePivotCell(Napi::Env env, const fm_pivot_cell_info_t& cell);

/// Translates an `fm_cf_match_info_t` into the JS shape used by embind.
Napi::Object TranslateCfMatch(Napi::Env env, const fm_cf_match_info_t& m);

/// Translates an `fm_color_t` into `{ kind, rgb, theme, tint, indexed }`.
Napi::Object ColorSpecToJs(Napi::Env env, const fm_color_t& spec);

/// `ColorSpecToJs(spec)`, or `null` when `spec.kind` is `FM_COLOR_KIND_NONE`.
Napi::Value ColorSpecOrNull(Napi::Env env, const fm_color_t& spec);

/// Builds a `CellRect` `{ firstRow, firstCol, lastRow, lastCol }`.
Napi::Object RangeToJs(Napi::Env env, const fm_cell_rect_t& r);

/// JS string from a C string; NULL becomes "".
Napi::Value JsString(Napi::Env env, const char* s);

/// JS number from a C numeric field.
Napi::Value JsNumber(Napi::Env env, double v);

/// JS boolean from a C `int32_t` flag (non-zero is true).
Napi::Value JsBool(Napi::Env env, int32_t v);
/// An `fm_tristate_t` read as a boolean: `FM_TRISTATE_TRUE` is `true`.
Napi::Value JsTristate(Napi::Env env, int32_t v);

/// JS string from a C string, or `null` when it is NULL or empty: the C ABI
/// spells an absent string attribute as "".
Napi::Value JsStringOrNull(Napi::Env env, const char* s);

/// JS number `v`, or `null` when it is negative: the C ABI's sentinel for an
/// absent index or member.
Napi::Value JsIndexOrNull(Napi::Env env, int32_t v);

/// JS number `v` when `present` is non-zero, else `null`.
Napi::Value JsNumberOrNull(Napi::Env env, int32_t present, double v);

/// JS boolean `v != 0` when `present` is non-zero, else `null`.
Napi::Value JsBoolOrNull(Napi::Env env, int32_t present, int32_t v);

/// Builds a `PrintTitles` `{ rows, columns }` object, each axis a
/// `{ first, last }` span or `null`; `null` when neither axis repeats.
Napi::Value PrintTitlesToJs(Napi::Env env, const fm_print_titles_t& titles);

// ---------------------------------------------------------------------
/// Names the public method whose arguments the readers are parsing, so a
/// reader failure reads `Class.method <field> ...`. Scopes nest: the previous
/// label returns when the scope ends. `describe` and `name` run only on a
/// failure; `name` gives the public parameter name of a raw positional index.
class ScopedReaderLabel final {
 public:
  using Describe = std::string (*)(const void* data);
  using Name = std::string (*)(const void* data, std::size_t raw_index);
  ScopedReaderLabel(Describe describe, Name name, const void* data) noexcept;
  ~ScopedReaderLabel();
  ScopedReaderLabel(const ScopedReaderLabel&) = delete;
  ScopedReaderLabel& operator=(const ScopedReaderLabel&) = delete;

 private:
  Describe previous_describe_;
  Name previous_name_;
  const void* previous_data_;
};

/// The `Class.method` label of the innermost `ScopedReaderLabel`, or an empty string outside one.
std::string CurrentCallLabel();

/// The d.ts name of the raw positional argument `raw_index` of the call being served, or an empty string when it has
/// none: an argument the facade supplies, or a raw layout the d.ts signature does not describe.
std::string CurrentArgumentName(std::size_t raw_index);

/// Throws a TypeError `<call> <name> <constraint>` for the raw positional argument `raw_index` of the call being
/// served (see `CurrentArgumentName`); an unnamed argument words it `<call> <constraint>`.
void ThrowPositionalTypeError(Napi::Env env, std::size_t raw_index, const std::string& constraint);

/// `ThrowIntegerRefusal` for the raw positional argument `raw_index` of the call being served.
void ThrowPositionalIntegerRefusal(Napi::Env env, std::size_t raw_index, double number, double min_value);

/// Throws a TypeError `<call> <path> <constraint>` for the call being served (see `ScopedReaderLabel`); an empty
/// `path` names the call's record parameter, when it takes one.
void ThrowArgumentTypeError(Napi::Env env, const std::string& path, const std::string& constraint);

/// Throws the refusal of the number `number` given for the integer parameter `path` of `call`: a non-integral value
/// is a RangeError "must be an integer"; a negative one for a parameter whose smallest value is `min_value >= 0` is an
/// InvalidArgument FormulonError "must not be negative"; any other value is a RangeError "is too large". An empty
/// `path` words the message `<call> <constraint>`.
void ThrowIntegerRefusal(Napi::Env env, const std::string& call, const std::string& path, double number,
                         double min_value);

/// Per-call reader for nested binding object fields. Its failures are
/// prefixed with the label of the innermost `ScopedReaderLabel`, if any.
/// It owns the pending-error It owns the pending-error
/// bit for one binding invocation, so nested parsers stop touching JS objects
/// as soon as a getter or checked conversion fails.
class CheckedSpecReader final {
 public:
  explicit CheckedSpecReader(Napi::Env env) : env_(env) {}

  bool ok() const noexcept { return !failed_ && !env_.IsExceptionPending(); }

  /// Copies a non-nullish value into `out`, recording any pending getter
  /// exception on this reader. Nullish values are absent and keep defaults.
  bool Value(const Napi::Value& value, const char* key, Napi::Value* out);
  /// Reads one array element while retaining the same per-call error state.
  /// Missing/nullish elements are absent and do not fail the reader.
  bool ArrayElement(const Napi::Array& owner, uint32_t index, Napi::Value* out);
  bool String(const Napi::Object& owner, const char* key, std::string* out);
  bool String(const Napi::Value& value, const char* key, std::string* out);
  bool Object(const Napi::Object& owner, const char* key, Napi::Object* out);
  bool Object(const Napi::Value& value, const char* key, Napi::Object* out);
  bool Array(const Napi::Object& owner, const char* key, Napi::Array* out);
  bool Array(const Napi::Value& value, const char* key, Napi::Array* out);
  bool Bool(const Napi::Object& owner, const char* key, bool dflt, bool* present = nullptr);
  bool Bool(const Napi::Value& value, const char* key, bool dflt);
  /// A `boolean | null | undefined` field as an `fm_tristate_t`: absent and `null` are `FM_TRISTATE_DEFAULT`.
  int32_t Tristate(const Napi::Object& owner, const char* key);
  int32_t Tristate(const Napi::Value& value, const char* key);
  double Double(const Napi::Object& owner, const char* key, double dflt, bool* present = nullptr);
  double Double(const Napi::Value& value, const char* key, double dflt);
  uint8_t U8(const Napi::Object& owner, const char* key, uint8_t dflt, bool* present = nullptr);
  uint8_t U8(const Napi::Value& value, const char* key, uint8_t dflt);
  uint16_t U16(const Napi::Object& owner, const char* key, uint16_t dflt, bool* present = nullptr);
  uint16_t U16(const Napi::Value& value, const char* key, uint16_t dflt);
  uint32_t U32(const Napi::Object& owner, const char* key, uint32_t dflt, bool* present = nullptr);
  uint32_t U32(const Napi::Value& value, const char* key, uint32_t dflt);
  int32_t I32(const Napi::Object& owner, const char* key, int32_t dflt, bool* present = nullptr);
  int32_t I32(const Napi::Value& value, const char* key, int32_t dflt);
  /// Reads a signed int64 from a JavaScript Number without narrowing through
  /// a wider unsigned intermediate. The valid interval is [-2^63, 2^63).
  int64_t I64(const Napi::Object& owner, const char* key, int64_t dflt, bool* present = nullptr);
  int64_t I64(const Napi::Value& value, const char* key, int64_t dflt);
  /// Reads a required uint32 field: absent or a non-number is a TypeError, a
  /// non-integer or too-large value a RangeError, a negative one an InvalidArgument error.
  uint32_t RequiredU32(const Napi::Object& owner, const char* key);
  /// `Object` / `Array` for a required field: an absent one is a TypeError.
  bool RequiredObject(const Napi::Object& owner, const char* key, Napi::Object* out);
  bool RequiredArray(const Napi::Object& owner, const char* key, Napi::Array* out);
  /// Reads a required string field: absent, `null` or a non-string is a TypeError.
  std::string RequiredString(const Napi::Object& owner, const char* key);
  /// `owner[key]` as stored, `undefined` when absent; unlike the typed
  /// readers it keeps `null` apart from `undefined`.
  Napi::Value Get(const Napi::Object& owner, const char* key);
  /// Records a TypeError for `key`: "must be `expected`", `expected` a full phrase such as "a number" or "an object".
  void FailType(const char* key, const char* expected);
  /// Records a TypeError for `key`: "must not be null".
  void FailNull(const char* key);
  /// Records the refusal of the number `number` read for the integer field `key`: a non-integral value is a
  /// RangeError "must be an integer"; a negative one for a field whose smallest value is `min_value >= 0` is an
  /// InvalidArgument FormulonError "must not be negative"; any other value is a RangeError "is too large".
  void FailInteger(const char* key, double number, double min_value);

 private:
  bool ReadOptional(const Napi::Object& owner, const char* key, Napi::Value* out, bool* present = nullptr);
  bool ReadNumber(const Napi::Object& owner, const char* key, double* out, bool* present = nullptr);
  bool ReadNumber(const Napi::Value& value, const char* key, double* out);
  bool ReadInteger(const Napi::Object& owner, const char* key, double* out, double min_value, double max_value,
                   bool* present = nullptr);
  bool ReadInteger(const Napi::Value& value, const char* key, double* out, double min_value, double max_value);
  template <typename T>
  T ReadIntegral(const Napi::Object& owner, const char* key, T dflt, bool* present = nullptr);
  template <typename T>
  T ReadIntegral(const Napi::Value& value, const char* key, T dflt);
  int64_t NarrowI64(double number, const char* key, int64_t dflt);
  void ObservePending();
  template <typename ErrorType>
  void ReportFailure(const char* key, const std::string& constraint);
  void ReportNegative(const char* key);
  /// Makes the fields read next belong to `owner`: the path it was reached by, or the root.
  void Enter(const Napi::Value& owner);
  /// Names `value` by the path of `key` under the current owner, for the fields read from it later.
  void Register(const Napi::Value& value, const char* key);
  std::string PathOf(const char* key) const;

  std::vector<std::pair<Napi::Value, std::string>> paths_;
  std::string prefix_;
  bool element_ = false;
  Napi::Env env_;
  bool failed_ = false;
};

/// The state of one key of a partial-update (patch) object.
enum class PatchKey : int32_t {
  kKeep = 0,   ///< `undefined` or omitted: the stored value stays.
  kClear = 1,  ///< `null` on a nullable field: the stored value is removed.
  kSet = 2,    ///< Any other value.
};

/// Classifies `value`, read from the patch key `key`, as `PatchKey`: `undefined` keeps, anything else but `null`
/// sets, and `null` clears when `nullable`; otherwise it records a TypeError on `reader` and reads as `kKeep`.
PatchKey ClassifyPatchKey(CheckedSpecReader& reader, const Napi::Value& value, const char* key, bool nullable);

/// Reads `spec[key]` as one patch field and returns its op: an absent key keeps, `null` clears when the field is
/// clearable (a TypeError otherwise), any other value is read into `*out` and sets. `PatchBool` / `PatchI32` /
/// `PatchU32` / `PatchI64` / `PatchDouble` take fields that cannot be cleared; `PatchNullable*` take fields that
/// can. The string form copies into `*store` and points `*out` at it, so the storage must outlive the C call.
fm_patch_op_t PatchBool(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int32_t* out);
fm_patch_op_t PatchI32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int32_t* out);
fm_patch_op_t PatchU32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, uint32_t* out);
fm_patch_op_t PatchI64(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int64_t* out);
fm_patch_op_t PatchColor(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, bool nullable,
                         fm_color_t* out);
/// `PatchDouble` for a field `null` clears.
fm_patch_op_t PatchNullableDouble(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, double* out);
/// `PatchI32` for a field `null` clears.
fm_patch_op_t PatchNullableI32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int32_t* out);
/// `PatchU32` for a field `null` clears.
fm_patch_op_t PatchNullableU32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, uint32_t* out);
fm_patch_op_t PatchDouble(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, double* out);
fm_patch_op_t PatchString(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, bool nullable,
                          std::string* store, const char** out);
/// Reads the rectangle array `spec[key]` of a patch into `*out`; true when the key is given (an absent key keeps,
/// `null` is a TypeError) and every element reads.
bool PatchRects(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, std::vector<fm_cell_rect_t>* out);

/// Reads a `CellRect` `{ firstRow, firstCol, lastRow, lastCol }` object; every
/// field is required.
bool ReadCellRect(CheckedSpecReader& reader, const Napi::Object& range, fm_cell_rect_t* out);

/// Reads the `{ kind, rgb, theme, tint, indexed }` colour specification out
/// of `owner[key]`. An absent object leaves `kind` at `FM_COLOR_KIND_NONE` (no
/// colour). A supplied selector is authoritative; this binding does not
/// resolve theme/indexed / auto colours.
fm_color_t PullColorSpec(CheckedSpecReader& reader, const Napi::Object& owner, const char* key);

/// `PullColorSpec` for a value already in hand; `key` names it in a rejection.
fm_color_t ReadColorSpec(CheckedSpecReader& reader, const Napi::Value& value, const char* key);

/// Reads the `CellRect` array `value`; `key` names it in a rejection. A
/// missing array, or a missing element, reads as nothing.
bool ReadCellRects(CheckedSpecReader& reader, const Napi::Value& value, const char* key,
                   std::vector<fm_cell_rect_t>* out);

/// Reads the range at `info[idx]`; a missing/nullish argument reads as all
/// zeros, while a supplied non-object is rejected by the reader.
bool CellRectArg(CheckedSpecReader& reader, const Napi::CallbackInfo& info, size_t idx, fm_cell_rect_t* out);

// ---------------------------------------------------------------------
// Positional arguments. The registration guard in `workbook_class.cc`
// checks presence and type for every position its masks name before a
// method runs; these accessors are the only place a binding unwraps a JS
// argument, and they never coerce. A value of another type reads as the
// default (or `false`), so a method branching on one is reading an
// optional position the guard leaves to it.

/// The number at `value`, or `dflt` when it is not a number.
double NumberValue(const Napi::Value& value, double dflt);
/// `info[idx]` as a number, or `dflt` when absent or not a number.
double NumberArg(const Napi::CallbackInfo& info, size_t idx, double dflt);
/// `info[idx]` as UTF-8 (lone surrogates become U+FFFD), or "" when it is
/// not a string.
std::string StringArg(const Napi::CallbackInfo& info, size_t idx);
bool ObjectArg(const Napi::CallbackInfo& info, size_t idx, Napi::Object* out);
bool ArrayArg(const Napi::CallbackInfo& info, size_t idx, Napi::Array* out);
bool FunctionArg(const Napi::CallbackInfo& info, size_t idx, Napi::Function* out);
/// `info[idx]` when it is a `Uint8Array` (not merely another typed array).
bool BytesArg(const Napi::CallbackInfo& info, size_t idx, Napi::Uint8Array* out);

/// Builds an `fm_pivot_data_field_t` from a JS spec object. The
/// `name_buf` / `nfmt_buf` strings keep the borrowed `const char*`
/// pointers alive for the caller.
void BuildDataFieldSpec(CheckedSpecReader& reader, const Napi::Object& spec, fm_pivot_data_field_t& out,
                        std::string& name_buf, std::string& nfmt_buf);

/// Reads the optional `options` object at `info[idx]` (absent or nullish reads as `{}`); false after throwing a
/// TypeError for another type.
bool OptionsArg(const Napi::CallbackInfo& info, size_t idx, Napi::Object* out);

/// The `{ r1c1?, baseRow?, baseCol? }` options of `parseReference`, `localizeFormula` and `canonicalizeFormula`.
struct NotationOptions {
  bool r1c1 = false;
  bool has_base = false;
  uint32_t base_row = 0;
  uint32_t base_col = 0;
};

/// Reads the optional notation options at `info[idx]`; false after throwing: a non-object, a non-boolean `r1c1`, a
/// coordinate that is not a uint32, or one base coordinate without the other.
bool NotationOptionsArg(const Napi::CallbackInfo& info, size_t idx, NotationOptions* out);

/// `options` as the C record `Spec` (`fm_parse_reference_spec_t`, `fm_formula_notation_spec_t`).
template <typename Spec>
Spec NotationSpec(const NotationOptions& options) {
  Spec spec = FM_STRUCT_INIT(Spec);
  spec.r1c1 = options.r1c1 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  spec.has_base = options.has_base ? 1 : 0;
  spec.base_row = options.base_row;
  spec.base_col = options.base_col;
  return spec;
}

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_ADDON_COMMON_H_
