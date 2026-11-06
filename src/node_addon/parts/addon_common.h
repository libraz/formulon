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

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "c_api/binding_status.h"
#include "c_api/formulon_c.h"

namespace formulon_node {

/// Copies borrowed bytes onto the JS heap; the source remains owned by the caller.
Napi::Uint8Array CopyBytes(Napi::Env env, const uint8_t* bytes, std::size_t len);

/// Copies an engine-owned buffer onto the JS heap and releases the source.
Napi::Uint8Array TakeSavedBytes(Napi::Env env, uint8_t* bytes, std::size_t len);

using formulon::c_api::kBindingInvalidHandle;
using formulon::c_api::kInvalidArgument;

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

/// Throws a `FormulonError` carrying `rc` and the last-error message and
/// context of the C-ABI call that failed, and returns `env.Undefined()`.
/// Throws nothing when an exception is already pending (a callback's own
/// exception, or a rejected argument), so that exception reaches JS as-is.
/// Without an injected error class it throws a `TypeError` instead.
Napi::Value ThrowStatus(Napi::Env env, fm_status_t rc);

/// Throws the `FormulonError` of a call on a disposed workbook
/// (`BindingInvalidHandle`, "the workbook is disposed") and returns
/// `env.Undefined()`. No C-ABI call ran, so no last-error state is read.
Napi::Value ThrowInvalidHandle(Napi::Env env);

/// Returns `ok_value` when `rc` is 0, else `ThrowStatus(env, rc)`.
Napi::Value ValueOrThrow(Napi::Env env, fm_status_t rc, Napi::Value ok_value);

/// Result of a lookup by key: `ThrowStatus` when `rc` is non-zero, `null`
/// when the key was not `present`, else `item`.
Napi::Value OptionalOrNull(Napi::Env env, fm_status_t rc, int32_t present, Napi::Value item);

/// Translates an `fm_value_t` into the JS `Value` union: `kind` plus only
/// the field that kind carries (`number`, `boolean`, `text` or `errorCode`).
Napi::Value ValueToJs(Napi::Env env, const fm_value_t& v);

/// Reads a JS `Value` into `out`. `kind` is required, and the field its kind
/// carries must be present with its type (`boolean` must be a JS boolean);
/// other fields are ignored. A text payload is copied into `text`, which
/// must outlive the C call reading `out`. Throws a `TypeError` and returns
/// false on a malformed value.
bool JsToValue(Napi::Env env, Napi::Value in, fm_value_t* out, std::string* text);

/// Fills `items` by asking `count_fn(size_t*)` for the element count and then
/// `at_fn(size_t, Item*)` for each element, converting it with `to_js(item)`.
/// Stops at the first failing status and returns it; returns 0 otherwise,
/// including when a conversion left an exception pending.
template <typename Item, typename CountFn, typename AtFn, typename ToJs>
fm_status_t CollectList(Napi::Array& items, CountFn count_fn, AtFn at_fn, ToJs to_js) {
  std::size_t count = 0;
  fm_status_t rc = count_fn(&count);
  for (std::size_t i = 0; rc == 0 && i < count; ++i) {
    Item item{};
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

/// The largest uint32 value as a JS Number, the upper bound of uint32 checks.
constexpr double kUint32MaxNumber = static_cast<double>(std::numeric_limits<std::uint32_t>::max());

/// Translates an `fm_pivot_cell_t` into the JS shape.
Napi::Object TranslatePivotCell(Napi::Env env, const fm_pivot_cell_t& cell);

/// Translates an `fm_cf_color_t` into the JS shape used by embind.
Napi::Object TranslateCfColor(Napi::Env env, const fm_cf_color_t& c);

/// Translates an `fm_cf_match_t` into the JS shape used by embind.
Napi::Object TranslateCfMatch(Napi::Env env, const fm_cf_match_t& m);

/// Translates an `fm_color_spec_t` into `{ kind, rgb, theme, tint, indexed }`.
Napi::Object ColorSpecToJs(Napi::Env env, const fm_color_spec_t& spec);

/// Builds a `CellRect` `{ firstRow, firstCol, lastRow, lastCol }`.
Napi::Object RangeToJs(Napi::Env env, const fm_cell_rect_t& r);

/// JS string from a C string; NULL becomes "".
Napi::Value JsString(Napi::Env env, const char* s);

/// JS number from a C numeric field.
Napi::Value JsNumber(Napi::Env env, double v);

/// JS boolean from a C `int32_t` flag (non-zero is true).
Napi::Value JsBool(Napi::Env env, int32_t v);

// ---------------------------------------------------------------------
/// Per-call reader for nested binding object fields. It owns the pending-error
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
  /// Reads a required uint32 field: absent, non-number or non-integer is a
  /// TypeError, a negative or too-large integer a RangeError.
  uint32_t RequiredU32(const Napi::Object& owner, const char* key);
  /// `Object` / `Array` for a required field: an absent one is a TypeError.
  bool RequiredObject(const Napi::Object& owner, const char* key, Napi::Object* out);
  bool RequiredArray(const Napi::Object& owner, const char* key, Napi::Array* out);

 private:
  bool ReadOptional(const Napi::Object& owner, const char* key, Napi::Value* out, bool* present = nullptr);
  bool ReadNumber(const Napi::Object& owner, const char* key, double* out, bool* present = nullptr);
  bool ReadNumber(const Napi::Value& value, const char* key, double* out);
  bool ReadInteger(const Napi::Object& owner, const char* key, double* out, double min_value, double max_value,
                   const char* range_name, bool* present = nullptr);
  bool ReadInteger(const Napi::Value& value, const char* key, double* out, double min_value, double max_value,
                   const char* range_name);
  template <typename T>
  T ReadIntegral(const Napi::Object& owner, const char* key, T dflt, const char* range_name, bool* present = nullptr);
  template <typename T>
  T ReadIntegral(const Napi::Value& value, const char* key, T dflt, const char* range_name);
  int64_t NarrowI64(double number, const char* key, int64_t dflt);
  void ObservePending();
  template <typename ErrorType>
  void ReportFailure(const std::string& message);
  void ReportType(const char* key, const char* expected);
  void ReportRange(const char* key, const char* range_name);

  Napi::Env env_;
  bool failed_ = false;
};

/// Reads a `CellRect` `{ firstRow, firstCol, lastRow, lastCol }` object; every
/// field is required.
bool ReadCellRect(CheckedSpecReader& reader, const Napi::Object& range, fm_cell_rect_t* out);

/// Reads the `{ kind, rgb, theme, tint, indexed }` colour specification out
/// of `owner[key]`. An absent object leaves `kind` at `kFmColorNone`, so
/// the writer emits the sibling `*Argb` as literal `rgb`. A supplied
/// selector is authoritative; this binding does not resolve theme/indexed /
/// auto colours.
fm_color_spec_t PullColorSpec(CheckedSpecReader& reader, const Napi::Object& owner, const char* key);

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

/// Builds an `fm_pivot_data_field_spec_t` from a JS spec object. The
/// `name_buf` / `nfmt_buf` strings keep the borrowed `const char*`
/// pointers alive for the caller.
void BuildDataFieldSpec(CheckedSpecReader& reader, const Napi::Object& spec, fm_pivot_data_field_spec_t& out,
                        std::string& name_buf, std::string& nfmt_buf);

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_ADDON_COMMON_H_
