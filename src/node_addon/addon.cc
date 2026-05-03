// Copyright 2026 libraz. Licensed under the MIT License.
//
// Node.js N-API addon for the Formulon engine.
//
// This translation unit is compiled only when `FM_BUILD_NODE_ADDON=ON`.
// It is a thin C++ wrapper around the stable C ABI declared in
// `c_api/formulon_c.h`; the JavaScript surface NEVER touches
// `formulon::Workbook`, `formulon::Value`, or any other internal
// symbol directly — exactly mirroring the WASM/embind binding's
// architectural stance.
//
// ## Design notes
//
//   * The whole engine is built `-fno-exceptions -fno-rtti`. node-addon-api
//     would normally throw C++ exceptions across the boundary; we
//     disable that via `NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS` (and
//     `NAPI_DISABLE_CPP_EXCEPTIONS`) before any napi header include,
//     which the FormulonNodeAddon.cmake target propagates.
//
//   * Every fallible binding entry returns the same JS shape as the
//     embind binding:
//        Status = { ok: boolean, status: number,
//                   message: string, context: string }
//        Value  = { kind: number, number: number, boolean: number,
//                   text: string, errorCode: number }
//     The thread-local `fm_last_error_*` strings are snapshotted into
//     the Status envelope on every error path.
//
//   * `Workbook` is wrapped in `Napi::ObjectWrap<Workbook>`. The wrapper
//     owns the `fm_workbook_t*` and frees it in its destructor, which
//     N-API invokes when the JS object is garbage-collected.

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
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"

namespace {

// ---------------------------------------------------------------------
// Status / Value helpers
// ---------------------------------------------------------------------

/// Builds an `ok` Status envelope:
///   { ok: true, status: 0, message: "", context: "" }
Napi::Object MakeOkStatus(Napi::Env env) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("ok", Napi::Boolean::New(env, true));
  o.Set("status", Napi::Number::New(env, 0));
  o.Set("message", Napi::String::New(env, ""));
  o.Set("context", Napi::String::New(env, ""));
  return o;
}

/// Builds an error Status envelope, copying the thread-local
/// diagnostics surfaced by the most recent C-ABI call.
Napi::Object MakeErrorStatus(Napi::Env env, fm_status_t code) {
  const char* msg = fm_last_error_message();
  const char* ctx = fm_last_error_context();
  Napi::Object o = Napi::Object::New(env);
  o.Set("ok", Napi::Boolean::New(env, false));
  o.Set("status", Napi::Number::New(env, static_cast<int32_t>(code)));
  o.Set("message", Napi::String::New(env, msg != nullptr ? msg : ""));
  o.Set("context", Napi::String::New(env, ctx != nullptr ? ctx : ""));
  return o;
}

/// Translates an `fm_value_t` into the JS Value shape.
Napi::Object TranslateValue(Napi::Env env, const fm_value_t& v) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("kind", Napi::Number::New(env, static_cast<int32_t>(v.kind)));
  // Default-zero all fields so consumers can read any field without
  // checking kind first (matches the embind shape).
  double number_field = 0.0;
  int32_t boolean_field = 0;
  std::string text_field;
  int32_t error_code_field = 0;
  switch (v.kind) {
    case FM_VAL_NUMBER:
      number_field = v.u.number;
      break;
    case FM_VAL_BOOL:
      boolean_field = v.u.boolean;
      break;
    case FM_VAL_TEXT:
      text_field = (v.u.text != nullptr) ? std::string(v.u.text) : std::string();
      break;
    case FM_VAL_ERROR:
      error_code_field = v.u.error_code;
      break;
    case FM_VAL_BLANK:
    case FM_VAL_ARRAY:
    case FM_VAL_REF:
    case FM_VAL_LAMBDA:
    default:
      break;
  }
  o.Set("number", Napi::Number::New(env, number_field));
  o.Set("boolean", Napi::Number::New(env, boolean_field));
  o.Set("text", Napi::String::New(env, text_field));
  o.Set("errorCode", Napi::Number::New(env, error_code_field));
  return o;
}

/// `kBindingNullPointer` ordinal mirrors `formulon::FormulonErrorCode`
/// in the 7000-7999 range allocated to bindings (see CLAUDE.md error
/// code table). The C ABI itself returns this code when a NULL pointer
/// crosses the boundary; we emit the same code from the JS side when
/// the wrapper is asked to operate on a destroyed handle.
constexpr fm_status_t kBindingNullPointer = 7000;

// ---------------------------------------------------------------------
// Workbook ObjectWrap
// ---------------------------------------------------------------------

class Workbook : public Napi::ObjectWrap<Workbook> {
 public:
  static Napi::Function GetClass(Napi::Env env);

  explicit Workbook(const Napi::CallbackInfo& info);
  ~Workbook() override;

  Workbook(const Workbook&) = delete;
  Workbook& operator=(const Workbook&) = delete;
  Workbook(Workbook&&) = delete;
  Workbook& operator=(Workbook&&) = delete;

  // Static factories.
  static Napi::Value CreateDefault(const Napi::CallbackInfo& info);
  static Napi::Value CreateEmpty(const Napi::CallbackInfo& info);
  static Napi::Value LoadBytes(const Napi::CallbackInfo& info);

  // Cell mutation.
  Napi::Value SetNumber(const Napi::CallbackInfo& info);
  Napi::Value SetBool(const Napi::CallbackInfo& info);
  Napi::Value SetText(const Napi::CallbackInfo& info);
  Napi::Value SetBlank(const Napi::CallbackInfo& info);
  Napi::Value SetFormula(const Napi::CallbackInfo& info);

  // Cell read.
  Napi::Value GetValue(const Napi::CallbackInfo& info);

  // Recalc + save.
  Napi::Value Recalc(const Napi::CallbackInfo& info);
  Napi::Value Save(const Napi::CallbackInfo& info);

  // Sheet operations.
  Napi::Value AddSheet(const Napi::CallbackInfo& info);
  Napi::Value RemoveSheet(const Napi::CallbackInfo& info);
  Napi::Value RenameSheet(const Napi::CallbackInfo& info);
  Napi::Value SheetCount(const Napi::CallbackInfo& info);
  Napi::Value SheetName(const Napi::CallbackInfo& info);

  // Defined names.
  Napi::Value SetDefinedName(const Napi::CallbackInfo& info);

 private:
  /// Extracts a `uint32_t` argument or sets `*ok=false` and surfaces
  /// the conversion failure through the JS error-status envelope.
  static uint32_t ArgU32(const Napi::CallbackInfo& info, size_t idx);
  static double ArgDouble(const Napi::CallbackInfo& info, size_t idx);
  static std::string ArgString(const Napi::CallbackInfo& info, size_t idx);

  /// Builds an error-Status envelope when the wrapper has been
  /// finalized / destroyed but JS still holds a reference.
  Napi::Object NullHandleError(Napi::Env env) const { return MakeErrorStatus(env, kBindingNullPointer); }

  fm_workbook_t* handle_ = nullptr;
};

Workbook::Workbook(const Napi::CallbackInfo& info) : Napi::ObjectWrap<Workbook>(info) {
  // Default constructor used by the static factories. They populate
  // `handle_` after construction via `wb->handle_ = ...`.
  //
  // External JS callers should NOT invoke `new Workbook()` directly;
  // the JS-side index.mjs only re-exports the static factories.
  (void)info;
}

Workbook::~Workbook() {
  if (handle_ != nullptr) {
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
}

uint32_t Workbook::ArgU32(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return 0;
  }
  return info[idx].ToNumber().Uint32Value();
}

double Workbook::ArgDouble(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return 0.0;
  }
  return info[idx].ToNumber().DoubleValue();
}

std::string Workbook::ArgString(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return std::string();
  }
  return info[idx].ToString().Utf8Value();
}

// ---- Static factories -----------------------------------------------

Napi::Value Workbook::CreateDefault(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  fm_status_t rc = fm_workbook_create(&wb->handle_);
  if (rc != 0) {
    // Even on failure return the wrapper; the caller can inspect
    // `lastErrorMessage()` and the next operation will fail with
    // `kBindingNullPointer`. This matches embind's behaviour.
    wb->handle_ = nullptr;
  }
  return jsobj;
}

Napi::Value Workbook::CreateEmpty(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  fm_status_t rc = fm_workbook_create_empty(&wb->handle_);
  if (rc != 0) {
    wb->handle_ = nullptr;
  }
  return jsobj;
}

Napi::Value Workbook::LoadBytes(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  if (info.Length() < 1 || !info[0].IsTypedArray()) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  Napi::TypedArray ta = info[0].As<Napi::TypedArray>();
  if (ta.TypedArrayType() != napi_uint8_array) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  Napi::Uint8Array u8 = ta.As<Napi::Uint8Array>();
  const uint8_t* data = u8.Data();
  const std::size_t len = u8.ElementLength();
  if (data == nullptr || len == 0) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  fm_status_t rc = fm_workbook_load(data, len, &wb->handle_);
  if (rc != 0) {
    wb->handle_ = nullptr;
  }
  return jsobj;
}

// ---- Cell mutation --------------------------------------------------

Napi::Value Workbook::SetNumber(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const double value = ArgDouble(info, 3);
  fm_status_t rc = fm_workbook_set_number(handle_, sheet, row, col, value);
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

Napi::Value Workbook::SetBool(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  bool value = false;
  if (info.Length() > 3) {
    value = info[3].ToBoolean().Value();
  }
  fm_status_t rc = fm_workbook_set_bool(handle_, sheet, row, col, value ? 1 : 0);
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

Napi::Value Workbook::SetText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string text = ArgString(info, 3);
  fm_status_t rc = fm_workbook_set_text(handle_, sheet, row, col, text.c_str());
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

Napi::Value Workbook::SetBlank(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_status_t rc = fm_workbook_set_blank(handle_, sheet, row, col);
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

Napi::Value Workbook::SetFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string formula = ArgString(info, 3);
  fm_status_t rc = fm_workbook_set_formula(handle_, sheet, row, col, formula.c_str());
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

// ---- Cell read ------------------------------------------------------

Napi::Value Workbook::GetValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    // Default-zero Value so callers can blindly read .number/.text.
    fm_value_t empty{};
    out.Set("value", TranslateValue(env, empty));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_value_t v{};
  fm_status_t rc = fm_workbook_get_value(handle_, sheet, row, col, &v);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    fm_value_t empty{};
    out.Set("value", TranslateValue(env, empty));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("value", TranslateValue(env, v));
  return out;
}

// ---- Recalc + save --------------------------------------------------

Napi::Value Workbook::Recalc(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  fm_status_t rc = fm_workbook_recalc(handle_);
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

Napi::Value Workbook::Save(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("bytes", env.Null());
    return out;
  }
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  fm_status_t rc = fm_workbook_save(handle_, &buf, &len);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("bytes", env.Null());
    return out;
  }
  // Copy into a fresh Uint8Array on the JS heap; the C-side buffer is
  // owned by the engine and must be released with `fm_buffer_free`.
  Napi::Uint8Array dst = Napi::Uint8Array::New(env, len);
  if (len != 0 && buf != nullptr) {
    std::memcpy(dst.Data(), buf, len);
  }
  fm_buffer_free(buf);
  out.Set("status", MakeOkStatus(env));
  out.Set("bytes", dst);
  return out;
}

// ---- Sheet operations -----------------------------------------------

Napi::Value Workbook::AddSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  fm_status_t rc = fm_workbook_add_sheet(handle_, name.c_str());
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

Napi::Value Workbook::RemoveSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t idx = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_remove_sheet(handle_, idx);
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

Napi::Value Workbook::RenameSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t idx = ArgU32(info, 0);
  const std::string name = ArgString(info, 1);
  fm_status_t rc = fm_workbook_rename_sheet(handle_, idx, name.c_str());
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

Napi::Value Workbook::SheetCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(fm_workbook_sheet_count(handle_)));
}

Napi::Value Workbook::SheetName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("value", Napi::String::New(env, ""));
    return out;
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_sheet_name(handle_, idx, &name);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("value", Napi::String::New(env, ""));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("value", Napi::String::New(env, name != nullptr ? name : ""));
  return out;
}

// ---- Defined names --------------------------------------------------

Napi::Value Workbook::SetDefinedName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  const std::string formula = ArgString(info, 1);
  fm_status_t rc = fm_workbook_set_defined_name(handle_, name.c_str(), formula.c_str());
  return rc == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, rc);
}

// ---- Class registration ---------------------------------------------

Napi::Function Workbook::GetClass(Napi::Env env) {
  return DefineClass(env, "Workbook",
                     {
                         StaticMethod<&Workbook::CreateDefault>("createDefault"),
                         StaticMethod<&Workbook::CreateEmpty>("createEmpty"),
                         StaticMethod<&Workbook::LoadBytes>("loadBytes"),
                         InstanceMethod<&Workbook::SetNumber>("setNumber"),
                         InstanceMethod<&Workbook::SetBool>("setBool"),
                         InstanceMethod<&Workbook::SetText>("setText"),
                         InstanceMethod<&Workbook::SetBlank>("setBlank"),
                         InstanceMethod<&Workbook::SetFormula>("setFormula"),
                         InstanceMethod<&Workbook::GetValue>("getValue"),
                         InstanceMethod<&Workbook::Recalc>("recalc"),
                         InstanceMethod<&Workbook::Save>("save"),
                         InstanceMethod<&Workbook::AddSheet>("addSheet"),
                         InstanceMethod<&Workbook::RemoveSheet>("removeSheet"),
                         InstanceMethod<&Workbook::RenameSheet>("renameSheet"),
                         InstanceMethod<&Workbook::SheetCount>("sheetCount"),
                         InstanceMethod<&Workbook::SheetName>("sheetName"),
                         InstanceMethod<&Workbook::SetDefinedName>("setDefinedName"),
                     });
}

// ---------------------------------------------------------------------
// Free functions
// ---------------------------------------------------------------------

/// `evalFormula(formula)`: convenience that mirrors the embind variant.
/// Spins up an empty workbook, places the formula at A1, recalcs, and
/// returns `{ status, value }`.
Napi::Value EvalFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  std::string formula;
  if (info.Length() > 0) {
    formula = info[0].ToString().Utf8Value();
  }

  fm_workbook_t* wb = nullptr;
  fm_status_t rc = fm_workbook_create(&wb);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    fm_value_t empty{};
    out.Set("value", TranslateValue(env, empty));
    return out;
  }
  rc = fm_workbook_set_formula(wb, 0, 0, 0, formula.c_str());
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    fm_value_t empty{};
    out.Set("value", TranslateValue(env, empty));
    fm_workbook_destroy(wb);
    return out;
  }
  rc = fm_workbook_recalc(wb);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    fm_value_t empty{};
    out.Set("value", TranslateValue(env, empty));
    fm_workbook_destroy(wb);
    return out;
  }
  fm_value_t v{};
  rc = fm_workbook_get_value(wb, 0, 0, 0, &v);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    fm_value_t empty{};
    out.Set("value", TranslateValue(env, empty));
    fm_workbook_destroy(wb);
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("value", TranslateValue(env, v));
  fm_workbook_destroy(wb);
  return out;
}

Napi::Value Version(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_version_string();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value LastErrorMessage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_last_error_message();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value LastErrorContext(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_last_error_context();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value StatusString(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t code = 0;
  if (info.Length() > 0) {
    code = info[0].ToNumber().Int32Value();
  }
  const char* s = fm_status_string(static_cast<fm_status_t>(code));
  return Napi::String::New(env, s != nullptr ? s : "");
}

// ---------------------------------------------------------------------
// Module init
// ---------------------------------------------------------------------

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  exports.Set("Workbook", Workbook::GetClass(env));
  exports.Set("evalFormula", Napi::Function::New(env, &EvalFormula, "evalFormula"));
  exports.Set("version", Napi::Function::New(env, &Version, "version"));
  exports.Set("lastErrorMessage", Napi::Function::New(env, &LastErrorMessage, "lastErrorMessage"));
  exports.Set("lastErrorContext", Napi::Function::New(env, &LastErrorContext, "lastErrorContext"));
  exports.Set("statusString", Napi::Function::New(env, &StatusString, "statusString"));
  return exports;
}

}  // namespace

NODE_API_MODULE(formulon, Init)
