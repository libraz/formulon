// Reference text: `parseReference` / `formatReference`, the A1 / R1C1
// reference grammar `INDIRECT` reads, exposed to hosts.

#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

constexpr const char* kDestroyedMessage = "the workbook handle was already destroyed";

/// Throws the C ABI's last error as a RangeError (an out-of-grid or unordered argument).
Napi::Value ThrowLastError(Napi::Env env) {
  Napi::RangeError::New(env, fm_get_last_error_message()).ThrowAsJavaScriptException();
  return env.Undefined();
}

/// Reads the optional `options` object at `info[idx]`; false after throwing for a non-object.
bool OptionsArg(const Napi::CallbackInfo& info, size_t idx, const char* api, Napi::Object* out) {
  const Napi::Env env = info.Env();
  if (info.Length() <= idx || info[idx].IsUndefined() || info[idx].IsNull()) {
    *out = Napi::Object::New(env);
    return true;
  }
  if (info[idx].IsArray() || !ObjectArg(info, idx, out)) {
    Napi::TypeError::New(env, std::string(api) + " expects an options object").ThrowAsJavaScriptException();
    return false;
  }
  return true;
}

/// Reads a required uint32 property; a missing one is a TypeError.
uint32_t RequiredU32(CheckedSpecReader& reader, const Napi::Object& owner, const char* key, const char* api) {
  bool present = false;
  const uint32_t v = reader.U32(owner, key, 0U, &present);
  if (reader.ok() && !present) {
    Napi::TypeError::New(owner.Env(), std::string(api) + ": `" + key + "` is required").ThrowAsJavaScriptException();
  }
  return v;
}

}  // namespace

Napi::Value Workbook::ParseReference(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    Napi::TypeError::New(env, kDestroyedMessage).ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Object options;
  if (!OptionsArg(info, 1, "parseReference", &options)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  const bool r1c1 = reader.Bool(options, "r1c1", false);
  bool has_row = false;
  bool has_col = false;
  const uint32_t base_row = reader.U32(options, "baseRow", 0U, &has_row);
  const uint32_t base_col = reader.U32(options, "baseCol", 0U, &has_col);
  if (!reader.ok()) {
    return env.Undefined();
  }
  if (has_row != has_col) {
    Napi::TypeError::New(env, "parseReference: `baseRow` and `baseCol` must be given together")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const std::string text = ArgString(info, 0);
  fm_parsed_reference_t ref{};
  const fm_status_t rc = fm_workbook_parse_reference(handle_, text.c_str(), r1c1 ? 1 : 0, has_row ? 1 : 0, base_row,
                                                     base_col, &ref);
  if (rc != 0) {
    return ThrowLastError(env);
  }
  if (ref.valid == 0) {
    return env.Null();
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("sheetName", ref.has_sheet != 0 ? JsString(env, ref.sheet_name) : env.Null());
  out.Set("row0", Napi::Number::New(env, ref.row0));
  out.Set("col0", Napi::Number::New(env, ref.col0));
  out.Set("row1", Napi::Number::New(env, ref.row1));
  out.Set("col1", Napi::Number::New(env, ref.col1));
  return out;
}

Napi::Value Workbook::FormatReference(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    Napi::TypeError::New(env, kDestroyedMessage).ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Object options;
  if (!OptionsArg(info, 1, "formatReference", &options)) {
    return env.Undefined();
  }
  const Napi::Object range = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  std::string sheet_name;
  const bool has_sheet = reader.String(range, "sheetName", &sheet_name);
  const uint32_t row0 = RequiredU32(reader, range, "row0", "formatReference");
  const uint32_t col0 = reader.ok() ? RequiredU32(reader, range, "col0", "formatReference") : 0U;
  const uint32_t row1 = reader.ok() ? RequiredU32(reader, range, "row1", "formatReference") : 0U;
  const uint32_t col1 = reader.ok() ? RequiredU32(reader, range, "col1", "formatReference") : 0U;
  const bool absolute = reader.ok() && reader.Bool(options, "absolute", false);
  if (!reader.ok()) {
    return env.Undefined();
  }
  const char* text = nullptr;
  const fm_status_t rc = fm_workbook_format_reference(handle_, has_sheet ? sheet_name.c_str() : nullptr, row0, col0,
                                                      row1, col1, absolute ? 1 : 0, &text);
  if (rc != 0) {
    return ThrowLastError(env);
  }
  return JsString(env, text);
}

}  // namespace formulon_node
