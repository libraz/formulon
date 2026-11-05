// Reference text: `parseReference` / `formatReference`, the A1 / R1C1
// reference grammar `INDIRECT` reads, exposed to hosts.

#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

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

}  // namespace

Napi::Value Workbook::ParseReference(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
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
  const fm_status_t rc =
      fm_workbook_parse_reference(handle_, text.c_str(), r1c1 ? 1 : 0, has_row ? 1 : 0, base_row, base_col, &ref);
  if (rc != 0 || ref.valid == 0) {
    return MakeOptionalResult(env, rc, 0, env.Null());
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("sheetName", ref.has_sheet != 0 ? JsString(env, ref.sheet_name) : env.Null());
  out.Set("rect", RangeToJs(env, ref.rect));
  return MakeOptionalResult(env, rc, 1, out);
}

Napi::Value Workbook::FormatReference(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "value", "");
  }
  Napi::Object options;
  if (!OptionsArg(info, 1, "formatReference", &options)) {
    return env.Undefined();
  }
  const Napi::Object range = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  std::string sheet_name;
  const bool has_sheet = reader.String(range, "sheetName", &sheet_name);
  fm_cell_rect_t rect{};
  Napi::Object rect_object;
  if (!reader.Object(range, "rect", &rect_object)) {
    if (reader.ok()) {
      Napi::TypeError::New(env, "formatReference: `rect` is required").ThrowAsJavaScriptException();
    }
    return env.Undefined();
  }
  if (!ReadCellRect(reader, rect_object, &rect)) {
    return env.Undefined();
  }
  const bool absolute = reader.ok() && reader.Bool(options, "absolute", false);
  if (!reader.ok()) {
    return env.Undefined();
  }
  const char* text = nullptr;
  const fm_status_t rc =
      fm_workbook_format_reference(handle_, has_sheet ? sheet_name.c_str() : nullptr, &rect, absolute ? 1 : 0, &text);
  return MakeStringResult(env, rc, text);
}

}  // namespace formulon_node
