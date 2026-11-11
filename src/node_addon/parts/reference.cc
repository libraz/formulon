// Reference text: `parseReference` / `formatReference`, the A1 / R1C1
// reference grammar `INDIRECT` reads, exposed to hosts, the sheet lookup by
// name, `getSheetIndex`, and the `{ r1c1, baseRow, baseCol }` options reader
// `localizeFormula` / `canonicalizeFormula` share with `parseReference`.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

bool OptionsArg(const Napi::CallbackInfo& info, size_t idx, Napi::Object* out) {
  const Napi::Env env = info.Env();
  if (info.Length() <= idx || info[idx].IsUndefined() || info[idx].IsNull()) {
    *out = Napi::Object::New(env);
    return true;
  }
  if (info[idx].IsArray() || !ObjectArg(info, idx, out)) {
    ThrowArgumentTypeError(env, "options", "must be an object");
    return false;
  }
  return true;
}

bool NotationOptionsArg(const Napi::CallbackInfo& info, size_t idx, NotationOptions* out) {
  const Napi::Env env = info.Env();
  Napi::Object options;
  if (!OptionsArg(info, idx, &options)) {
    return false;
  }
  CheckedSpecReader reader(env);
  const bool r1c1 = reader.Bool(options, "r1c1", false);
  bool has_row = false;
  bool has_col = false;
  const uint32_t base_row = reader.U32(options, "baseRow", 0U, &has_row);
  const uint32_t base_col = reader.U32(options, "baseCol", 0U, &has_col);
  if (!reader.ok()) {
    return false;
  }
  if (has_row != has_col) {
    ThrowArgumentTypeError(env, "options.baseRow and options.baseCol", "must be given together");
    return false;
  }
  out->r1c1 = r1c1;
  out->has_base = has_row;
  out->base_row = base_row;
  out->base_col = base_col;
  return true;
}

Napi::Value Workbook::ParseReference(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  NotationOptions notation;
  if (!NotationOptionsArg(info, 1, &notation)) {
    return env.Undefined();
  }
  const std::string text = ArgString(info, 0);
  fm_parsed_reference_info_t ref = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  const fm_parse_reference_spec_t spec = NotationSpec<fm_parse_reference_spec_t>(notation);
  const fm_status_t rc = fm_workbook_parse_reference(handle_, text.c_str(), &spec, &ref);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("sheetName", JsStringOrNull(env, ref.sheet_name));
  out.Set("rect", RangeToJs(env, ref.rect));
  return out;
}

Napi::Value Workbook::FormatReference(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object options;
  if (!OptionsArg(info, 1, &options)) {
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
      ThrowArgumentTypeError(env, "range.rect", "is required");
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
  return ValueOrThrow(env, rc, JsString(env, text));
}

Napi::Value Workbook::GetSheetIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string name = ArgString(info, 0);
  std::size_t index = 0;
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_sheet_index(handle_, name.c_str(), &present, &index);
  return OptionalOrNull(env, rc, present, Napi::Number::New(env, static_cast<double>(index)));
}

}  // namespace formulon_node
