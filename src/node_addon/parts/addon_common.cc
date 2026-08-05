// Shared translation helpers and module-global state for the Node.js
// N-API addon. See `addon_common.h` for the contract.

#include "node_addon/parts/addon_common.h"

namespace formulon_node {

Napi::Object MakeOkStatus(Napi::Env env) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("ok", Napi::Boolean::New(env, true));
  o.Set("status", Napi::Number::New(env, 0));
  o.Set("message", Napi::String::New(env, ""));
  o.Set("context", Napi::String::New(env, ""));
  return o;
}

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

Napi::Object MakeStatus(Napi::Env env, fm_status_t code) {
  return code == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, code);
}

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

Napi::Object MakeFieldResult(Napi::Env env, Napi::Object status, const char* field, Napi::Value value) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set(field, value);
  return out;
}

Napi::Object MakeNumberFieldResult(Napi::Env env, Napi::Object status, const char* field, double value) {
  return MakeFieldResult(env, status, field, Napi::Number::New(env, value));
}

Napi::Object MakeStringFieldResult(Napi::Env env, Napi::Object status, const char* field, const char* value) {
  return MakeFieldResult(env, status, field, Napi::String::New(env, value != nullptr ? value : ""));
}

Napi::Object MakeValueResult(Napi::Env env, Napi::Object status, const fm_value_t& value) {
  return MakeFieldResult(env, status, "value", TranslateValue(env, value));
}

Napi::Object MakeEmptyValueResult(Napi::Env env, Napi::Object status) {
  fm_value_t empty{};
  return MakeValueResult(env, status, empty);
}

Napi::Object MakeIndexResult(Napi::Env env, Napi::Object status, uint32_t index) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set("index", Napi::Number::New(env, index));
  return out;
}

Napi::Object EmptyPivotLayoutResult(Napi::Env env, Napi::Object status) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set("top", Napi::Number::New(env, 0));
  out.Set("left", Napi::Number::New(env, 0));
  out.Set("rows", Napi::Number::New(env, 0));
  out.Set("cols", Napi::Number::New(env, 0));
  out.Set("cells", Napi::Array::New(env));
  return out;
}

Napi::Object TranslatePivotCell(Napi::Env env, const fm_pivot_cell_t& cell) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("row", Napi::Number::New(env, cell.row));
  out.Set("col", Napi::Number::New(env, cell.col));
  out.Set("value", TranslateValue(env, cell.value));
  out.Set("kind", Napi::Number::New(env, static_cast<int32_t>(cell.kind)));
  out.Set("depth", Napi::Number::New(env, cell.depth));
  out.Set("fieldName", Napi::String::New(env, cell.field_name != nullptr ? cell.field_name : ""));
  out.Set("numberFormat", Napi::String::New(env, cell.number_format != nullptr ? cell.number_format : ""));
  return out;
}

Napi::Object TranslateCfColor(Napi::Env env, const fm_cf_color_t& c) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("r", Napi::Number::New(env, static_cast<int32_t>(c.r)));
  o.Set("g", Napi::Number::New(env, static_cast<int32_t>(c.g)));
  o.Set("b", Napi::Number::New(env, static_cast<int32_t>(c.b)));
  o.Set("a", Napi::Number::New(env, static_cast<int32_t>(c.a)));
  return o;
}

Napi::Object TranslateCfMatch(Napi::Env env, const fm_cf_match_t& m) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("kind", Napi::Number::New(env, static_cast<int32_t>(m.kind)));
  o.Set("priority", Napi::Number::New(env, m.priority));
  o.Set("dxfIdEngaged", Napi::Number::New(env, m.dxf_id_engaged));
  o.Set("dxfId", Napi::Number::New(env, m.dxf_id));
  o.Set("color", TranslateCfColor(env, m.color));
  o.Set("barLengthPct", Napi::Number::New(env, m.bar_length_pct));
  o.Set("barAxisPositionPct", Napi::Number::New(env, m.bar_axis_position_pct));
  o.Set("barIsNegative", Napi::Number::New(env, m.bar_is_negative));
  o.Set("barFill", TranslateCfColor(env, m.bar_fill));
  o.Set("barBorderEngaged", Napi::Number::New(env, m.bar_border_engaged));
  o.Set("barBorder", TranslateCfColor(env, m.bar_border));
  o.Set("barGradient", Napi::Number::New(env, m.bar_gradient));
  o.Set("iconSetName", Napi::Number::New(env, m.icon_set_name));
  o.Set("iconIndex", Napi::Number::New(env, static_cast<int32_t>(m.icon_index)));
  return o;
}

int32_t SpecPullInt32(const Napi::Object& spec, const char* key, int32_t dflt) {
  if (!spec.Has(key)) {
    return dflt;
  }
  Napi::Value v = spec.Get(key);
  if (v.IsUndefined() || v.IsNull()) {
    return dflt;
  }
  return v.As<Napi::Number>().Int32Value();
}

uint32_t SpecPullU32(const Napi::Object& spec, const char* key, uint32_t dflt) {
  if (!spec.Has(key)) {
    return dflt;
  }
  Napi::Value v = spec.Get(key);
  if (v.IsUndefined() || v.IsNull()) {
    return dflt;
  }
  return v.As<Napi::Number>().Uint32Value();
}

double SpecPullDouble(const Napi::Object& spec, const char* key, double dflt) {
  if (!spec.Has(key)) {
    return dflt;
  }
  Napi::Value v = spec.Get(key);
  if (v.IsUndefined() || v.IsNull()) {
    return dflt;
  }
  return v.As<Napi::Number>().DoubleValue();
}

bool SpecPullBool(const Napi::Object& spec, const char* key, bool dflt) {
  if (!spec.Has(key)) {
    return dflt;
  }
  Napi::Value v = spec.Get(key);
  if (v.IsUndefined() || v.IsNull()) {
    return dflt;
  }
  return v.ToBoolean().Value();
}

bool SpecHas(const Napi::Object& spec, const char* key) {
  if (!spec.Has(key)) {
    return false;
  }
  Napi::Value v = spec.Get(key);
  return !v.IsUndefined() && !v.IsNull();
}

std::vector<uint32_t> ReadU32Array(const Napi::CallbackInfo& info, size_t idx) {
  std::vector<uint32_t> out;
  if (idx >= info.Length()) {
    return out;
  }
  Napi::Value v = info[idx];
  if (!v.IsArray()) {
    return out;
  }
  Napi::Array arr = v.As<Napi::Array>();
  const uint32_t len = arr.Length();
  out.reserve(len);
  for (uint32_t i = 0; i < len; ++i) {
    out.push_back(arr.Get(i).As<Napi::Number>().Uint32Value());
  }
  return out;
}

void BuildDataFieldSpec(const Napi::Object& spec, fm_pivot_data_field_spec_t& out, std::string& name_buf,
                        std::string& nfmt_buf, bool& has_nfmt) {
  name_buf = spec.Get("name").ToString().Utf8Value();
  has_nfmt = SpecHas(spec, "numberFormat");
  nfmt_buf = has_nfmt ? spec.Get("numberFormat").ToString().Utf8Value() : std::string();
  out.name = name_buf.c_str();
  out.field_index = SpecPullU32(spec, "fieldIndex", 0U);
  out.aggregation = static_cast<fm_pivot_aggregation_t>(SpecPullU32(spec, "aggregation", 0U));
  out.number_format = has_nfmt ? nfmt_buf.c_str() : nullptr;
  out.show_as = static_cast<fm_pivot_show_as_t>(SpecPullU32(spec, "showAs", 0U));
  out.show_as_base_field = SpecPullInt32(spec, "showAsBaseField", -1);
  out.show_as_base_item = SpecPullInt32(spec, "showAsBaseItem", -1);
}

}  // namespace formulon_node
