// Shared translation helpers and module-global state for the Node.js
// N-API addon. See `addon_common.h` for the contract.

#include "node_addon/parts/addon_common.h"

#include <cmath>
#include <limits>

namespace formulon_node {

namespace {

// Builds `{ ok, status, message, context }`; NULL strings become "".
Napi::Object MakeStatusEnvelope(Napi::Env env, bool ok, fm_status_t code, const char* message, const char* context) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("ok", Napi::Boolean::New(env, ok));
  o.Set("status", Napi::Number::New(env, static_cast<int32_t>(code)));
  o.Set("message", Napi::String::New(env, message != nullptr ? message : ""));
  o.Set("context", Napi::String::New(env, context != nullptr ? context : ""));
  return o;
}

}  // namespace

// FORMULON-ALLOW: N-API uses JavaScript exceptions for argument validation.
void CheckedSpecReader::ObservePending() {
  if (!failed_ && env_.IsExceptionPending()) {
    failed_ = true;
  }
}

void CheckedSpecReader::ReportType(const char* key, const char* expected) {
  if (failed_ || env_.IsExceptionPending()) {
    failed_ = true;
    return;
  }
  Napi::TypeError::New(env_, std::string(key) + " must be a " + expected).ThrowAsJavaScriptException();
  failed_ = true;
}

void CheckedSpecReader::ReportRange(const char* key, const char* range_name) {
  if (failed_ || env_.IsExceptionPending()) {
    failed_ = true;
    return;
  }
  Napi::RangeError::New(env_, std::string(key) + " is outside " + range_name + " range").ThrowAsJavaScriptException();
  failed_ = true;
}

bool CheckedSpecReader::ReadOptional(const Napi::Object& owner, const char* key, Napi::Value* out, bool* present) {
  if (present != nullptr) {
    *present = false;
  }
  if (!ok()) {
    return false;
  }
  const bool has_property = owner.Has(key);
  ObservePending();
  if (!ok() || !has_property) {
    return false;
  }
  *out = owner.Get(key);
  ObservePending();
  if (!ok() || out->IsUndefined() || out->IsNull()) {
    return false;
  }
  if (present != nullptr) {
    *present = true;
  }
  return true;
}

bool CheckedSpecReader::Value(const Napi::Value& value, const char* key, Napi::Value* out) {
  if (!ok() || value.IsUndefined() || value.IsNull()) {
    return false;
  }
  *out = value;
  ObservePending();
  if (!ok()) {
    return false;
  }
  (void)key;
  return true;
}

bool CheckedSpecReader::ArrayElement(const Napi::Array& owner, uint32_t index, Napi::Value* out) {
  if (!ok()) {
    return false;
  }
  *out = owner.Get(index);
  ObservePending();
  if (!ok()) {
    return false;
  }
  return Value(*out, "array element", out);
}

bool CheckedSpecReader::ReadNumber(const Napi::Object& owner, const char* key, double* out, bool* present) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value, present)) {
    return ok();
  }
  return ReadNumber(value, key, out);
}

bool CheckedSpecReader::ReadNumber(const Napi::Value& value, const char* key, double* out) {
  Napi::Value present;
  if (!Value(value, key, &present)) {
    return ok();
  }
  if (!present.IsNumber()) {
    ReportType(key, "number");
    return false;
  }
  *out = present.As<Napi::Number>().DoubleValue();
  ObservePending();
  return ok();
}

bool CheckedSpecReader::ReadInteger(const Napi::Object& owner, const char* key, double* out, double min_value,
                                    double max_value, const char* range_name, bool* present) {
  if (!ReadNumber(owner, key, out, present)) {
    return false;
  }
  if (!std::isfinite(*out) || std::trunc(*out) != *out || *out < min_value || *out > max_value) {
    ReportRange(key, range_name);
    return false;
  }
  return true;
}

bool CheckedSpecReader::String(const Napi::Object& owner, const char* key, std::string* out) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value)) {
    out->clear();
    return false;
  }
  return String(value, key, out);
}

bool CheckedSpecReader::String(const Napi::Value& value, const char* key, std::string* out) {
  Napi::Value present;
  if (!Value(value, key, &present)) {
    out->clear();
    return false;
  }
  // Same acceptance as the WASM surface: a string or a one-byte buffer, never
  // a coerced value.
  const bool byte_view = present.IsTypedArray() && [&present] {
    const napi_typedarray_type type = present.As<Napi::TypedArray>().TypedArrayType();
    return type == napi_int8_array || type == napi_uint8_array || type == napi_uint8_clamped_array;
  }();
  if (present.IsString()) {
    *out = present.As<Napi::String>().Utf8Value();
  } else if (present.IsArrayBuffer()) {
    Napi::ArrayBuffer buffer = present.As<Napi::ArrayBuffer>();
    out->assign(static_cast<const char*>(buffer.Data()), buffer.ByteLength());
  } else if (byte_view) {
    Napi::TypedArray view = present.As<Napi::TypedArray>();
    const auto* data = static_cast<const char*>(view.ArrayBuffer().Data()) + view.ByteOffset();
    out->assign(data, view.ByteLength());
  } else {
    out->clear();
    ReportType(key, "string or one-byte buffer");
    return false;
  }
  ObservePending();
  return ok();
}

bool CheckedSpecReader::Object(const Napi::Object& owner, const char* key, Napi::Object* out) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value)) {
    return false;
  }
  return Object(value, key, out);
}

bool CheckedSpecReader::Object(const Napi::Value& value, const char* key, Napi::Object* out) {
  Napi::Value present;
  if (!Value(value, key, &present)) {
    return false;
  }
  if (!present.IsObject()) {
    ReportType(key, "object");
    return false;
  }
  *out = present.As<Napi::Object>();
  ObservePending();
  return ok();
}

bool CheckedSpecReader::Array(const Napi::Object& owner, const char* key, Napi::Array* out) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value)) {
    return false;
  }
  return Array(value, key, out);
}

bool CheckedSpecReader::Array(const Napi::Value& value, const char* key, Napi::Array* out) {
  Napi::Value present;
  if (!Value(value, key, &present)) {
    return false;
  }
  if (!present.IsArray()) {
    ReportType(key, "array");
    return false;
  }
  *out = present.As<Napi::Array>();
  ObservePending();
  return ok();
}

bool CheckedSpecReader::Bool(const Napi::Object& owner, const char* key, bool dflt, bool* present) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value, present)) {
    return dflt;
  }
  return Bool(value, key, dflt);
}

bool CheckedSpecReader::Bool(const Napi::Value& value, const char* key, bool dflt) {
  Napi::Value present;
  if (!Value(value, key, &present)) {
    return dflt;
  }
  const bool result = present.ToBoolean().Value();
  ObservePending();
  return result;
}

double CheckedSpecReader::Double(const Napi::Object& owner, const char* key, double dflt, bool* present) {
  double value = dflt;
  if (!ReadNumber(owner, key, &value, present)) {
    return dflt;
  }
  return value;
}

double CheckedSpecReader::Double(const Napi::Value& value, const char* key, double dflt) {
  double result = dflt;
  if (!ReadNumber(value, key, &result)) {
    return dflt;
  }
  return result;
}

uint8_t CheckedSpecReader::U8(const Napi::Object& owner, const char* key, uint8_t dflt, bool* present) {
  double value = static_cast<double>(dflt);
  if (!ReadInteger(owner, key, &value, 0.0, static_cast<double>(std::numeric_limits<uint8_t>::max()), "uint8",
                   present)) {
    return dflt;
  }
  return static_cast<uint8_t>(value);
}

uint8_t CheckedSpecReader::U8(const Napi::Value& value, const char* key, uint8_t dflt) {
  double result = static_cast<double>(dflt);
  if (!ReadNumber(value, key, &result)) {
    return dflt;
  }
  if (!std::isfinite(result) || std::trunc(result) != result || result < 0.0 ||
      result > static_cast<double>(std::numeric_limits<uint8_t>::max())) {
    ReportRange(key, "uint8");
    return dflt;
  }
  return static_cast<uint8_t>(result);
}

uint16_t CheckedSpecReader::U16(const Napi::Object& owner, const char* key, uint16_t dflt, bool* present) {
  double value = static_cast<double>(dflt);
  if (!ReadInteger(owner, key, &value, 0.0, static_cast<double>(std::numeric_limits<uint16_t>::max()), "uint16",
                   present)) {
    return dflt;
  }
  return static_cast<uint16_t>(value);
}

uint16_t CheckedSpecReader::U16(const Napi::Value& value, const char* key, uint16_t dflt) {
  double result = static_cast<double>(dflt);
  if (!ReadNumber(value, key, &result)) {
    return dflt;
  }
  if (!std::isfinite(result) || std::trunc(result) != result || result < 0.0 ||
      result > static_cast<double>(std::numeric_limits<uint16_t>::max())) {
    ReportRange(key, "uint16");
    return dflt;
  }
  return static_cast<uint16_t>(result);
}

uint32_t CheckedSpecReader::U32(const Napi::Object& owner, const char* key, uint32_t dflt, bool* present) {
  double value = static_cast<double>(dflt);
  if (!ReadInteger(owner, key, &value, 0.0, static_cast<double>(std::numeric_limits<uint32_t>::max()), "uint32",
                   present)) {
    return dflt;
  }
  return static_cast<uint32_t>(value);
}

uint32_t CheckedSpecReader::U32(const Napi::Value& value, const char* key, uint32_t dflt) {
  double result = static_cast<double>(dflt);
  if (!ReadNumber(value, key, &result)) {
    return dflt;
  }
  if (!std::isfinite(result) || std::trunc(result) != result || result < 0.0 ||
      result > static_cast<double>(std::numeric_limits<uint32_t>::max())) {
    ReportRange(key, "uint32");
    return dflt;
  }
  return static_cast<uint32_t>(result);
}

int32_t CheckedSpecReader::I32(const Napi::Object& owner, const char* key, int32_t dflt, bool* present) {
  double value = static_cast<double>(dflt);
  if (!ReadInteger(owner, key, &value, static_cast<double>(std::numeric_limits<int32_t>::min()),
                   static_cast<double>(std::numeric_limits<int32_t>::max()), "int32", present)) {
    return dflt;
  }
  return static_cast<int32_t>(value);
}

int32_t CheckedSpecReader::I32(const Napi::Value& value, const char* key, int32_t dflt) {
  double result = static_cast<double>(dflt);
  if (!ReadNumber(value, key, &result)) {
    return dflt;
  }
  if (!std::isfinite(result) || std::trunc(result) != result ||
      result < static_cast<double>(std::numeric_limits<int32_t>::min()) ||
      result > static_cast<double>(std::numeric_limits<int32_t>::max())) {
    ReportRange(key, "int32");
    return dflt;
  }
  return static_cast<int32_t>(result);
}

int64_t CheckedSpecReader::I64(const Napi::Object& owner, const char* key, int64_t dflt, bool* present) {
  double value = static_cast<double>(dflt);
  if (!ReadNumber(owner, key, &value, present)) {
    return dflt;
  }
  if (!std::isfinite(value) || std::trunc(value) != value || value < -0x1p63 || value >= 0x1p63) {
    ReportRange(key, "int64");
    return dflt;
  }
  if (value == -0x1p63) {
    return std::numeric_limits<int64_t>::min();
  }
  return static_cast<int64_t>(value);
}

int64_t CheckedSpecReader::I64(const Napi::Value& value, const char* key, int64_t dflt) {
  Napi::Value present;
  if (!Value(value, key, &present)) {
    return dflt;
  }
  if (!present.IsNumber()) {
    ReportType(key, "number");
    return dflt;
  }
  const double number = present.As<Napi::Number>().DoubleValue();
  ObservePending();
  if (!ok()) {
    return dflt;
  }
  if (!std::isfinite(number) || std::trunc(number) != number || number < -0x1p63 || number >= 0x1p63) {
    ReportRange(key, "int64");
    return dflt;
  }
  if (number == -0x1p63) {
    return std::numeric_limits<int64_t>::min();
  }
  return static_cast<int64_t>(number);
}

Napi::Object MakeOkStatus(Napi::Env env) {
  return MakeStatusEnvelope(env, true, 0, "", "");
}

Napi::Object MakeBindingError(Napi::Env env, fm_status_t code, const char* message) {
  return MakeStatusEnvelope(env, false, code, message, "");
}

Napi::Object MakeBindingArgumentError(Napi::Env env, const char* message) {
  return MakeBindingError(env, kBindingNullPointer, message);
}

Napi::Object MakeCallbackThrewStatus(Napi::Env env) {
  return MakeStatusEnvelope(env, false, kBindingCallbackException,
                            "the iterative progress callback threw; the solve was aborted",
                            "Workbook.setIterativeProgress");
}

Napi::Object MakeErrorStatus(Napi::Env env, fm_status_t code) {
  if (code == kBindingInvalidHandle) {
    // No C-ABI call ran, so the thread-local diagnostics still describe
    // whatever call came before. Synthesise the message here instead.
    return MakeBindingError(env, code, "the workbook handle was already destroyed");
  }
  return MakeStatusEnvelope(env, false, code, fm_last_error_message(), fm_last_error_context());
}

Napi::Object MakeStatus(Napi::Env env, fm_status_t code) {
  return code == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, code);
}

Napi::Array FinishListResult(Napi::Env env, Napi::Array items, fm_status_t code) {
  items.Set("status", MakeStatus(env, code));
  return items;
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

Napi::Object MakeNumberFieldResult(Napi::Env env, fm_status_t code, const char* field, double value) {
  return MakeNumberFieldResult(env, MakeStatus(env, code), field, code == 0 ? value : 0.0);
}

Napi::Object MakeStringFieldResult(Napi::Env env, fm_status_t code, const char* field, const char* value) {
  return MakeStringFieldResult(env, MakeStatus(env, code), field, code == 0 ? value : nullptr);
}

Napi::Object MakeNumberResult(Napi::Env env, fm_status_t code, double value) {
  return MakeNumberFieldResult(env, code, "value", value);
}

Napi::Object MakeStringResult(Napi::Env env, fm_status_t code, const char* value) {
  return MakeStringFieldResult(env, code, "value", value);
}

Napi::Object MakeValueResult(Napi::Env env, Napi::Object status, const fm_value_t& value) {
  return MakeFieldResult(env, status, "value", TranslateValue(env, value));
}

Napi::Object MakeEmptyValueResult(Napi::Env env, Napi::Object status) {
  fm_value_t empty{};
  return MakeValueResult(env, status, empty);
}

Napi::Object MakeValueResult(Napi::Env env, fm_status_t code, const fm_value_t& value) {
  return code == 0 ? MakeValueResult(env, MakeOkStatus(env), value)
                   : MakeEmptyValueResult(env, MakeErrorStatus(env, code));
}

Napi::Object MakeIndexResult(Napi::Env env, Napi::Object status, uint32_t index) {
  return MakeNumberFieldResult(env, status, "index", index);
}

Napi::Object MakeIndexResult(Napi::Env env, fm_status_t code, uint32_t index) {
  return MakeNumberFieldResult(env, code, "index", index);
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
  o.Set("barDirection", Napi::Number::New(env, static_cast<int32_t>(m.bar_direction)));
  o.Set("iconSetName", Napi::Number::New(env, m.icon_set_name));
  o.Set("iconIndex", Napi::Number::New(env, static_cast<int32_t>(m.icon_index)));
  return o;
}

void BuildDataFieldSpec(CheckedSpecReader& reader, const Napi::Object& spec, fm_pivot_data_field_spec_t& out,
                        std::string& name_buf, std::string& nfmt_buf, bool& has_nfmt) {
  // `name` is required; an omitted key passes NULL through rather than
  // the coerced literal string "undefined" `.ToString()` would otherwise
  // produce, so the C ABI's own kBindingNullPointer check is what rejects
  // the call. See the `sourceName` comment in `pivot_table.cc`'s
  // `PivotFieldAdd`.
  const bool has_name = reader.String(spec, "name", &name_buf);
  has_nfmt = reader.String(spec, "numberFormat", &nfmt_buf);
  out.name = has_name ? name_buf.c_str() : nullptr;
  out.field_index = reader.U32(spec, "fieldIndex", 0U);
  out.aggregation = static_cast<fm_pivot_aggregation_t>(reader.U32(spec, "aggregation", 0U));
  out.number_format = has_nfmt ? nfmt_buf.c_str() : nullptr;
  out.show_as = static_cast<fm_pivot_show_as_t>(reader.U32(spec, "showAs", 0U));
  out.show_as_base_field = reader.I32(spec, "showAsBaseField", -1);
  out.show_as_base_item = reader.I32(spec, "showAsBaseItem", -1);
}

// Reads a `{ firstRow, lastRow, firstCol, lastCol }` range object.
bool ReadMergeRange(CheckedSpecReader& reader, const Napi::Object& range, fm_merge_range* out) {
  out->first_row = reader.U32(range, "firstRow", 0U);
  out->last_row = reader.U32(range, "lastRow", 0U);
  out->first_col = reader.U32(range, "firstCol", 0U);
  out->last_col = reader.U32(range, "lastCol", 0U);
  return reader.ok();
}

// Reads the range at `info[idx]`; a missing/nullish argument reads as all
// zeros, while a supplied non-object is rejected by the reader.
bool MergeRangeArg(CheckedSpecReader& reader, const Napi::CallbackInfo& info, size_t idx, fm_merge_range* out) {
  *out = fm_merge_range{};
  if (info.Length() <= idx) {
    return true;
  }
  Napi::Value value = info[idx];
  Napi::Value present;
  if (!reader.Value(value, "range", &present)) {
    return reader.ok();
  }
  Napi::Object object;
  if (!reader.Object(present, "range", &object)) {
    return false;
  }
  return ReadMergeRange(reader, object, out);
}

}  // namespace formulon_node
