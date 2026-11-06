// Shared translation helpers and module-global state for the Node.js
// N-API addon. See `addon_common.h` for the contract.

#include "node_addon/parts/addon_common.h"

#include <cmath>
#include <cstring>
#include <limits>

namespace formulon_node {

Napi::Uint8Array CopyBytes(Napi::Env env, const uint8_t* bytes, std::size_t len) {
  Napi::Uint8Array copy = Napi::Uint8Array::New(env, len);
  if (len != 0 && bytes != nullptr) {
    std::memcpy(copy.Data(), bytes, len);
  }
  return copy;
}

Napi::Uint8Array TakeSavedBytes(Napi::Env env, uint8_t* bytes, std::size_t len) {
  Napi::Uint8Array copy = CopyBytes(env, bytes, len);
  fm_buffer_destroy(bytes);
  return copy;
}

// FORMULON-ALLOW: N-API uses JavaScript exceptions for argument validation.
void CheckedSpecReader::ObservePending() {
  if (!failed_ && env_.IsExceptionPending()) {
    failed_ = true;
  }
}

template <typename ErrorType>
void CheckedSpecReader::ReportFailure(const std::string& message) {
  if (failed_ || env_.IsExceptionPending()) {
    failed_ = true;
    return;
  }
  ErrorType::New(env_, message).ThrowAsJavaScriptException();
  failed_ = true;
}

void CheckedSpecReader::ReportType(const char* key, const char* expected) {
  ReportFailure<Napi::TypeError>(std::string(key) + " must be a " + expected);
}

void CheckedSpecReader::ReportRange(const char* key, const char* range_name) {
  ReportFailure<Napi::RangeError>(std::string(key) + " is outside " + range_name + " range");
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
  Napi::Value value;
  if (!ReadOptional(owner, key, &value, present)) {
    return ok();
  }
  return ReadInteger(value, key, out, min_value, max_value, range_name);
}

bool CheckedSpecReader::ReadInteger(const Napi::Value& value, const char* key, double* out, double min_value,
                                    double max_value, const char* range_name) {
  if (!ReadNumber(value, key, out)) {
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

template <typename T>
T CheckedSpecReader::ReadIntegral(const Napi::Object& owner, const char* key, T dflt, const char* range_name,
                                  bool* present) {
  double value = static_cast<double>(dflt);
  if (!ReadInteger(owner, key, &value, static_cast<double>(std::numeric_limits<T>::min()),
                   static_cast<double>(std::numeric_limits<T>::max()), range_name, present)) {
    return dflt;
  }
  return static_cast<T>(value);
}

template <typename T>
T CheckedSpecReader::ReadIntegral(const Napi::Value& value, const char* key, T dflt, const char* range_name) {
  double result = static_cast<double>(dflt);
  if (!ReadInteger(value, key, &result, static_cast<double>(std::numeric_limits<T>::min()),
                   static_cast<double>(std::numeric_limits<T>::max()), range_name)) {
    return dflt;
  }
  return static_cast<T>(result);
}

uint8_t CheckedSpecReader::U8(const Napi::Object& owner, const char* key, uint8_t dflt, bool* present) {
  return ReadIntegral<uint8_t>(owner, key, dflt, "uint8", present);
}

uint8_t CheckedSpecReader::U8(const Napi::Value& value, const char* key, uint8_t dflt) {
  return ReadIntegral<uint8_t>(value, key, dflt, "uint8");
}

uint16_t CheckedSpecReader::U16(const Napi::Object& owner, const char* key, uint16_t dflt, bool* present) {
  return ReadIntegral<uint16_t>(owner, key, dflt, "uint16", present);
}

uint16_t CheckedSpecReader::U16(const Napi::Value& value, const char* key, uint16_t dflt) {
  return ReadIntegral<uint16_t>(value, key, dflt, "uint16");
}

uint32_t CheckedSpecReader::U32(const Napi::Object& owner, const char* key, uint32_t dflt, bool* present) {
  return ReadIntegral<uint32_t>(owner, key, dflt, "uint32", present);
}

uint32_t CheckedSpecReader::U32(const Napi::Value& value, const char* key, uint32_t dflt) {
  return ReadIntegral<uint32_t>(value, key, dflt, "uint32");
}

int32_t CheckedSpecReader::I32(const Napi::Object& owner, const char* key, int32_t dflt, bool* present) {
  return ReadIntegral<int32_t>(owner, key, dflt, "int32", present);
}

int32_t CheckedSpecReader::I32(const Napi::Value& value, const char* key, int32_t dflt) {
  return ReadIntegral<int32_t>(value, key, dflt, "int32");
}

bool CheckedSpecReader::RequiredObject(const Napi::Object& owner, const char* key, Napi::Object* out) {
  if (Object(owner, key, out)) {
    return true;
  }
  if (ok()) {
    ReportType(key, "object");
  }
  return false;
}

bool CheckedSpecReader::RequiredArray(const Napi::Object& owner, const char* key, Napi::Array* out) {
  if (Array(owner, key, out)) {
    return true;
  }
  if (ok()) {
    ReportType(key, "array");
  }
  return false;
}

std::string CheckedSpecReader::RequiredString(const Napi::Object& owner, const char* key) {
  const Napi::Value value = Get(owner, key);
  if (!ok()) {
    return std::string();
  }
  if (!value.IsString()) {
    ReportType(key, "string");
    return std::string();
  }
  std::string out;
  (void)String(value, key, &out);
  return out;
}

Napi::Value CheckedSpecReader::Get(const Napi::Object& owner, const char* key) {
  if (!ok()) {
    return env_.Undefined();
  }
  const Napi::Value value = owner.Get(key);
  ObservePending();
  return ok() ? value : env_.Undefined();
}

uint32_t CheckedSpecReader::RequiredU32(const Napi::Object& owner, const char* key) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value)) {
    if (ok()) {
      ReportType(key, "whole number");
    }
    return 0U;
  }
  const double number = value.IsNumber() ? value.As<Napi::Number>().DoubleValue() : 0.5;
  if (!std::isfinite(number) || std::trunc(number) != number) {
    ReportType(key, "whole number");
    return 0U;
  }
  if (number < 0.0 || number > kUint32MaxNumber) {
    ReportRange(key, "uint32");
    return 0U;
  }
  return static_cast<uint32_t>(number);
}

int64_t CheckedSpecReader::NarrowI64(double number, const char* key, int64_t dflt) {
  if (!std::isfinite(number) || std::trunc(number) != number || number < -0x1p63 || number >= 0x1p63) {
    ReportRange(key, "int64");
    return dflt;
  }
  if (number == -0x1p63) {
    return std::numeric_limits<int64_t>::min();
  }
  return static_cast<int64_t>(number);
}

int64_t CheckedSpecReader::I64(const Napi::Object& owner, const char* key, int64_t dflt, bool* present) {
  double value = static_cast<double>(dflt);
  if (!ReadNumber(owner, key, &value, present)) {
    return dflt;
  }
  return NarrowI64(value, key, dflt);
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
  return NarrowI64(number, key, dflt);
}

double NumberValue(const Napi::Value& value, double dflt) {
  return value.IsNumber() ? value.As<Napi::Number>().DoubleValue() : dflt;
}

double NumberArg(const Napi::CallbackInfo& info, size_t idx, double dflt) {
  return NumberValue(info[idx], dflt);
}

std::string StringArg(const Napi::CallbackInfo& info, size_t idx) {
  const Napi::Value value = info[idx];
  return value.IsString() ? value.As<Napi::String>().Utf8Value() : std::string();
}

bool ObjectArg(const Napi::CallbackInfo& info, size_t idx, Napi::Object* out) {
  const Napi::Value value = info[idx];
  if (!value.IsObject()) {
    return false;
  }
  *out = value.As<Napi::Object>();
  return true;
}

bool ArrayArg(const Napi::CallbackInfo& info, size_t idx, Napi::Array* out) {
  const Napi::Value value = info[idx];
  if (!value.IsArray()) {
    return false;
  }
  *out = value.As<Napi::Array>();
  return true;
}

bool FunctionArg(const Napi::CallbackInfo& info, size_t idx, Napi::Function* out) {
  const Napi::Value value = info[idx];
  if (!value.IsFunction()) {
    return false;
  }
  *out = value.As<Napi::Function>();
  return true;
}

bool BytesArg(const Napi::CallbackInfo& info, size_t idx, Napi::Uint8Array* out) {
  const Napi::Value value = info[idx];
  if (!value.IsTypedArray() || value.As<Napi::TypedArray>().TypedArrayType() != napi_uint8_array) {
    return false;
  }
  *out = value.As<Napi::Uint8Array>();
  return true;
}

namespace {

// Message of the `TypeError` thrown when a failure is due but the JS loader
// never injected the error class.
constexpr const char* kMissingErrorClassMessage = "FormulonError was not provided to the addon";

// Throws `new FormulonError(code, message, context)`; NULL strings become "".
Napi::Value ThrowFormulonError(Napi::Env env, fm_status_t code, const char* message, const char* context) {
  const Napi::FunctionReference& error_class = GetAddonData(env).error_class;
  if (error_class.IsEmpty()) {
    Napi::TypeError::New(env, kMissingErrorClassMessage).ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const Napi::Object error = error_class.New({
      Napi::Number::New(env, static_cast<double>(code)),
      JsString(env, message),
      JsString(env, context),
  });
  if (!env.IsExceptionPending()) {
    napi_throw(env, error);
  }
  return env.Undefined();
}

// Rejects a malformed input `Value`.
bool RejectValue(Napi::Env env, const char* message) {
  Napi::TypeError::New(env, message).ThrowAsJavaScriptException();
  return false;
}

// Reads `object[key]`; false when the getter threw.
bool ValueField(Napi::Env env, const Napi::Object& object, const char* key, Napi::Value* out) {
  *out = object.Get(key);
  return !env.IsExceptionPending();
}

// True when `value` is an integral Number in [min_value, max_value].
bool IsIntegerIn(const Napi::Value& value, double min_value, double max_value) {
  if (!value.IsNumber()) {
    return false;
  }
  const double number = value.As<Napi::Number>().DoubleValue();
  return std::isfinite(number) && std::trunc(number) == number && number >= min_value && number <= max_value;
}

}  // namespace

AddonData& GetAddonData(Napi::Env env) {
  AddonData* data = env.GetInstanceData<AddonData>();
  if (data == nullptr) {
    data = new AddonData{};
    env.SetInstanceData(data);
  }
  return *data;
}

Napi::Value ThrowStatus(Napi::Env env, fm_status_t rc) {
  if (env.IsExceptionPending()) {
    return env.Undefined();
  }
  return ThrowFormulonError(env, rc, fm_get_last_error_message(), fm_get_last_error_context());
}

Napi::Value ThrowInvalidHandle(Napi::Env env) {
  if (env.IsExceptionPending()) {
    return env.Undefined();
  }
  return ThrowFormulonError(env, kBindingInvalidHandle, "the workbook is disposed", "");
}

Napi::Value ThrowReentrant(Napi::Env env, const char* message) {
  if (env.IsExceptionPending()) {
    return env.Undefined();
  }
  return ThrowFormulonError(env, FM_STATUS_GRAPH_RECALC_REENTRANT, message, "");
}

Napi::Value ValueOrThrow(Napi::Env env, fm_status_t rc, Napi::Value ok_value) {
  return rc == 0 ? ok_value : ThrowStatus(env, rc);
}

Napi::Value OptionalOrNull(Napi::Env env, fm_status_t rc, int32_t present, Napi::Value item) {
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  return present != 0 ? item : env.Null();
}

Napi::Value ValueToJs(Napi::Env env, const fm_value_t& v) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("kind", Napi::Number::New(env, static_cast<int32_t>(v.kind)));
  switch (v.kind) {
    case FM_VAL_NUMBER:
      o.Set("number", Napi::Number::New(env, v.u.number));
      break;
    case FM_VAL_BOOL:
      o.Set("boolean", Napi::Boolean::New(env, v.u.boolean != 0));
      break;
    case FM_VAL_TEXT:
      o.Set("text", JsString(env, v.u.text));
      break;
    case FM_VAL_ERROR:
      o.Set("errorCode", Napi::Number::New(env, v.u.error_code));
      break;
    default:
      break;
  }
  return o;
}

bool JsToValue(Napi::Env env, Napi::Value in, fm_value_t* out, std::string* text) {
  *out = fm_value_t{};
  if (env.IsExceptionPending()) {
    return false;
  }
  if (!in.IsObject()) {
    return RejectValue(env, "value must be an object");
  }
  const Napi::Object object = in.As<Napi::Object>();
  Napi::Value kind;
  if (!ValueField(env, object, "kind", &kind)) {
    return false;
  }
  if (!IsIntegerIn(kind, static_cast<double>(FM_VAL_BLANK), static_cast<double>(FM_VAL_LAMBDA))) {
    return RejectValue(env, "value.kind must be a ValueKind");
  }
  out->kind = static_cast<int32_t>(kind.As<Napi::Number>().DoubleValue());
  Napi::Value field;
  switch (out->kind) {
    case FM_VAL_NUMBER:
      if (!ValueField(env, object, "number", &field)) {
        return false;
      }
      if (!field.IsNumber()) {
        return RejectValue(env, "value.number must be a number");
      }
      out->u.number = field.As<Napi::Number>().DoubleValue();
      break;
    case FM_VAL_BOOL:
      if (!ValueField(env, object, "boolean", &field)) {
        return false;
      }
      if (!field.IsBoolean()) {
        return RejectValue(env, "value.boolean must be a boolean");
      }
      out->u.boolean = field.As<Napi::Boolean>().Value() ? 1 : 0;
      break;
    case FM_VAL_TEXT:
      if (!ValueField(env, object, "text", &field)) {
        return false;
      }
      if (!field.IsString()) {
        return RejectValue(env, "value.text must be a string");
      }
      *text = field.As<Napi::String>().Utf8Value();
      out->u.text = text->c_str();
      break;
    case FM_VAL_ERROR:
      if (!ValueField(env, object, "errorCode", &field)) {
        return false;
      }
      if (!IsIntegerIn(field, static_cast<double>(std::numeric_limits<int32_t>::min()),
                       static_cast<double>(std::numeric_limits<int32_t>::max()))) {
        return RejectValue(env, "value.errorCode must be an ErrorCode");
      }
      out->u.error_code = static_cast<int32_t>(field.As<Napi::Number>().DoubleValue());
      break;
    default:
      break;
  }
  return true;
}

Napi::Object TranslatePivotCell(Napi::Env env, const fm_pivot_cell_t& cell) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("row", Napi::Number::New(env, cell.row));
  out.Set("col", Napi::Number::New(env, cell.col));
  out.Set("value", ValueToJs(env, cell.value));
  out.Set("kind", Napi::Number::New(env, static_cast<int32_t>(cell.kind)));
  out.Set("depth", Napi::Number::New(env, cell.depth));
  out.Set("fieldName", JsString(env, cell.field_name));
  out.Set("numberFormat", JsStringOrNull(env, cell.number_format));
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

Napi::Object ColorSpecToJs(Napi::Env env, const fm_color_spec_t& spec) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("kind", Napi::Number::New(env, static_cast<uint32_t>(spec.kind)));
  out.Set("rgb", Napi::Number::New(env, spec.rgb));
  out.Set("theme", Napi::Number::New(env, spec.theme));
  out.Set("tint", Napi::Number::New(env, spec.tint));
  out.Set("indexed", Napi::Number::New(env, spec.indexed));
  return out;
}

Napi::Object RangeToJs(Napi::Env env, const fm_cell_rect_t& r) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("firstRow", Napi::Number::New(env, r.first_row));
  item.Set("firstCol", Napi::Number::New(env, r.first_col));
  item.Set("lastRow", Napi::Number::New(env, r.last_row));
  item.Set("lastCol", Napi::Number::New(env, r.last_col));
  return item;
}

Napi::Value JsString(Napi::Env env, const char* s) {
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value JsNumber(Napi::Env env, double v) {
  return Napi::Number::New(env, v);
}

Napi::Value JsBool(Napi::Env env, int32_t v) {
  return Napi::Boolean::New(env, v != 0);
}

Napi::Value JsStringOrNull(Napi::Env env, const char* s) {
  return s != nullptr && s[0] != '\0' ? Napi::Value(Napi::String::New(env, s)) : env.Null();
}

Napi::Value JsIndexOrNull(Napi::Env env, int32_t v) {
  return v >= 0 ? Napi::Value(Napi::Number::New(env, v)) : env.Null();
}

Napi::Value JsNumberOrNull(Napi::Env env, int32_t present, double v) {
  return present != 0 ? Napi::Value(Napi::Number::New(env, v)) : env.Null();
}

Napi::Value JsBoolOrNull(Napi::Env env, int32_t present, int32_t v) {
  return present != 0 ? Napi::Value(Napi::Boolean::New(env, v != 0)) : env.Null();
}

Napi::Value PrintTitlesToJs(Napi::Env env, const fm_print_titles_t& titles) {
  if (titles.has_rows == 0 && titles.has_cols == 0) {
    return env.Null();
  }
  const auto span = [&](int32_t present, uint32_t first, uint32_t last) -> Napi::Value {
    if (present == 0) {
      return env.Null();
    }
    Napi::Object o = Napi::Object::New(env);
    o.Set("first", Napi::Number::New(env, first));
    o.Set("last", Napi::Number::New(env, last));
    return o;
  };
  Napi::Object out = Napi::Object::New(env);
  out.Set("rows", span(titles.has_rows, titles.first_row, titles.last_row));
  out.Set("columns", span(titles.has_cols, titles.first_col, titles.last_col));
  return out;
}

PatchKey ClassifyPatchKey(CheckedSpecReader& reader, const Napi::Value& value, const char* key, bool nullable) {
  if (!reader.ok() || value.IsUndefined()) {
    return PatchKey::kKeep;
  }
  if (value.IsNull()) {
    if (nullable) {
      return PatchKey::kClear;
    }
    reader.FailType(key, "value (null cannot clear this field)");
    return PatchKey::kKeep;
  }
  return PatchKey::kSet;
}

void BuildDataFieldSpec(CheckedSpecReader& reader, const Napi::Object& spec, fm_pivot_data_field_spec_t& out,
                        std::string& name_buf, std::string& nfmt_buf) {
  // An omitted `name` passes NULL through rather than the coerced literal
  // string "undefined" `.ToString()` would otherwise produce, so the C ABI
  // derives the data field's name from the workbook profile.
  const bool has_name = reader.String(spec, "name", &name_buf);
  const bool has_nfmt = reader.String(spec, "numberFormat", &nfmt_buf);
  out.name = has_name ? name_buf.c_str() : nullptr;
  out.field_index = reader.U32(spec, "fieldIndex", 0U);
  out.aggregation = static_cast<fm_pivot_aggregation_t>(reader.U32(spec, "aggregation", 0U));
  out.number_format = has_nfmt ? nfmt_buf.c_str() : nullptr;
  out.show_as = static_cast<fm_pivot_show_values_as_t>(reader.U32(spec, "showAs", 0U));
  out.show_as_base_field = reader.I32(spec, "showAsBaseField", -1);
  out.show_as_base_item = reader.I32(spec, "showAsBaseItem", -1);
}

// Reads a `{ firstRow, lastRow, firstCol, lastCol }` range object.
bool ReadCellRect(CheckedSpecReader& reader, const Napi::Object& range, fm_cell_rect_t* out) {
  out->first_row = reader.RequiredU32(range, "firstRow");
  out->first_col = reader.RequiredU32(range, "firstCol");
  out->last_row = reader.RequiredU32(range, "lastRow");
  out->last_col = reader.RequiredU32(range, "lastCol");
  return reader.ok();
}

bool ReadCellRects(CheckedSpecReader& reader, const Napi::Value& value, const char* key,
                   std::vector<fm_cell_rect_t>* out) {
  out->clear();
  Napi::Array array;
  if (!reader.Array(value, key, &array)) {
    return reader.ok();
  }
  const uint32_t n = array.Length();
  out->reserve(n);
  for (uint32_t i = 0; i < n && reader.ok(); ++i) {
    Napi::Value element;
    Napi::Object object;
    fm_cell_rect_t range{};
    if (reader.ArrayElement(array, i, &element) && reader.Object(element, key, &object) &&
        ReadCellRect(reader, object, &range)) {
      out->push_back(range);
    }
  }
  return reader.ok();
}

fm_color_spec_t PullColorSpec(CheckedSpecReader& reader, const Napi::Object& owner, const char* key) {
  fm_color_spec_t spec{};
  Napi::Object v;
  if (!reader.Object(owner, key, &v)) {
    return spec;
  }
  spec.kind = reader.U8(v, "kind", 0U);
  spec.rgb = reader.U32(v, "rgb", 0U);
  spec.theme = reader.U32(v, "theme", 0U);
  spec.tint = reader.Double(v, "tint", 0.0);
  spec.indexed = reader.U32(v, "indexed", 0U);
  return spec;
}

// Reads the range at `info[idx]`; a missing/nullish argument reads as all
// zeros, while a supplied non-object is rejected by the reader.
bool CellRectArg(CheckedSpecReader& reader, const Napi::CallbackInfo& info, size_t idx, fm_cell_rect_t* out) {
  *out = fm_cell_rect_t{};
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
  return ReadCellRect(reader, object, out);
}

}  // namespace formulon_node
