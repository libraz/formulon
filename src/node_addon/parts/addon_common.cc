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

namespace {
thread_local ScopedReaderLabel::Describe g_label_describe = nullptr;
thread_local ScopedReaderLabel::Name g_label_name = nullptr;
thread_local const void* g_label_data = nullptr;
}  // namespace

ScopedReaderLabel::ScopedReaderLabel(Describe describe, Name name, const void* data) noexcept
    : previous_describe_(g_label_describe), previous_name_(g_label_name), previous_data_(g_label_data) {
  g_label_describe = describe;
  g_label_name = name;
  g_label_data = data;
}

ScopedReaderLabel::~ScopedReaderLabel() {
  g_label_describe = previous_describe_;
  g_label_name = previous_name_;
  g_label_data = previous_data_;
}

std::string CurrentCallLabel() {
  return g_label_describe != nullptr ? g_label_describe(g_label_data) : std::string();
}

std::string CurrentArgumentName(std::size_t raw_index) {
  return g_label_name != nullptr ? g_label_name(g_label_data, raw_index) : std::string();
}

namespace {
// `parent.segment`, or `parent[i]` when the segment is an element index.
std::string JoinPath(const std::string& parent, const std::string& segment) {
  if (parent.empty() || segment.empty()) {
    return parent.empty() ? segment : parent;
  }
  return segment.front() == '[' ? parent + segment : formulon::binding::field_path(parent, segment);
}

bool EndsWithElementMark(const char* key) {
  const std::size_t length = std::strlen(key);
  return length >= 2 && std::strcmp(key + length - 2, "[]") == 0;
}
}  // namespace

// FORMULON-ALLOW: N-API uses JavaScript exceptions for argument validation.
void CheckedSpecReader::ObservePending() {
  if (!failed_ && env_.IsExceptionPending()) {
    failed_ = true;
  }
}

void CheckedSpecReader::Enter(const Napi::Value& owner) {
  prefix_.clear();
  element_ = false;
  for (const auto& entry : paths_) {
    if (entry.first.StrictEquals(owner)) {
      prefix_ = entry.second;
      return;
    }
  }
}

std::string CheckedSpecReader::PathOf(const char* key) const {
  return element_ || EndsWithElementMark(key) ? prefix_ : JoinPath(prefix_, key);
}

void CheckedSpecReader::Register(const Napi::Value& value, const char* key) {
  paths_.emplace_back(value, PathOf(key));
}

template <typename ErrorType>
void CheckedSpecReader::ReportFailure(const char* key, const std::string& constraint) {
  if (failed_ || env_.IsExceptionPending()) {
    failed_ = true;
    return;
  }
  const std::string call = CurrentCallLabel();
  const std::string path = formulon::binding::record_path(call, PathOf(key));
  const std::string message = call.empty() ? (path.empty() ? constraint : path + " " + constraint)
                                           : formulon::binding::argument_message(call, path, constraint);
  ErrorType::New(env_, message).ThrowAsJavaScriptException();
  failed_ = true;
}

void CheckedSpecReader::FailType(const char* key, const char* expected) {
  ReportFailure<Napi::TypeError>(key, std::string("must be ") + expected);
}

void CheckedSpecReader::ReportNegative(const char* key) {
  if (failed_ || env_.IsExceptionPending()) {
    failed_ = true;
    return;
  }
  const std::string call = CurrentCallLabel();
  const std::string path = formulon::binding::record_path(call, PathOf(key));
  ThrowBindingStatus(env_, FM_STATUS_INVALID_ARGUMENT,
                     call.empty() ? path + " must not be negative"
                                  : formulon::binding::argument_message(call, path, "must not be negative"));
  failed_ = true;
}

void ThrowArgumentTypeError(Napi::Env env, const std::string& path, const std::string& constraint) {
  if (env.IsExceptionPending()) {
    return;
  }
  const std::string call = CurrentCallLabel();
  const std::string where = path.empty() ? std::string(formulon::binding::record_param(call)) : path;
  Napi::TypeError::New(env, call.empty() ? (where.empty() ? constraint : where + " " + constraint)
                                         : formulon::binding::argument_message(call, where, constraint))
      .ThrowAsJavaScriptException();
}

void ThrowIntegerRefusal(Napi::Env env, const std::string& call, const std::string& path, double number,
                         double min_value) {
  if (env.IsExceptionPending()) {
    return;
  }
  const auto message = [&](const char* constraint) {
    return call.empty() ? (path.empty() ? std::string(constraint) : path + " " + constraint)
                        : formulon::binding::argument_message(call, path, constraint);
  };
  if (!std::isfinite(number) || std::trunc(number) != number) {
    Napi::RangeError::New(env, message("must be an integer")).ThrowAsJavaScriptException();
  } else if (number < 0.0 && min_value >= 0.0) {
    ThrowBindingStatus(env, FM_STATUS_INVALID_ARGUMENT, message("must not be negative"));
  } else {
    Napi::RangeError::New(env, message("is too large")).ThrowAsJavaScriptException();
  }
}

void ThrowPositionalTypeError(Napi::Env env, std::size_t raw_index, const std::string& constraint) {
  if (env.IsExceptionPending()) {
    return;
  }
  Napi::TypeError::New(
      env, formulon::binding::argument_message(CurrentCallLabel(), CurrentArgumentName(raw_index), constraint))
      .ThrowAsJavaScriptException();
}

void ThrowPositionalIntegerRefusal(Napi::Env env, std::size_t raw_index, double number, double min_value) {
  ThrowIntegerRefusal(env, CurrentCallLabel(), CurrentArgumentName(raw_index), number, min_value);
}

void CheckedSpecReader::FailInteger(const char* key, double number, double min_value) {
  if (!std::isfinite(number) || std::trunc(number) != number) {
    ReportFailure<Napi::RangeError>(key, "must be an integer");
  } else if (number < 0.0 && min_value >= 0.0) {
    ReportNegative(key);
  } else {
    ReportFailure<Napi::RangeError>(key, "is too large");
  }
}

bool CheckedSpecReader::ReadOptional(const Napi::Object& owner, const char* key, Napi::Value* out, bool* present) {
  if (present != nullptr) {
    *present = false;
  }
  if (!ok()) {
    return false;
  }
  Enter(owner);
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
  Enter(owner);
  prefix_ = JoinPath(prefix_, "[" + std::to_string(index) + "]");
  element_ = true;
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
    FailType(key, "a number");
    return false;
  }
  *out = present.As<Napi::Number>().DoubleValue();
  ObservePending();
  return ok();
}

bool CheckedSpecReader::ReadInteger(const Napi::Object& owner, const char* key, double* out, double min_value,
                                    double max_value, bool* present) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value, present)) {
    return ok();
  }
  return ReadInteger(value, key, out, min_value, max_value);
}

bool CheckedSpecReader::ReadInteger(const Napi::Value& value, const char* key, double* out, double min_value,
                                    double max_value) {
  if (!ReadNumber(value, key, out)) {
    return false;
  }
  if (!std::isfinite(*out) || std::trunc(*out) != *out || *out < min_value || *out > max_value) {
    FailInteger(key, *out, min_value);
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
    FailType(key, "a string");
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
    FailType(key, "an object");
    return false;
  }
  *out = present.As<Napi::Object>();
  ObservePending();
  Register(*out, key);
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
    FailType(key, "an array");
    return false;
  }
  *out = present.As<Napi::Array>();
  ObservePending();
  Register(*out, key);
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
  if (!present.IsBoolean()) {
    FailType(key, "a boolean");
    return dflt;
  }
  const bool result = present.As<Napi::Boolean>().Value();
  ObservePending();
  return result;
}

int32_t CheckedSpecReader::Tristate(const Napi::Object& owner, const char* key) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value)) {
    return FM_TRISTATE_DEFAULT;
  }
  return Tristate(value, key);
}

int32_t CheckedSpecReader::Tristate(const Napi::Value& value, const char* key) {
  if (value.IsUndefined() || value.IsNull()) {
    return FM_TRISTATE_DEFAULT;
  }
  return Bool(value, key, false) ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
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
T CheckedSpecReader::ReadIntegral(const Napi::Object& owner, const char* key, T dflt, bool* present) {
  double value = static_cast<double>(dflt);
  if (!ReadInteger(owner, key, &value, static_cast<double>(std::numeric_limits<T>::min()),
                   static_cast<double>(std::numeric_limits<T>::max()), present)) {
    return dflt;
  }
  return static_cast<T>(value);
}

template <typename T>
T CheckedSpecReader::ReadIntegral(const Napi::Value& value, const char* key, T dflt) {
  double result = static_cast<double>(dflt);
  if (!ReadInteger(value, key, &result, static_cast<double>(std::numeric_limits<T>::min()),
                   static_cast<double>(std::numeric_limits<T>::max()))) {
    return dflt;
  }
  return static_cast<T>(result);
}

uint8_t CheckedSpecReader::U8(const Napi::Object& owner, const char* key, uint8_t dflt, bool* present) {
  return ReadIntegral<uint8_t>(owner, key, dflt, present);
}

uint8_t CheckedSpecReader::U8(const Napi::Value& value, const char* key, uint8_t dflt) {
  return ReadIntegral<uint8_t>(value, key, dflt);
}

uint16_t CheckedSpecReader::U16(const Napi::Object& owner, const char* key, uint16_t dflt, bool* present) {
  return ReadIntegral<uint16_t>(owner, key, dflt, present);
}

uint16_t CheckedSpecReader::U16(const Napi::Value& value, const char* key, uint16_t dflt) {
  return ReadIntegral<uint16_t>(value, key, dflt);
}

uint32_t CheckedSpecReader::U32(const Napi::Object& owner, const char* key, uint32_t dflt, bool* present) {
  return ReadIntegral<uint32_t>(owner, key, dflt, present);
}

uint32_t CheckedSpecReader::U32(const Napi::Value& value, const char* key, uint32_t dflt) {
  return ReadIntegral<uint32_t>(value, key, dflt);
}

int32_t CheckedSpecReader::I32(const Napi::Object& owner, const char* key, int32_t dflt, bool* present) {
  return ReadIntegral<int32_t>(owner, key, dflt, present);
}

int32_t CheckedSpecReader::I32(const Napi::Value& value, const char* key, int32_t dflt) {
  return ReadIntegral<int32_t>(value, key, dflt);
}

bool CheckedSpecReader::RequiredObject(const Napi::Object& owner, const char* key, Napi::Object* out) {
  if (Object(owner, key, out)) {
    return true;
  }
  if (ok()) {
    FailType(key, "an object");
  }
  return false;
}

bool CheckedSpecReader::RequiredArray(const Napi::Object& owner, const char* key, Napi::Array* out) {
  if (Array(owner, key, out)) {
    return true;
  }
  if (ok()) {
    FailType(key, "an array");
  }
  return false;
}

std::string CheckedSpecReader::RequiredString(const Napi::Object& owner, const char* key) {
  const Napi::Value value = Get(owner, key);
  if (!ok()) {
    return std::string();
  }
  if (!value.IsString()) {
    FailType(key, "a string");
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
  Enter(owner);
  const Napi::Value value = owner.Get(key);
  ObservePending();
  return ok() ? value : env_.Undefined();
}

uint32_t CheckedSpecReader::RequiredU32(const Napi::Object& owner, const char* key) {
  Napi::Value value;
  if (!ReadOptional(owner, key, &value)) {
    if (ok()) {
      FailType(key, "a number");
    }
    return 0U;
  }
  if (!value.IsNumber()) {
    FailType(key, "a number");
    return 0U;
  }
  const double number = value.As<Napi::Number>().DoubleValue();
  if (!std::isfinite(number) || std::trunc(number) != number || number < 0.0 || number > kUint32MaxNumber) {
    FailInteger(key, number, 0.0);
    return 0U;
  }
  return static_cast<uint32_t>(number);
}

int64_t CheckedSpecReader::NarrowI64(double number, const char* key, int64_t dflt) {
  if (!std::isfinite(number) || std::trunc(number) != number || number < -0x1p63 || number >= 0x1p63) {
    FailInteger(key, number, -0x1p63);
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
    FailType(key, "a number");
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

// The `parseError` argument of a failure: `{ code, start, end, line, column }`
// when the failing call recorded a formula syntax error, else `undefined`.
Napi::Value LastParseErrorToJs(Napi::Env env) {
  fm_parse_error_info_t record = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;
  if (fm_get_last_parse_error(&present, &record) != 0 || present == 0) {
    return env.Undefined();
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("code", Napi::Number::New(env, record.code));
  out.Set("start", Napi::Number::New(env, record.start));
  out.Set("end", Napi::Number::New(env, record.end));
  out.Set("line", Napi::Number::New(env, record.line));
  out.Set("column", Napi::Number::New(env, record.column));
  return out;
}

// Throws `new FormulonError(code, message, context, parseError)`; NULL strings become "".
Napi::Value ThrowFormulonError(Napi::Env env, fm_status_t code, const char* message, const char* context,
                               Napi::Value parse_error) {
  const Napi::FunctionReference& error_class = GetAddonData(env).error_class;
  if (error_class.IsEmpty()) {
    Napi::TypeError::New(env, kMissingErrorClassMessage).ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const Napi::Object error = error_class.New({
      Napi::Number::New(env, static_cast<double>(code)),
      JsString(env, message),
      JsString(env, context),
      parse_error,
  });
  if (!env.IsExceptionPending()) {
    napi_throw(env, error);
  }
  return env.Undefined();
}

// Rejects a malformed input `Value`.
bool RejectValue(Napi::Env env, const char* path, const char* constraint) {
  ThrowArgumentTypeError(env, path, constraint);
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
  // The parse-error record first, then message and context, before any other `fm_*` call.
  const Napi::Value parse_error = LastParseErrorToJs(env);
  const char* message = fm_get_last_error_message();
  const char* context = fm_get_last_error_context();
  return ThrowFormulonError(env, rc, message, context, parse_error);
}

Napi::Value ThrowBindingStatus(Napi::Env env, fm_status_t code, const std::string& message) {
  if (env.IsExceptionPending()) {
    return env.Undefined();
  }
  return ThrowFormulonError(env, code, message.c_str(), "", env.Undefined());
}

Napi::Value ThrowInvalidHandle(Napi::Env env) {
  if (env.IsExceptionPending()) {
    return env.Undefined();
  }
  return ThrowFormulonError(env, kBindingInvalidHandle, "the workbook is disposed", "", env.Undefined());
}

Napi::Value ThrowReentrant(Napi::Env env, const char* message) {
  if (env.IsExceptionPending()) {
    return env.Undefined();
  }
  return ThrowFormulonError(env, FM_STATUS_GRAPH_RECALC_REENTRANT, message, "", env.Undefined());
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
    case FM_VALUE_KIND_NUMBER:
      o.Set("number", Napi::Number::New(env, v.u.number));
      break;
    case FM_VALUE_KIND_BOOL:
      o.Set("boolean", Napi::Boolean::New(env, v.u.boolean != 0));
      break;
    case FM_VALUE_KIND_TEXT:
      o.Set("text", JsString(env, v.u.text));
      break;
    case FM_VALUE_KIND_ERROR:
      o.Set("errorCode", Napi::Number::New(env, v.u.error_code));
      break;
    default:
      break;
  }
  return o;
}

Napi::Object ValueArrayToJs(Napi::Env env, const fm_value_array_t* array) {
  uint32_t rows = 0;
  uint32_t cols = 0;
  fm_value_array_get_dimensions(array, &rows, &cols);
  Napi::Array cells = Napi::Array::New(env, rows);
  for (uint32_t r = 0; r < rows; ++r) {
    Napi::Array js_row = Napi::Array::New(env, cols);
    for (uint32_t c = 0; c < cols; ++c) {
      fm_value_t v{};
      fm_value_array_get_value_at(array, static_cast<std::size_t>(r) * cols + c, &v);
      js_row.Set(c, ValueToJs(env, v));
    }
    cells.Set(r, js_row);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("rows", Napi::Number::New(env, rows));
  out.Set("cols", Napi::Number::New(env, cols));
  out.Set("cells", cells);
  return out;
}

bool JsToValue(Napi::Env env, Napi::Value in, fm_value_t* out, std::string* text) {
  *out = fm_value_t{};
  if (env.IsExceptionPending()) {
    return false;
  }
  if (!in.IsObject()) {
    return RejectValue(env, "value", "must be an object");
  }
  const Napi::Object object = in.As<Napi::Object>();
  Napi::Value kind;
  if (!ValueField(env, object, "kind", &kind)) {
    return false;
  }
  if (!IsIntegerIn(kind, static_cast<double>(FM_VALUE_KIND_BLANK), static_cast<double>(FM_VALUE_KIND_LAMBDA))) {
    return RejectValue(env, "value.kind", "must be a ValueKind");
  }
  out->kind = static_cast<int32_t>(kind.As<Napi::Number>().DoubleValue());
  Napi::Value field;
  switch (out->kind) {
    case FM_VALUE_KIND_NUMBER:
      if (!ValueField(env, object, "number", &field)) {
        return false;
      }
      if (!field.IsNumber()) {
        return RejectValue(env, "value.number", "must be a number");
      }
      out->u.number = field.As<Napi::Number>().DoubleValue();
      break;
    case FM_VALUE_KIND_BOOL:
      if (!ValueField(env, object, "boolean", &field)) {
        return false;
      }
      if (!field.IsBoolean()) {
        return RejectValue(env, "value.boolean", "must be a boolean");
      }
      out->u.boolean = field.As<Napi::Boolean>().Value() ? 1 : 0;
      break;
    case FM_VALUE_KIND_TEXT:
      if (!ValueField(env, object, "text", &field)) {
        return false;
      }
      if (!field.IsString()) {
        return RejectValue(env, "value.text", "must be a string");
      }
      *text = field.As<Napi::String>().Utf8Value();
      out->u.text = text->c_str();
      break;
    case FM_VALUE_KIND_ERROR:
      if (!ValueField(env, object, "errorCode", &field)) {
        return false;
      }
      if (!IsIntegerIn(field, static_cast<double>(std::numeric_limits<int32_t>::min()),
                       static_cast<double>(std::numeric_limits<int32_t>::max()))) {
        return RejectValue(env, "value.errorCode", "must be an ErrorCode");
      }
      out->u.error_code = static_cast<int32_t>(field.As<Napi::Number>().DoubleValue());
      break;
    default:
      break;
  }
  return true;
}

Napi::Object TranslatePivotCell(Napi::Env env, const fm_pivot_cell_info_t& cell) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("row", Napi::Number::New(env, cell.row));
  out.Set("col", Napi::Number::New(env, cell.col));
  out.Set("value", ValueToJs(env, cell.value));
  out.Set("kind", Napi::Number::New(env, static_cast<int32_t>(cell.kind)));
  out.Set("depth", Napi::Number::New(env, cell.depth));
  out.Set("fieldName", JsStringOrNull(env, cell.field_name));
  out.Set("numberFormat", JsStringOrNull(env, cell.number_format));
  return out;
}

Napi::Object TranslateCfMatch(Napi::Env env, const fm_cf_match_info_t& m) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("kind", Napi::Number::New(env, static_cast<int32_t>(m.kind)));
  o.Set("priority", Napi::Number::New(env, m.priority));
  // Each kind carries only its own fields.
  switch (m.kind) {
    case FM_CF_MATCH_KIND_DIFFERENTIAL_FORMAT:
      o.Set("dxfIndex", m.has_dxf_index != 0 ? Napi::Value(Napi::Number::New(env, m.dxf_index)) : env.Null());
      break;
    case FM_CF_MATCH_KIND_COLOR_SCALE:
      o.Set("color", ColorSpecToJs(env, m.color));
      break;
    case FM_CF_MATCH_KIND_DATA_BAR: {
      Napi::Object bar = Napi::Object::New(env);
      bar.Set("lengthPct", Napi::Number::New(env, m.bar_length_pct));
      bar.Set("axisPositionPct", Napi::Number::New(env, m.bar_axis_position_pct));
      bar.Set("isNegative", Napi::Boolean::New(env, m.bar_is_negative != 0));
      bar.Set("fill", ColorSpecToJs(env, m.bar_fill));
      bar.Set("border",
              m.bar_border.kind != FM_COLOR_KIND_NONE ? Napi::Value(ColorSpecToJs(env, m.bar_border)) : env.Null());
      bar.Set("gradient", Napi::Boolean::New(env, m.bar_gradient != 0));
      bar.Set("direction", Napi::Number::New(env, static_cast<int32_t>(m.bar_direction)));
      o.Set("bar", bar);
      break;
    }
    case FM_CF_MATCH_KIND_ICON_SET: {
      Napi::Object icon = Napi::Object::New(env);
      icon.Set("iconSet", Napi::Number::New(env, m.icon_set));
      icon.Set("iconIndex", Napi::Number::New(env, static_cast<int32_t>(m.icon_index)));
      o.Set("icon", icon);
      break;
    }
    default:
      break;
  }
  return o;
}

Napi::Object ColorSpecToJs(Napi::Env env, const fm_color_t& spec) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("kind", Napi::Number::New(env, static_cast<uint32_t>(spec.kind)));
  out.Set("argb", Napi::Number::New(env, spec.argb));
  out.Set("theme", Napi::Number::New(env, spec.theme));
  out.Set("tint", Napi::Number::New(env, spec.tint));
  out.Set("indexed", Napi::Number::New(env, spec.indexed));
  return out;
}

Napi::Value ColorSpecOrNull(Napi::Env env, const fm_color_t& spec) {
  return spec.kind != FM_COLOR_KIND_NONE ? Napi::Value(ColorSpecToJs(env, spec)) : env.Null();
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

Napi::Value JsTristate(Napi::Env env, int32_t v) {
  return Napi::Boolean::New(env, v == FM_TRISTATE_TRUE);
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
  if (titles.has_rows == 0 && titles.has_columns == 0) {
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
  out.Set("columns", span(titles.has_columns, titles.first_col, titles.last_col));
  return out;
}

PatchKey ClassifyPatchKey(CheckedSpecReader& reader, const Napi::Value& value) {
  if (!reader.ok() || value.IsUndefined()) {
    return PatchKey::kKeep;
  }
  // `null` always packs as CLEAR; the C ABI refuses it on a field that cannot be cleared.
  return value.IsNull() ? PatchKey::kClear : PatchKey::kSet;
}

namespace {

template <typename Read>
fm_patch_op_t PatchField(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, Read&& read) {
  const Napi::Value value = reader.Get(spec, key);
  const PatchKey state = ClassifyPatchKey(reader, value);
  if (state == PatchKey::kSet) {
    read(value);
  }
  return static_cast<fm_patch_op_t>(state);
}

}  // namespace

fm_patch_op_t PatchBool(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int32_t* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = reader.Bool(v, key, false) ? 1 : 0; });
}

fm_patch_op_t PatchI32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int32_t* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = reader.I32(v, key, 0); });
}

fm_patch_op_t PatchU32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, uint32_t* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = reader.U32(v, key, 0U); });
}

fm_patch_op_t PatchI64(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int64_t* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = reader.I64(v, key, 0); });
}

fm_patch_op_t PatchColor(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, fm_color_t* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = ReadColorSpec(reader, v, key); });
}

fm_patch_op_t PatchNullableDouble(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, double* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = reader.Double(v, key, 0.0); });
}

fm_patch_op_t PatchNullableI32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int32_t* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = reader.I32(v, key, 0); });
}

fm_patch_op_t PatchNullableU32(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, uint32_t* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = reader.U32(v, key, 0U); });
}

fm_patch_op_t PatchDouble(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, double* out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) { *out = reader.Double(v, key, 0.0); });
}

fm_patch_op_t PatchString(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, std::string* store,
                          const char** out) {
  return PatchField(reader, spec, key, [&](const Napi::Value& v) {
    if (reader.String(v, key, store)) {
      *out = store->c_str();
    }
  });
}

void BuildDataFieldSpec(CheckedSpecReader& reader, const Napi::Object& spec, fm_pivot_data_field_t& out,
                        std::string& name_buf, std::string& nfmt_buf) {
  // An omitted `name` passes NULL through rather than the coerced literal
  // string "undefined" `.ToString()` would otherwise produce, so the C ABI
  // derives the data field's name from the workbook profile.
  const bool has_name = reader.String(spec, "name", &name_buf);
  const bool has_nfmt = reader.String(spec, "numberFormat", &nfmt_buf);
  out.name = has_name ? name_buf.c_str() : nullptr;
  out.field_index = reader.RequiredU32(spec, "fieldIndex");
  out.aggregation = static_cast<fm_pivot_aggregation_t>(reader.U32(spec, "aggregation", FM_PIVOT_AGGREGATION_SUM));
  out.number_format = has_nfmt ? nfmt_buf.c_str() : nullptr;
  out.show_as = static_cast<fm_pivot_show_values_as_t>(reader.U32(spec, "showAs", 0U));
  bool base_field_present = false;
  out.show_as_base_field = reader.U32(spec, "showAsBaseField", 0U, &base_field_present);
  out.has_show_as_base_field = base_field_present ? 1 : 0;
  bool base_item_present = false;
  out.show_as_base_item = reader.U32(spec, "showAsBaseItem", 0U, &base_item_present);
  out.has_show_as_base_item = base_item_present ? 1 : 0;
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

namespace {

fm_color_t ColorSpecFields(CheckedSpecReader& reader, const Napi::Object& v) {
  fm_color_t spec{};
  spec.kind = reader.I32(v, "kind", 0);
  spec.argb = reader.U32(v, "argb", 0U);
  spec.theme = reader.U32(v, "theme", 0U);
  spec.tint = reader.Double(v, "tint", 0.0);
  spec.indexed = reader.U32(v, "indexed", 0U);
  return spec;
}

}  // namespace

fm_color_t PullColorSpec(CheckedSpecReader& reader, const Napi::Object& owner, const char* key) {
  Napi::Object v;
  return reader.Object(owner, key, &v) ? ColorSpecFields(reader, v) : fm_color_t{};
}

fm_color_t ReadColorSpec(CheckedSpecReader& reader, const Napi::Value& value, const char* key) {
  Napi::Object v;
  return reader.Object(value, key, &v) ? ColorSpecFields(reader, v) : fm_color_t{};
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
