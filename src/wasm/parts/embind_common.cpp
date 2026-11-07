//
// Out-of-line implementation of the shared embind helpers declared in
// `parts/embind_common.h`. The failure builders record the thread-local
// diagnostic surface that follows every C-ABI call for the pre-js wrapper
// to throw, the translation helpers project a `fm_*` POD into a JS value,
// and the field readers pull caller-supplied JS records apart.

#include "wasm/parts/embind_common.h"

#include <emscripten/bind.h>
#include <emscripten/em_js.h>
#include <emscripten/val.h>

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

// The operation codes are deliberately kept private to this translation
// unit. Every operation that can invoke caller-controlled JavaScript runs in
// this one try/catch boundary; no C++ `val` property/call path is allowed to
// observe a user getter directly while exceptions are disabled.
enum SafeReaderOperation : int32_t {
  kSafeGet = 0,
  kSafeHasOwn = 1,
  kSafeIsArray = 2,
  kSafeStringSnapshot = 3,
  kSafeBytesSnapshot = 4,
};

// The error class the pre-js wrapper raises for a reader rejection.
enum ArgumentErrorKind : int32_t {
  kArgumentErrorType = 0,
  kArgumentErrorRange = 1,
  kArgumentErrorAccess = 2,
};

}  // namespace

// `success` reports whether the operation itself completed. A completed
// string/bytes snapshot may still return `undefined` for an unsupported
// input; the C++ reader turns that into its normal type error. A thrown proxy,
// getter, or typed-array hook reports success == 0 and returns an owned
// `undefined` handle. A string snapshot replaces each lone surrogate with
// U+FFFD, as N-API's UTF-8 conversion does, so ill-formed UTF-8 never reaches
// the engine. The body is JavaScript source; keep formatting disabled so
// clang-format cannot rewrite its operators.
// FORMULON-ALLOW: this JavaScript try/catch is the required no-escape boundary
// for caller-owned getters and typed-array operations in a -fno-exceptions WASM
// build; a C++ try/catch cannot safely intercept those throws here.
// clang-format off
EM_JS(emscripten::EM_VAL, fm_wasm_safe_reader_operation,
      (emscripten::EM_VAL owner_handle, emscripten::EM_VAL key_handle, int32_t operation, int32_t* success), {
  try {
    var owner = Emval.toValue(owner_handle);
    var key = Emval.toValue(key_handle);
    var result;
    if (operation === 0) {
      result = owner[key];
    } else if (operation === 1) {
      result = Object.prototype.hasOwnProperty.call(owner, key);
    } else if (operation === 2) {
      result = Array.isArray(owner);
    } else if (operation === 3) {
      if (typeof owner === 'string') {
        if (typeof owner.toWellFormed === 'function') {
          result = owner.toWellFormed();
        } else {
          result = '';
          for (var i = 0; i < owner.length; i++) {
            var unit = owner.charCodeAt(i);
            var next = i + 1 < owner.length ? owner.charCodeAt(i + 1) : 0;
            if (unit >= 0xD800 && unit <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) {
              result += owner.substring(i, i + 2);
              i++;
            } else {
              result += unit >= 0xD800 && unit <= 0xDFFF ? String.fromCharCode(0xFFFD) : owner.charAt(i);
            }
          }
        }
      } else if (owner instanceof ArrayBuffer) {
        var bufferCopy = new Uint8Array(owner.byteLength);
        bufferCopy.set(new Uint8Array(owner));
        result = bufferCopy;
      } else if (ArrayBuffer.isView(owner) && owner.BYTES_PER_ELEMENT === 1) {
        var viewCopy = new Uint8Array(owner.byteLength);
        viewCopy.set(new Uint8Array(owner.buffer, owner.byteOffset, owner.byteLength));
        result = viewCopy;
      } else {
        result = undefined;
      }
    } else if (operation === 4) {
      // `isView` is false for a Proxy, which the Node binding also rejects.
      if (ArrayBuffer.isView(owner) && owner instanceof Uint8Array) {
        var bytesCopy = new Uint8Array(owner.length);
        bytesCopy.set(owner);
        result = bytesCopy;
      } else {
        result = undefined;
      }
    } else {
      result = undefined;
    }
    HEAP32[success >> 2] = 1;
    return Emval.toHandle(result);
  } catch (e) {
    var errors = Module[Symbol.for('formulon.argumentErrors')];
    if (errors) {
      errors.thrown = { value: e };
    }
    HEAP32[success >> 2] = 0;
    return Emval.toHandle(undefined);
  }
});
// clang-format on

EM_JS_DEPS(fm_wasm_safe_reader, "$Emval");

// Hands a reader's first rejection to the pre-js method wrapper, which throws
// it once the native method has returned: kind 0 is a TypeError, 1 a
// RangeError, 2 rethrows the caller's getter exception. It displaces a
// recorded status failure (kind 3), so an argument error wins within a call.
// Outside a wrapped call there is no wrapper to throw, and it is dropped.
// clang-format off
EM_JS(void, fm_wasm_record_argument_error, (int32_t kind, const char* message), {
  var errors = Module[Symbol.for('formulon.argumentErrors')];
  if (!errors || errors.depth === 0 || (errors.pending !== undefined && errors.pending.kind !== 3)) {
    return;
  }
  var thrown = kind === 2 ? errors.thrown : undefined;
  errors.thrown = undefined;
  errors.pending = { kind: kind, message: UTF8ToString(message), thrown: thrown };
});
// clang-format on

EM_JS_DEPS(fm_wasm_argument_error, "$UTF8ToString");

// Records a failed status for the pre-js method wrapper, which throws it as
// `FormulonError(code, message, context, parseError)` once the native method
// has returned; `parse_error` is NULL unless the call recorded a formula
// syntax error. Any failure already recorded in the call is kept.
// clang-format off
EM_JS(void, fm_wasm_record_status_error,
      (int32_t code, const char* message, const char* context, const fm_parse_error_t* parse_error), {
  var errors = Module[Symbol.for('formulon.argumentErrors')];
  if (!errors || errors.depth === 0 || errors.pending !== undefined) {
    return;
  }
  var parsed = undefined;
  if (parse_error) {
    var base = parse_error >> 2;
    parsed = { code: HEAP32[base + 1], start: HEAPU32[base + 2], end: HEAPU32[base + 3],
               line: HEAPU32[base + 4], column: HEAPU32[base + 5] };
  }
  errors.pending = { kind: 3, code: code, message: UTF8ToString(message), context: UTF8ToString(context),
                     parseError: parsed };
});
// clang-format on

EM_JS_DEPS(fm_wasm_status_error, "$UTF8ToString");

// Copies a previously snapshotted plain Uint8Array into a C++ vector's
// allocation. The source is produced by the trampoline above, and both the
// length read and the typed-array copy are kept inside this second boundary so
// a modified intrinsic cannot throw through a C++ frame.
// FORMULON-ALLOW: this JavaScript try/catch is required to catch a hostile
// typed-array operation while C++ exceptions are disabled for the WASM ABI.
// clang-format off
EM_JS(int32_t, fm_wasm_copy_safe_bytes,
      (emscripten::EM_VAL source_handle, uintptr_t destination, uint32_t capacity, uint32_t* written), {
  try {
    var source = Emval.toValue(source_handle);
    if (!(source instanceof Uint8Array) || source.length > capacity) {
      HEAPU32[written >> 2] = 0;
      return 0;
    }
    HEAPU8.set(source, destination);
    HEAPU32[written >> 2] = source.length;
    return 1;
  } catch (e) {
    HEAPU32[written >> 2] = 0;
    return -1;
  }
});
// clang-format on

EM_JS_DEPS(fm_wasm_safe_bytes, "$Emval");

bool js_value_present(const emscripten::val& value) {
  return !value.isUndefined() && !value.isNull();
}

emscripten::val js_object_or_empty(const emscripten::val& value) {
  return js_value_present(value) ? value : emscripten::val::object();
}

void record_binding_failure(int32_t code, const char* message) {
  fm_wasm_record_status_error(code, message != nullptr ? message : "", "", nullptr);
}

void record_type_error(const char* message) {
  fm_wasm_record_argument_error(kArgumentErrorType, message != nullptr ? message : "");
}

void record_range_error(const char* message) {
  fm_wasm_record_argument_error(kArgumentErrorRange, message != nullptr ? message : "");
}

void record_invalid_handle() {
  record_binding_failure(kBindingInvalidHandle, "the workbook is disposed");
}

void record_failure(fm_status_t rc) {
  if (rc == kBindingInvalidHandle) {
    // No C-ABI call ran, so the thread-local diagnostics still describe
    // whatever call came before.
    record_invalid_handle();
    return;
  }
  // The parse-error record first, then message and context, before any other `fm_*` call.
  fm_parse_error_t parse_error = FM_STRUCT_INIT(fm_parse_error_t);
  const bool has_parse_error = fm_get_last_parse_error(&parse_error) == 0;
  const char* message = fm_get_last_error_message();
  const char* context = fm_get_last_error_context();
  fm_wasm_record_status_error(rc, message != nullptr ? message : "", context != nullptr ? context : "",
                              has_parse_error ? &parse_error : nullptr);
}

bool succeeded(fm_status_t rc) {
  if (rc == 0) {
    return true;
  }
  record_failure(rc);
  return false;
}

emscripten::val optional_or_null(fm_status_t rc, int32_t present, emscripten::val item) {
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  return present != 0 ? item : emscripten::val::null();
}

JsNarrowNumericReader::JsNarrowNumericReader(const char* operation)
    : operation_(operation != nullptr ? operation : "WASM binding") {}

void JsNarrowNumericReader::reject_access(const char* key, const char* field) {
  if (!message_.empty()) {
    return;
  }
  const char* name = field != nullptr ? field : key;
  message_ = operation_ + ": `" + (name != nullptr ? name : "<field>") + "` field access threw";
  fm_wasm_record_argument_error(kArgumentErrorAccess, message_.c_str());
}

void JsNarrowNumericReader::reject(const char* key, const char* field, const char* range) {
  if (!message_.empty()) {
    return;
  }
  const char* name = field != nullptr ? field : key;
  message_ = operation_ + ": `" + (name != nullptr ? name : "<field>") + "` must be a finite integer in " + range;
  fm_wasm_record_argument_error(kArgumentErrorRange, message_.c_str());
}

void JsNarrowNumericReader::reject_type(const char* field, const char* expected) {
  if (!message_.empty()) {
    return;
  }
  message_ = operation_ + ": `" + (field != nullptr ? field : "<field>") + "` must be " +
             (expected != nullptr ? expected : "the expected type");
  fm_wasm_record_argument_error(kArgumentErrorType, message_.c_str());
}

emscripten::val JsNarrowNumericReader::safe_operation(const emscripten::val& owner, const emscripten::val& key,
                                                      int32_t operation, bool* succeeded) {
  int32_t success = 0;
  const emscripten::EM_VAL handle =
      fm_wasm_safe_reader_operation(owner.as_handle(), key.as_handle(), operation, &success);
  if (succeeded != nullptr) {
    *succeeded = success != 0;
  }
  return emscripten::val::take_ownership(handle);
}

emscripten::val JsNarrowNumericReader::safe_get(const emscripten::val& owner, const emscripten::val& key,
                                                const char* field, bool* succeeded) {
  bool completed = false;
  emscripten::val result = safe_operation(owner, key, kSafeGet, &completed);
  if (!completed) {
    reject_access(nullptr, field);
  }
  if (succeeded != nullptr) {
    *succeeded = completed;
  }
  return result;
}

bool JsNarrowNumericReader::present_number(const emscripten::val& value, const char* key, const char* field) {
  if (!js_value_present(value)) {
    return false;
  }
  // Check the JavaScript primitive before asking embind for a C++ value.
  // In particular, Symbol and BigInt must never reach val::as on this path.
  if (!value.isNumber()) {
    reject_type(field != nullptr ? field : key, "a Number");
    return false;
  }
  return true;
}

bool JsNarrowNumericReader::read_integer(const emscripten::val& value, const char* key, const char* field, double lower,
                                         double upper, const char* range, double* number) {
  if (!present_number(value, key, field)) {
    return false;
  }
  const double candidate = value.as<double>();
  if (!std::isfinite(candidate) || std::floor(candidate) != candidate || candidate < lower || candidate > upper) {
    reject(key, field, range);
    return false;
  }
  *number = candidate;
  return true;
}

bool JsNarrowNumericReader::read_number(const emscripten::val& value, const char* key, const char* field,
                                        double* number) {
  if (!present_number(value, key, field)) {
    return false;
  }
  *number = value.as<double>();
  return true;
}

emscripten::val JsNarrowNumericReader::fetch(const emscripten::val& owner, const char* key, const char* field,
                                             bool* fetched) {
  *fetched = false;
  if (!ok()) {
    return emscripten::val::undefined();
  }
  return safe_get(owner, emscripten::val(key), field != nullptr ? field : key, fetched);
}

double JsNarrowNumericReader::integer_value(const emscripten::val& value, double dflt, const char* field, double lower,
                                            double upper, const char* range) {
  if (!ok()) {
    return dflt;
  }
  double number = 0.0;
  if (!read_integer(value, field, field, lower, upper, range, &number)) {
    return !js_value_present(value) ? dflt : 0.0;
  }
  return number;
}

uint32_t JsNarrowNumericReader::u32(const emscripten::val& owner, const char* key, uint32_t dflt, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  return fetched ? u32_value(value, dflt, field != nullptr ? field : key) : dflt;
}

int32_t JsNarrowNumericReader::i32(const emscripten::val& owner, const char* key, int32_t dflt, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  return fetched ? i32_value(value, dflt, field != nullptr ? field : key) : dflt;
}

uint8_t JsNarrowNumericReader::u8(const emscripten::val& owner, const char* key, uint8_t dflt, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  return fetched ? u8_value(value, dflt, field != nullptr ? field : key) : dflt;
}

uint16_t JsNarrowNumericReader::u16(const emscripten::val& owner, const char* key, uint16_t dflt, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  return fetched ? u16_value(value, dflt, field != nullptr ? field : key) : dflt;
}

uint8_t JsNarrowNumericReader::u8_value(const emscripten::val& value, uint8_t dflt, const char* field) {
  return static_cast<uint8_t>(integer_value(value, dflt, field, 0.0, 255.0, "[0, 255]"));
}

uint16_t JsNarrowNumericReader::u16_value(const emscripten::val& value, uint16_t dflt, const char* field) {
  return static_cast<uint16_t>(integer_value(value, dflt, field, 0.0, 65535.0, "[0, 65535]"));
}

uint32_t JsNarrowNumericReader::u32_value(const emscripten::val& value, uint32_t dflt, const char* field) {
  return static_cast<uint32_t>(integer_value(value, dflt, field, 0.0, kMaxU32, "[0, 4294967295]"));
}

std::string JsNarrowNumericReader::required_string(const emscripten::val& owner, const char* key, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  if (!fetched) {
    return std::string();
  }
  if (!value.isString()) {
    reject_type(field, "a string");
    return std::string();
  }
  return string_value(value, field);
}

void JsNarrowNumericReader::fail_type(const char* field, const char* expected) {
  reject_type(field, expected);
}

void JsNarrowNumericReader::fail_range(const char* field, const char* range) {
  reject(nullptr, field, range);
}

uint32_t JsNarrowNumericReader::required_u32(const emscripten::val& owner, const char* key, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  if (!fetched) {
    return 0U;
  }
  if (!js_value_present(value)) {
    reject_type(field, "a whole Number");
    return 0U;
  }
  if (!present_number(value, key, field)) {
    return 0U;
  }
  const double number = value.as<double>();
  if (!std::isfinite(number) || std::floor(number) != number) {
    reject_type(field, "a whole Number");
    return 0U;
  }
  if (number < 0.0 || number > kMaxU32) {
    reject(key, field, "[0, 4294967295]");
    return 0U;
  }
  return static_cast<uint32_t>(number);
}

int32_t JsNarrowNumericReader::i32_value(const emscripten::val& value, int32_t dflt, const char* field) {
  return static_cast<int32_t>(
      integer_value(value, dflt, field, -2147483648.0, 2147483647.0, "[-2147483648, 2147483647]"));
}

int64_t JsNarrowNumericReader::i64(const emscripten::val& owner, const char* key, int64_t dflt, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  if (!fetched) {
    return dflt;
  }
  double number = 0.0;
  if (!read_integer(value, key, field, -0x1p63, 0x1p63, "[-2^63, 2^63)", &number)) {
    return dflt;
  }
  if (number >= 0x1p63) {
    reject(key, field, "[-2^63, 2^63)");
    return dflt;
  }
  // `-2^63` is exactly representable as a JS Number, but converting that
  // boundary through a floating-point cast is implementation-defined on
  // some targets. Spell out the one value that needs the full signed range.
  if (number == -0x1p63) {
    return std::numeric_limits<int64_t>::min();
  }
  return static_cast<int64_t>(number);
}

double JsNarrowNumericReader::number(const emscripten::val& owner, const char* key, double dflt, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  if (!fetched) {
    return dflt;
  }
  return number_value(value, dflt, field != nullptr ? field : key);
}

double JsNarrowNumericReader::number_value(const emscripten::val& value, double dflt, const char* field) {
  if (!ok()) {
    return dflt;
  }
  double number = 0.0;
  if (!read_number(value, field, field, &number)) {
    return !js_value_present(value) ? dflt : 0.0;
  }
  return number;
}

bool JsNarrowNumericReader::boolean(const emscripten::val& owner, const char* key, bool dflt, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  if (!fetched) {
    return dflt;
  }
  return boolean_value(value, dflt);
}

bool JsNarrowNumericReader::boolean_value(const emscripten::val& value, bool dflt) {
  if (!ok()) {
    return dflt;
  }
  if (!js_value_present(value)) {
    return dflt;
  }
  // JavaScript `!` is a primitive truthiness operation and never performs
  // user-defined valueOf/toString conversion, including for Symbol values.
  return !value.operator!();
}

std::string JsNarrowNumericReader::string(const emscripten::val& owner, const char* key, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  if (!fetched) {
    return std::string();
  }
  return string_value(value, field != nullptr ? field : key);
}

std::string JsNarrowNumericReader::string_value(const emscripten::val& value, const char* field) {
  if (!ok() || !js_value_present(value)) {
    return std::string();
  }
  bool completed = false;
  const emscripten::val safe_value =
      safe_operation(value, emscripten::val::undefined(), kSafeStringSnapshot, &completed);
  if (!completed) {
    reject_access(nullptr, field);
    return std::string();
  }
  if (!js_value_present(safe_value)) {
    reject_type(field, "a string or one-byte buffer");
    return std::string();
  }
  return safe_value.as<std::string>();
}

const char* JsNarrowNumericReader::optional_string(const emscripten::val& owner, const char* key, std::string& storage,
                                                   const char* field) {
  if (!js_value_present(owner)) {
    return nullptr;
  }
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  if (!fetched) {
    return nullptr;
  }
  return optional_string_value(value, storage, field != nullptr ? field : key);
}

const char* JsNarrowNumericReader::optional_string_value(const emscripten::val& value, std::string& storage,
                                                         const char* field) {
  if (!ok() || !js_value_present(value)) {
    return nullptr;
  }
  storage = string_value(value, field);
  return ok() ? storage.c_str() : nullptr;
}

emscripten::val JsNarrowNumericReader::value(const emscripten::val& owner, const char* key, const char* field) {
  if (!ok()) {
    return emscripten::val::undefined();
  }
  if (!js_value_present(owner)) {
    reject_type(field != nullptr ? field : key, "an object");
    return emscripten::val::undefined();
  }
  const emscripten::val type = owner.typeOf();
  const std::string type_name = type.as<std::string>();
  if (type_name != "object" && type_name != "function") {
    reject_type(field != nullptr ? field : key, "an object");
    return emscripten::val::undefined();
  }
  bool completed = false;
  return safe_get(owner, emscripten::val(key), field != nullptr ? field : key, &completed);
}

bool JsNarrowNumericReader::safe_predicate(const emscripten::val& owner, const emscripten::val& key, int32_t operation,
                                           const char* key_name, const char* field) {
  if (!ok() || !js_value_present(owner)) {
    return false;
  }
  bool completed = false;
  const emscripten::val result = safe_operation(owner, key, operation, &completed);
  if (!completed) {
    reject_access(key_name, field);
    return false;
  }
  return result.isTrue();
}

bool JsNarrowNumericReader::has_own(const emscripten::val& owner, const char* key, const char* field) {
  return safe_predicate(owner, emscripten::val(key), kSafeHasOwn, key, field);
}

bool JsNarrowNumericReader::is_array(const emscripten::val& value, const char* field) {
  return safe_predicate(value, emscripten::val::undefined(), kSafeIsArray, nullptr, field);
}

uint32_t JsNarrowNumericReader::length(const emscripten::val& array, const char* field) {
  if (!ok()) {
    return 0U;
  }
  bool completed = false;
  const emscripten::val length_value = safe_get(array, emscripten::val("length"), field, &completed);
  if (!completed) {
    return 0U;
  }
  return u32_value(length_value, 0U, field != nullptr ? field : "length");
}

emscripten::val JsNarrowNumericReader::array_element(const emscripten::val& array, uint32_t index, const char* field) {
  if (!ok()) {
    return emscripten::val::undefined();
  }
  if (!require_array(array, field)) {
    return emscripten::val::undefined();
  }
  bool completed = false;
  return safe_get(array, emscripten::val(index), field, &completed);
}

bool JsNarrowNumericReader::require_array(const emscripten::val& value, const char* field) {
  if (is_array(value, field)) {
    return true;
  }
  if (ok()) {
    reject_type(field, "an array");
  }
  return false;
}

bool JsNarrowNumericReader::require_object(const emscripten::val& value, const char* field) {
  if (!ok()) {
    return false;
  }
  if (js_value_present(value) && value.typeOf().strictlyEquals(emscripten::val("object"))) {
    return true;
  }
  reject_type(field, "an object");
  return false;
}

bool JsNarrowNumericReader::ok() const {
  return message_.empty();
}

const std::string& JsNarrowNumericReader::message() const {
  return message_;
}

fm_status_t js_collect_list(const fm_workbook_t* handle, size_t sheet, JsListCount count, JsListItem item,
                            emscripten::val& list) {
  if (handle == nullptr) {
    return kBindingInvalidHandle;
  }
  size_t n = 0;
  fm_status_t rc = count(handle, sheet, &n);
  for (size_t i = 0; rc == 0 && i < n; ++i) {
    emscripten::val entry = emscripten::val::undefined();
    rc = item(handle, sheet, i, &entry);
    if (rc == 0) {
      list.set(i, entry);
    }
  }
  return rc;
}

double js_count(const fm_workbook_t* handle, size_t sheet, JsListCount count) {
  size_t n = 0;
  if (!succeeded(handle != nullptr ? count(handle, sheet, &n) : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(n);
}

emscripten::val value_to_val(const fm_value_t& v) {
  emscripten::val out = emscripten::val::object();
  out.set("kind", static_cast<int32_t>(v.kind));
  switch (v.kind) {
    case FM_VAL_NUMBER:
      out.set("number", v.u.number);
      break;
    case FM_VAL_BOOL:
      out.set("boolean", v.u.boolean != 0);
      break;
    case FM_VAL_TEXT:
      out.set("text", string_from_cstr(v.u.text));
      break;
    case FM_VAL_ERROR:
      out.set("errorCode", v.u.error_code);
      break;
    case FM_VAL_BLANK:
    case FM_VAL_ARRAY:
    case FM_VAL_REF:
    case FM_VAL_LAMBDA:
    default:
      // Other variants carry no payload across this boundary.
      break;
  }
  return out;
}

emscripten::val value_array_to_val(const fm_value_array_t* array) {
  uint32_t rows = 0;
  uint32_t cols = 0;
  fm_value_array_get_dimensions(array, &rows, &cols);
  emscripten::val cells = emscripten::val::array();
  for (uint32_t r = 0; r < rows; ++r) {
    emscripten::val row = emscripten::val::array();
    for (uint32_t c = 0; c < cols; ++c) {
      fm_value_t v{};
      fm_value_array_get_value_at(array, static_cast<std::size_t>(r) * cols + c, &v);
      row.set(c, value_to_val(v));
    }
    cells.set(r, row);
  }
  emscripten::val out = emscripten::val::object();
  out.set("rows", rows);
  out.set("cols", cols);
  out.set("cells", cells);
  return out;
}

JsCfMatch translate_cf_match(const fm_cf_match_t& m) {
  JsCfMatch out;
  out.kind = static_cast<int32_t>(m.kind);
  out.priority = m.priority;
  out.dxfIdEngaged = m.dxf_id_engaged;
  out.dxfId = m.dxf_id;
  out.color = js_color_spec(m.color);
  out.barLengthPct = m.bar_length_pct;
  out.barAxisPositionPct = m.bar_axis_position_pct;
  out.barIsNegative = m.bar_is_negative;
  out.barFill = js_color_spec(m.bar_fill);
  out.barBorderEngaged = m.bar_border_engaged;
  out.barBorder = js_color_spec(m.bar_border);
  out.barGradient = m.bar_gradient;
  out.barDirection = static_cast<int32_t>(m.bar_direction);
  out.iconSetName = m.icon_set_name;
  out.iconIndex = static_cast<int32_t>(m.icon_index);
  return out;
}

emscripten::val cell_rect_to_val(const fm_cell_rect_t& m) {
  emscripten::val item = emscripten::val::object();
  item.set("firstRow", m.first_row);
  item.set("firstCol", m.first_col);
  item.set("lastRow", m.last_row);
  item.set("lastCol", m.last_col);
  return item;
}

emscripten::val cell_rects_to_val(const fm_cell_rect_t* rects, std::size_t count) {
  emscripten::val arr = emscripten::val::array();
  for (std::size_t i = 0; i < count; ++i) {
    arr.call<void>("push", cell_rect_to_val(rects[i]));
  }
  return arr;
}

emscripten::val pivot_cell_to_val(const fm_pivot_cell_t& cell) {
  emscripten::val item = emscripten::val::object();
  item.set("row", cell.row);
  item.set("col", cell.col);
  item.set("value", value_to_val(cell.value));
  item.set("kind", static_cast<int32_t>(cell.kind));
  item.set("depth", cell.depth);
  js_set_cstr(item, "fieldName", cell.field_name);
  js_set_cstr_or_null(item, "numberFormat", cell.number_format);
  return item;
}

bool val_to_bytes_checked(const emscripten::val& v, std::vector<uint8_t>* out) {
  out->clear();
  int32_t completed = 0;
  const emscripten::val snapshot = emscripten::val::take_ownership(fm_wasm_safe_reader_operation(
      v.as_handle(), emscripten::val::undefined().as_handle(), kSafeBytesSnapshot, &completed));
  if (completed == 0) {
    fm_wasm_record_argument_error(kArgumentErrorAccess, "bytes: field access threw");
    return false;
  }
  if (!js_value_present(snapshot)) {
    fm_wasm_record_argument_error(kArgumentErrorType, "bytes must be a Uint8Array");
    return false;
  }

  JsNarrowNumericReader reader("bytes");
  const uint32_t length = reader.length(snapshot, "bytes.length");
  if (!reader.ok()) {
    return false;
  }
  out->resize(length);
  if (length == 0U) {
    return true;
  }
  uint32_t written = 0;
  const int32_t copied =
      fm_wasm_copy_safe_bytes(snapshot.as_handle(), reinterpret_cast<uintptr_t>(out->data()), length, &written);
  if (copied != 1 || written != length) {
    out->clear();
    // The copy trampoline reports a throw as -1 without capturing it, so
    // both outcomes surface as a type error on the argument.
    fm_wasm_record_argument_error(kArgumentErrorType,
                                  copied < 0 ? "bytes: field access threw" : "bytes must be a Uint8Array");
    return false;
  }
  return true;
}

emscripten::val bytes_to_val(const uint8_t* data, std::size_t len) {
  emscripten::val u8 = emscripten::val::global("Uint8Array").new_(len);
  if (len != 0) {
    // The destination is a standalone JS buffer; `set` copies the transient
    // WASM view before the caller releases its C++ storage.
    u8.call<void>("set", emscripten::val(emscripten::typed_memory_view(len, data)));
  }
  return u8;
}

emscripten::val snapshot_bytes(fm_status_t rc, uint8_t* data, std::size_t len) {
  emscripten::val bytes = succeeded(rc) ? bytes_to_val(data, len) : emscripten::val::undefined();
  fm_buffer_destroy(data);
  return bytes;
}

fm_cell_rect_t js_pull_range(const emscripten::val& v, JsNarrowNumericReader* reader) {
  if (reader == nullptr) {
    return fm_cell_rect_t{};
  }
  fm_cell_rect_t m{};
  if (!js_value_present(v)) {
    return m;
  }
  m.first_row = reader->required_u32(v, "firstRow", "range.firstRow");
  m.first_col = reader->required_u32(v, "firstCol", "range.firstCol");
  m.last_row = reader->required_u32(v, "lastRow", "range.lastRow");
  m.last_col = reader->required_u32(v, "lastCol", "range.lastCol");
  return m;
}

std::vector<fm_cell_rect_t> js_pull_ranges(const emscripten::val& v, const char* key, JsNarrowNumericReader* reader) {
  if (reader == nullptr) {
    return {};
  }
  const emscripten::val array = reader->value(v, key, key);
  if (!reader->ok()) {
    return {};
  }
  return js_pull_range_array(array, key, reader);
}

std::vector<fm_cell_rect_t> js_pull_range_array(const emscripten::val& array, const char* key,
                                                JsNarrowNumericReader* reader) {
  if (reader == nullptr) {
    return {};
  }
  std::vector<fm_cell_rect_t> out;
  if (!js_value_present(array)) {
    return out;
  }
  if (!reader->require_array(array, key)) {
    return out;
  }
  const uint32_t n = reader->length(array, key);
  if (!reader->ok()) {
    return out;
  }
  out.reserve(n);
  for (uint32_t i = 0; i < n; ++i) {
    const emscripten::val range = reader->array_element(array, i, "ranges[]");
    if (!reader->ok()) {
      return out;
    }
    if (!js_value_present(range)) {
      continue;
    }
    out.push_back(js_pull_range(range, reader));
  }
  return out;
}

fm_color_spec_t js_pull_color_spec(const emscripten::val& v, const char* key, JsNarrowNumericReader* reader) {
  if (reader == nullptr) {
    return fm_color_spec_t{};
  }
  const emscripten::val f = reader->value(v, key, key);
  return js_pull_color_spec_value(f, key, reader);
}

fm_color_spec_t js_pull_color_spec_value(const emscripten::val& f, const char* field, JsNarrowNumericReader* reader) {
  fm_color_spec_t spec{};
  if (reader == nullptr || !reader->ok() || !js_value_present(f)) {
    return spec;
  }
  if (!reader->require_object(f, field)) {
    return spec;
  }
  spec.kind = reader->u8(f, "kind", 0, "color.kind");
  spec.rgb = reader->u32(f, "rgb", 0U, "color.rgb");
  spec.theme = reader->u32(f, "theme", 0U, "color.theme");
  spec.tint = reader->number(f, "tint", 0.0, "color.tint");
  spec.indexed = reader->u32(f, "indexed", 0U, "color.indexed");
  return spec;
}

fm_border_side_t js_pull_border_side(const emscripten::val& v, JsNarrowNumericReader* reader) {
  fm_border_side_t s{};
  if (reader == nullptr || !js_value_present(v)) {
    return s;
  }
  s.style = reader->u8(v, "style", 0, "border.style");
  s.color = js_pull_color_spec(v, "color", reader);
  return s;
}

std::string string_from_cstr(const char* s) {
  return s != nullptr ? std::string(s) : std::string();
}

void js_set_cstr(emscripten::val& o, const char* key, const char* s) {
  o.set(key, string_from_cstr(s));
}

void js_set_cstr_fields(emscripten::val& o, const JsStrField* fields, std::size_t n) {
  for (std::size_t i = 0; i < n; ++i) {
    js_set_cstr(o, fields[i].key, fields[i].value);
  }
}

JsPatchKey js_patch_key(JsNarrowNumericReader& reader, const emscripten::val& value, const char* field, bool nullable) {
  if (!reader.ok() || value.isUndefined()) {
    return JsPatchKey::kKeep;
  }
  if (value.isNull()) {
    if (nullable) {
      return JsPatchKey::kClear;
    }
    reader.fail_type(field, "a value (null cannot clear this field)");
    return JsPatchKey::kKeep;
  }
  return JsPatchKey::kSet;
}

emscripten::val print_titles_to_val(const fm_print_titles_t& titles) {
  if (titles.has_rows == 0 && titles.has_cols == 0) {
    return emscripten::val::null();
  }
  const auto span = [](int32_t present, uint32_t first, uint32_t last) {
    if (present == 0) {
      return emscripten::val::null();
    }
    emscripten::val o = emscripten::val::object();
    o.set("first", first);
    o.set("last", last);
    return o;
  };
  emscripten::val out = emscripten::val::object();
  out.set("rows", span(titles.has_rows, titles.first_row, titles.last_row));
  out.set("columns", span(titles.has_cols, titles.first_col, titles.last_col));
  return out;
}

emscripten::val cstr_or_null(const char* s) {
  return s != nullptr && s[0] != '\0' ? emscripten::val(std::string(s)) : emscripten::val::null();
}

void js_set_cstr_or_null(emscripten::val& o, const char* key, const char* s) {
  o.set(key, cstr_or_null(s));
}

void js_set_cstr_or_null_fields(emscripten::val& o, const JsStrField* fields, std::size_t n) {
  for (std::size_t i = 0; i < n; ++i) {
    js_set_cstr_or_null(o, fields[i].key, fields[i].value);
  }
}

void js_set_index_or_null(emscripten::val& o, const char* key, int32_t value) {
  if (value >= 0) {
    o.set(key, value);
  } else {
    o.set(key, emscripten::val::null());
  }
}

emscripten::val js_color_spec(const fm_color_spec_t& spec) {
  emscripten::val o = emscripten::val::object();
  o.set("kind", static_cast<uint32_t>(spec.kind));
  o.set("rgb", spec.rgb);
  o.set("theme", spec.theme);
  o.set("tint", spec.tint);
  o.set("indexed", spec.indexed);
  return o;
}

emscripten::val js_color_spec_or_null(const fm_color_spec_t& spec) {
  return spec.kind != kFmColorNone ? js_color_spec(spec) : emscripten::val::null();
}

emscripten::val js_border_side(const fm_border_side_t& s) {
  emscripten::val o = emscripten::val::object();
  o.set("style", static_cast<uint32_t>(s.style));
  o.set("color", js_color_spec(s.color));
  return o;
}

namespace {

/// Records a `Value` type error and returns false.
bool reject_value(const char* message) {
  fm_wasm_record_argument_error(kArgumentErrorType, message);
  return false;
}

}  // namespace

bool val_to_value(const emscripten::val& in, fm_value_t* out, std::string& text) {
  *out = fm_value_t{};
  JsNarrowNumericReader reader("Value");
  if (!reader.require_object(in, "value")) {
    return false;
  }
  const emscripten::val kind = reader.value(in, "kind", "value.kind");
  if (!reader.ok()) {
    return false;
  }
  if (!kind.isNumber()) {
    return reject_value("Value: `value.kind` must be a ValueKind");
  }
  const double kind_number = kind.as<double>();
  if (!(kind_number >= static_cast<double>(FM_VAL_BLANK) && kind_number <= static_cast<double>(FM_VAL_LAMBDA)) ||
      std::floor(kind_number) != kind_number) {
    return reject_value("Value: `value.kind` must be a ValueKind");
  }
  out->kind = static_cast<int32_t>(kind_number);
  switch (out->kind) {
    case FM_VAL_NUMBER: {
      const emscripten::val number = reader.value(in, "number", "value.number");
      if (!reader.ok()) {
        return false;
      }
      if (!number.isNumber()) {
        return reject_value("Value: `value.number` must be a Number");
      }
      out->u.number = number.as<double>();
      return true;
    }
    case FM_VAL_BOOL: {
      const emscripten::val flag = reader.value(in, "boolean", "value.boolean");
      if (!reader.ok()) {
        return false;
      }
      if (!flag.isTrue() && !flag.isFalse()) {
        return reject_value("Value: `value.boolean` must be a boolean");
      }
      out->u.boolean = flag.isTrue() ? 1 : 0;
      return true;
    }
    case FM_VAL_TEXT: {
      const emscripten::val string = reader.value(in, "text", "value.text");
      if (!reader.ok()) {
        return false;
      }
      if (!string.isString()) {
        return reject_value("Value: `value.text` must be a string");
      }
      text = reader.string_value(string, "value.text");
      if (!reader.ok()) {
        return false;
      }
      out->u.text = text.c_str();
      return true;
    }
    case FM_VAL_ERROR: {
      const emscripten::val code = reader.value(in, "errorCode", "value.errorCode");
      if (!reader.ok()) {
        return false;
      }
      if (!code.isNumber()) {
        return reject_value("Value: `value.errorCode` must be an ErrorCode");
      }
      out->u.error_code = reader.i32_value(code, 0, "value.errorCode");
      return reader.ok();
    }
    default:
      // The remaining kinds carry no payload.
      return true;
  }
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
