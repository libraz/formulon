//
// Out-of-line implementation of the shared embind helpers declared in
// `parts/embind_common.h`. The status builders capture the thread-local
// diagnostic surface that follows every C-ABI call, the translation
// helpers project a `fm_*` POD into an embind-friendly mirror, and the
// field readers pull caller-supplied JS records apart.

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
// `undefined` handle. The body is JavaScript source; keep formatting disabled
// so clang-format cannot rewrite its operators.
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
        result = owner;
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
      if (owner instanceof Uint8Array) {
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
// RangeError, 2 rethrows the caller's getter exception. Outside a wrapped call
// there is no wrapper to throw, so the envelope's status stays the report.
// clang-format off
EM_JS(void, fm_wasm_record_argument_error, (int32_t kind, const char* message), {
  var errors = Module[Symbol.for('formulon.argumentErrors')];
  if (!errors || errors.depth === 0 || errors.pending !== undefined) {
    return;
  }
  var thrown = kind === 2 ? errors.thrown : undefined;
  errors.thrown = undefined;
  errors.pending = { kind: kind, message: UTF8ToString(message), thrown: thrown };
});
// clang-format on

EM_JS_DEPS(fm_wasm_argument_error, "$UTF8ToString");

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

JsStatus ok_status() {
  return JsStatus{true, 0, std::string(), std::string()};
}

JsStatus binding_error_status(int32_t code, const char* message) {
  JsStatus s;
  s.ok = false;
  s.status = code;
  s.message = string_from_cstr(message);
  return s;
}

JsStatus error_status(int32_t code) {
  if (code == kBindingInvalidHandle) {
    // No C-ABI call ran, so the thread-local diagnostics still describe
    // whatever call came before. Synthesise the message here instead.
    return binding_error_status(code, "the workbook handle was already destroyed");
  }
  JsStatus s;
  s.ok = false;
  s.status = code;
  s.message = string_from_cstr(fm_last_error_message());
  s.context = string_from_cstr(fm_last_error_context());
  return s;
}

JsStatus status_from_rc(fm_status_t rc) {
  return rc == 0 ? ok_status() : error_status(rc);
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

bool JsNarrowNumericReader::read_integer(const emscripten::val& value, const char* key, const char* field, double lower,
                                         double upper, const char* range, double* number) {
  if (value.isUndefined() || value.isNull()) {
    return false;
  }
  // Check the JavaScript primitive before asking embind for a C++ value.
  // In particular, Symbol and BigInt must never reach val::as on this path.
  if (!value.isNumber()) {
    reject_type(field != nullptr ? field : key, "a Number");
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
  if (value.isUndefined() || value.isNull()) {
    return false;
  }
  if (!value.isNumber()) {
    reject_type(field != nullptr ? field : key, "a Number");
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
    return (value.isUndefined() || value.isNull()) ? dflt : 0.0;
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
  return static_cast<uint32_t>(integer_value(value, dflt, field, 0.0, 4294967295.0, "[0, 4294967295]"));
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
    return (value.isUndefined() || value.isNull()) ? dflt : 0.0;
  }
  return number;
}

bool JsNarrowNumericReader::boolean(const emscripten::val& owner, const char* key, bool dflt, const char* field) {
  bool fetched = false;
  const emscripten::val value = fetch(owner, key, field, &fetched);
  if (!fetched) {
    return dflt;
  }
  return boolean_value(value, dflt, field != nullptr ? field : key);
}

bool JsNarrowNumericReader::boolean_value(const emscripten::val& value, bool dflt, const char* field) {
  if (!ok()) {
    return dflt;
  }
  (void)field;
  if (value.isUndefined() || value.isNull()) {
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
  if (!ok() || value.isUndefined() || value.isNull()) {
    return std::string();
  }
  bool completed = false;
  const emscripten::val safe_value =
      safe_operation(value, emscripten::val::undefined(), kSafeStringSnapshot, &completed);
  if (!completed) {
    reject_access(nullptr, field);
    return std::string();
  }
  if (safe_value.isUndefined() || safe_value.isNull()) {
    reject_type(field, "a string or one-byte buffer");
    return std::string();
  }
  return safe_value.as<std::string>();
}

const char* JsNarrowNumericReader::optional_string(const emscripten::val& owner, const char* key, std::string& storage,
                                                   const char* field) {
  if (owner.isUndefined() || owner.isNull()) {
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
  if (!ok() || value.isUndefined() || value.isNull()) {
    return nullptr;
  }
  storage = string_value(value, field);
  return ok() ? storage.c_str() : nullptr;
}

emscripten::val JsNarrowNumericReader::value(const emscripten::val& owner, const char* key, const char* field) {
  if (!ok()) {
    return emscripten::val::undefined();
  }
  if (owner.isUndefined() || owner.isNull()) {
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
  if (!ok() || owner.isUndefined() || owner.isNull()) {
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
  if (!is_array(array, field)) {
    if (!ok()) {
      return emscripten::val::undefined();
    }
    reject_type(field, "an array");
    return emscripten::val::undefined();
  }
  bool completed = false;
  return safe_get(array, emscripten::val(index), field, &completed);
}

bool JsNarrowNumericReader::ok() const {
  return message_.empty();
}

const std::string& JsNarrowNumericReader::message() const {
  return message_;
}

JsNumberResult number_result(fm_status_t rc, double value) {
  JsNumberResult out;
  out.status = status_from_rc(rc);
  if (rc == 0) {
    out.value = value;
  }
  return out;
}

JsStringResult string_result(fm_status_t rc, const char* value) {
  JsStringResult out;
  out.status = status_from_rc(rc);
  if (rc == 0 && value != nullptr) {
    out.value = value;
  }
  return out;
}

JsAddStyleResult index_result(fm_status_t rc, std::size_t index) {
  JsAddStyleResult out;
  out.status = status_from_rc(rc);
  if (rc == 0) {
    out.index = static_cast<uint32_t>(index);
  }
  return out;
}

JsValue translate_value(const fm_value_t& v) {
  JsValue out;
  out.kind = static_cast<int32_t>(v.kind);
  switch (v.kind) {
    case FM_VAL_NUMBER:
      out.number = v.u.number;
      break;
    case FM_VAL_BOOL:
      out.boolean = v.u.boolean;
      break;
    case FM_VAL_TEXT:
      out.text = (v.u.text != nullptr) ? std::string(v.u.text) : std::string();
      break;
    case FM_VAL_ERROR:
      out.errorCode = v.u.error_code;
      break;
    case FM_VAL_BLANK:
    case FM_VAL_ARRAY:
    case FM_VAL_REF:
    case FM_VAL_LAMBDA:
    default:
      // Other variants carry no scalar payload across this boundary.
      break;
  }
  return out;
}

JsCfColor translate_cf_color(const fm_cf_color_t& c) {
  JsCfColor out;
  out.r = static_cast<int32_t>(c.r);
  out.g = static_cast<int32_t>(c.g);
  out.b = static_cast<int32_t>(c.b);
  out.a = static_cast<int32_t>(c.a);
  return out;
}

JsCfMatch translate_cf_match(const fm_cf_match_t& m) {
  JsCfMatch out;
  out.kind = static_cast<int32_t>(m.kind);
  out.priority = m.priority;
  out.dxfIdEngaged = m.dxf_id_engaged;
  out.dxfId = m.dxf_id;
  out.color = translate_cf_color(m.color);
  out.barLengthPct = m.bar_length_pct;
  out.barAxisPositionPct = m.bar_axis_position_pct;
  out.barIsNegative = m.bar_is_negative;
  out.barFill = translate_cf_color(m.bar_fill);
  out.barBorderEngaged = m.bar_border_engaged;
  out.barBorder = translate_cf_color(m.bar_border);
  out.barGradient = m.bar_gradient;
  out.barDirection = static_cast<int32_t>(m.bar_direction);
  out.iconSetName = m.icon_set_name;
  out.iconIndex = static_cast<int32_t>(m.icon_index);
  return out;
}

emscripten::val merge_range_to_val(const fm_merge_range& m) {
  emscripten::val item = emscripten::val::object();
  item.set("firstRow", m.first_row);
  item.set("lastRow", m.last_row);
  item.set("firstCol", m.first_col);
  item.set("lastCol", m.last_col);
  return item;
}

emscripten::val empty_pivot_layout_result(JsStatus status) {
  emscripten::val o = emscripten::val::object();
  o.set("status", status);
  o.set("top", static_cast<uint32_t>(0));
  o.set("left", static_cast<uint32_t>(0));
  o.set("rows", static_cast<uint32_t>(0));
  o.set("cols", static_cast<uint32_t>(0));
  o.set("cells", emscripten::val::array());
  return o;
}

emscripten::val pivot_cell_to_val(const fm_pivot_cell_t& cell) {
  emscripten::val item = emscripten::val::object();
  item.set("row", cell.row);
  item.set("col", cell.col);
  item.set("value", translate_value(cell.value));
  item.set("kind", static_cast<int32_t>(cell.kind));
  item.set("depth", cell.depth);
  js_set_cstr(item, "fieldName", cell.field_name);
  js_set_cstr(item, "numberFormat", cell.number_format);
  return item;
}

JsBytesReadResult val_to_bytes_checked(const emscripten::val& v) {
  JsBytesReadResult result;
  int32_t completed = 0;
  const emscripten::val snapshot = emscripten::val::take_ownership(fm_wasm_safe_reader_operation(
      v.as_handle(), emscripten::val::undefined().as_handle(), kSafeBytesSnapshot, &completed));
  if (completed == 0) {
    result.ok = false;
    result.message = "bytes: field access threw";
    return result;
  }
  if (snapshot.isUndefined() || snapshot.isNull()) {
    result.ok = false;
    result.message = "bytes must be a Uint8Array";
    return result;
  }

  JsNarrowNumericReader reader("bytes");
  const uint32_t length = reader.length(snapshot, "bytes.length");
  if (!reader.ok()) {
    result.ok = false;
    result.message = reader.message();
    return result;
  }
  result.bytes.resize(length);
  if (length == 0U) {
    return result;
  }
  uint32_t written = 0;
  const int32_t copied =
      fm_wasm_copy_safe_bytes(snapshot.as_handle(), reinterpret_cast<uintptr_t>(result.bytes.data()), length, &written);
  if (copied != 1 || written != length) {
    result.ok = false;
    result.bytes.clear();
    result.message = copied < 0 ? "bytes: field access threw" : "bytes must be a Uint8Array";
  }
  return result;
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

fm_merge_range js_pull_range(const emscripten::val& v, JsNarrowNumericReader* reader) {
  if (reader == nullptr) {
    return fm_merge_range{};
  }
  fm_merge_range m{};
  if (v.isUndefined() || v.isNull()) {
    return m;
  }
  m.first_row = reader->u32_value(reader->value(v, "firstRow", "range.firstRow"), 0U, "range.firstRow");
  m.last_row = reader->u32_value(reader->value(v, "lastRow", "range.lastRow"), 0U, "range.lastRow");
  m.first_col = reader->u32_value(reader->value(v, "firstCol", "range.firstCol"), 0U, "range.firstCol");
  m.last_col = reader->u32_value(reader->value(v, "lastCol", "range.lastCol"), 0U, "range.lastCol");
  return m;
}

std::vector<fm_merge_range> js_pull_ranges(const emscripten::val& v, const char* key, JsNarrowNumericReader* reader) {
  if (reader == nullptr) {
    return {};
  }
  std::vector<fm_merge_range> out;
  const emscripten::val array = reader->value(v, key, key);
  if (!reader->ok() || array.isUndefined() || array.isNull()) {
    return out;
  }
  if (!reader->is_array(array, key)) {
    return out;
  }
  const uint32_t n = reader->length(array, key);
  if (!reader->ok()) {
    return out;
  }
  out.reserve(n);
  for (uint32_t i = 0; i < n; ++i) {
    const emscripten::val range = reader->array_element(array, i, "sqref[]");
    if (!reader->ok()) {
      return out;
    }
    if (range.isUndefined() || range.isNull()) {
      continue;
    }
    fm_merge_range m{};
    m.first_row = reader->u32(range, "firstRow", 0U, "sqref.firstRow");
    m.last_row = reader->u32(range, "lastRow", 0U, "sqref.lastRow");
    m.first_col = reader->u32(range, "firstCol", 0U, "sqref.firstCol");
    m.last_col = reader->u32(range, "lastCol", 0U, "sqref.lastCol");
    out.push_back(m);
  }
  return out;
}

fm_color_spec js_pull_color_spec(const emscripten::val& v, const char* key, JsNarrowNumericReader* reader) {
  fm_color_spec spec{};
  if (reader == nullptr) {
    return spec;
  }
  const emscripten::val f = reader->value(v, key, key);
  if (!reader->ok()) {
    return spec;
  }
  if (f.isUndefined() || f.isNull()) {
    return spec;
  }
  spec.kind = reader->u8(f, "kind", 0, "color.kind");
  spec.rgb = reader->u32(f, "rgb", 0U, "color.rgb");
  spec.theme = reader->u32(f, "theme", 0U, "color.theme");
  spec.tint = reader->number(f, "tint", 0.0, "color.tint");
  spec.indexed = reader->u32(f, "indexed", 0U, "color.indexed");
  return spec;
}

fm_border_side js_pull_border_side(const emscripten::val& v, JsNarrowNumericReader* reader) {
  fm_border_side s{};
  if (reader == nullptr || v.isUndefined() || v.isNull()) {
    return s;
  }
  s.style = reader->u8(v, "style", 0, "border.style");
  s.color_argb = reader->u32(v, "colorArgb", 0U, "border.colorArgb");
  s.color = js_pull_color_spec(v, "color", reader);
  return s;
}

std::string string_from_cstr(const char* s) {
  return s != nullptr ? std::string(s) : std::string();
}

void js_set_cstr(emscripten::val& o, const char* key, const char* s) {
  o.set(key, string_from_cstr(s));
}

emscripten::val js_text_result(fm_status_t rc, const char* key, const char* text) {
  emscripten::val out = emscripten::val::object();
  out.set("status", status_from_rc(rc));
  js_set_cstr(out, key, rc == 0 ? text : nullptr);
  return out;
}

void js_set_cstr_fields(emscripten::val& o, const JsStrField* fields, std::size_t n) {
  for (std::size_t i = 0; i < n; ++i) {
    js_set_cstr(o, fields[i].key, fields[i].value);
  }
}

emscripten::val js_color_spec(const fm_color_spec& spec) {
  emscripten::val o = emscripten::val::object();
  o.set("kind", static_cast<uint32_t>(spec.kind));
  o.set("rgb", spec.rgb);
  o.set("theme", spec.theme);
  o.set("tint", spec.tint);
  o.set("indexed", spec.indexed);
  return o;
}

emscripten::val js_border_side(const fm_border_side& s) {
  emscripten::val o = emscripten::val::object();
  o.set("style", static_cast<uint32_t>(s.style));
  o.set("colorArgb", s.color_argb);
  o.set("color", js_color_spec(s.color));
  return o;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
