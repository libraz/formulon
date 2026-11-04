//
// JsWorkbook reference-text surface: `parseReference` / `formatReference`,
// the A1 / R1C1 grammar `INDIRECT` reads, over `fm_workbook_parse_reference`
// and `fm_workbook_format_reference`. Both return plain values, so an invalid
// argument is raised as a TypeError / RangeError by the pre-js method wrapper
// rather than reported through a status envelope.

#include <emscripten/em_js.h>
#include <emscripten/val.h>

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

enum ReferenceErrorKind : int32_t {
  kReferenceTypeError = 0,
  kReferenceRangeError = 1,
};

// Hands an error to the pre-js method wrapper, which throws it after the
// native method returns (the same channel the numeric readers use); outside a
// wrapped call there is no wrapper and the call just returns `undefined`.
// clang-format off
EM_JS(void, fm_wasm_record_reference_error, (int32_t kind, const char* message), {
  var errors = Module[Symbol.for('formulon.argumentErrors')];
  if (!errors || errors.depth === 0 || errors.pending !== undefined) {
    return;
  }
  errors.pending = { kind: kind, message: UTF8ToString(message), thrown: undefined };
});
// clang-format on

EM_JS_DEPS(fm_wasm_reference_error, "$UTF8ToString");

emscripten::val reject(int32_t kind, const std::string& message) {
  fm_wasm_record_reference_error(kind, message.c_str());
  return emscripten::val::undefined();
}

emscripten::val reject_last_error() {
  return reject(kReferenceRangeError, fm_last_error_message());
}

/// `options` as an object (absent reads as `{}`); false after recording a TypeError for another type.
bool options_object(const emscripten::val& options, const char* api, emscripten::val* out) {
  *out = js_object_or_empty(options);
  if (out->typeOf().strictlyEquals(emscripten::val("object"))) {
    return true;
  }
  reject(kReferenceTypeError, std::string(api) + " expects an options object");
  return false;
}

/// A required uint32 property of `range`; a missing one is a TypeError.
uint32_t required_u32(JsNarrowNumericReader& reader, const emscripten::val& range, const char* key) {
  const emscripten::val v = reader.value(range, key, key);
  if (!reader.ok()) {
    return 0U;
  }
  if (!js_value_present(v)) {
    reject(kReferenceTypeError, std::string("formatReference: `") + key + "` is required");
    return 0U;
  }
  return reader.u32_value(v, 0U, key);
}

}  // namespace

emscripten::val JsWorkbook::parseReference(const std::string& text, emscripten::val options) const {
  using emscripten::val;
  if (handle_ == nullptr) {
    return reject(kReferenceTypeError, "the workbook handle was already destroyed");
  }
  val opts = val::undefined();
  if (!options_object(options, "parseReference", &opts)) {
    return val::undefined();
  }
  JsNarrowNumericReader reader("parseReference");
  const bool r1c1 = reader.boolean(opts, "r1c1", false, "r1c1");
  const val row_value = reader.value(opts, "baseRow", "baseRow");
  const val col_value = reader.value(opts, "baseCol", "baseCol");
  const bool has_row = js_value_present(row_value);
  const bool has_col = js_value_present(col_value);
  const uint32_t base_row = has_row ? reader.u32_value(row_value, 0U, "baseRow") : 0U;
  const uint32_t base_col = has_col ? reader.u32_value(col_value, 0U, "baseCol") : 0U;
  if (!reader.ok()) {
    return val::undefined();
  }
  if (has_row != has_col) {
    return reject(kReferenceTypeError, "parseReference: `baseRow` and `baseCol` must be given together");
  }
  fm_reference_t ref{};
  const fm_status_t rc =
      fm_workbook_parse_reference(handle_, text.c_str(), r1c1 ? 1 : 0, has_row ? 1 : 0, base_row, base_col, &ref);
  if (rc != 0) {
    return reject_last_error();
  }
  if (ref.valid == 0) {
    return val::null();
  }
  val out = val::object();
  out.set("sheetName", ref.has_sheet != 0 ? val(std::string(ref.sheet_name)) : val::null());
  out.set("row0", ref.row0);
  out.set("col0", ref.col0);
  out.set("row1", ref.row1);
  out.set("col1", ref.col1);
  return out;
}

emscripten::val JsWorkbook::formatReference(emscripten::val range, emscripten::val options) const {
  using emscripten::val;
  if (handle_ == nullptr) {
    return reject(kReferenceTypeError, "the workbook handle was already destroyed");
  }
  val opts = val::undefined();
  if (!options_object(options, "formatReference", &opts)) {
    return val::undefined();
  }
  JsNarrowNumericReader reader("formatReference");
  std::string sheet_store;
  const char* sheet_name = reader.optional_string(range, "sheetName", sheet_store, "sheetName");
  const uint32_t row0 = required_u32(reader, range, "row0");
  const uint32_t col0 = required_u32(reader, range, "col0");
  const uint32_t row1 = required_u32(reader, range, "row1");
  const uint32_t col1 = required_u32(reader, range, "col1");
  const bool absolute = reader.boolean(opts, "absolute", false, "absolute");
  if (!reader.ok()) {
    return val::undefined();
  }
  const char* text = nullptr;
  const fm_status_t rc =
      fm_workbook_format_reference(handle_, sheet_name, row0, col0, row1, col1, absolute ? 1 : 0, &text);
  if (rc != 0) {
    return reject_last_error();
  }
  return val(std::string(text));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
