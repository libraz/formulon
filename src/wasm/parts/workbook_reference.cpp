//
// JsWorkbook reference-text surface: `parseReference` / `formatReference`,
// the A1 / R1C1 grammar `INDIRECT` reads, over `fm_workbook_parse_reference`
// and `fm_workbook_format_reference`, and the sheet lookup by name. An
// argument JS cannot represent (a wrong type, a non-integer) is raised as a
// TypeError / RangeError by the pre-js method wrapper; an engine rejection
// is thrown there as a `FormulonError`.

#include <emscripten/em_js.h>
#include <emscripten/val.h>

#include <cstddef>
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
};

// Hands an error to the pre-js method wrapper, which throws it after the
// native method returns (the same channel the numeric readers use, so it
// displaces a recorded status failure); outside a wrapped call there is no
// wrapper and the call just returns `undefined`.
// clang-format off
EM_JS(void, fm_wasm_record_reference_error, (int32_t kind, const char* message), {
  var errors = Module[Symbol.for('formulon.argumentErrors')];
  if (!errors || errors.depth === 0 || (errors.pending !== undefined && errors.pending.kind !== 3)) {
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

/// `options` as an object (absent reads as `{}`); false after recording a TypeError for another type.
bool options_object(const emscripten::val& options, const char* api, emscripten::val* out) {
  *out = js_object_or_empty(options);
  if (out->typeOf().strictlyEquals(emscripten::val("object"))) {
    return true;
  }
  reject(kReferenceTypeError, std::string(api) + " expects an options object");
  return false;
}

}  // namespace

emscripten::val JsWorkbook::parseReference(const std::string& text, emscripten::val options) const {
  using emscripten::val;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return val::undefined();
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
  fm_parsed_reference_t ref = FM_STRUCT_INIT(fm_parsed_reference_t);
  fm_parse_reference_spec_t spec = FM_STRUCT_INIT(fm_parse_reference_spec_t);
  spec.r1c1 = r1c1 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  spec.has_base = has_row ? 1 : 0;
  spec.base_row = base_row;
  spec.base_col = base_col;
  const fm_status_t rc = fm_workbook_parse_reference(handle_, text.c_str(), &spec, &ref);
  if (rc != 0 || ref.valid == 0) {
    return optional_or_null(rc, 0, val::null());
  }
  val out = val::object();
  out.set("sheetName", ref.has_sheet != 0 ? val(std::string(ref.sheet_name)) : val::null());
  out.set("rect", cell_rect_to_val(ref.rect));
  return optional_or_null(rc, 1, out);
}

std::string JsWorkbook::formatReference(emscripten::val range, emscripten::val options) const {
  using emscripten::val;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return std::string();
  }
  val opts = val::undefined();
  if (!options_object(options, "formatReference", &opts)) {
    return std::string();
  }
  JsNarrowNumericReader reader("formatReference");
  std::string sheet_store;
  const char* sheet_name = reader.optional_string(range, "sheetName", sheet_store, "sheetName");
  const val rect_value = reader.value(range, "rect", "rect");
  if (reader.ok() && !js_value_present(rect_value)) {
    reject(kReferenceTypeError, "formatReference: `rect` is required");
    return std::string();
  }
  const fm_cell_rect_t rect = js_pull_range(rect_value, &reader);
  const bool absolute = reader.boolean(opts, "absolute", false, "absolute");
  if (!reader.ok()) {
    return std::string();
  }
  const char* text = nullptr;
  if (!succeeded(fm_workbook_format_reference(handle_, sheet_name, &rect, absolute ? 1 : 0, &text))) {
    return std::string();
  }
  return string_from_cstr(text);
}

emscripten::val JsWorkbook::getSheetIndex(const std::string& name) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  std::size_t index = 0;
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_sheet_index(handle_, name.c_str(), &index, &present);
  return optional_or_null(rc, present, emscripten::val(static_cast<uint32_t>(index)));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
