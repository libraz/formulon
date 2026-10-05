//
// Out-of-line implementation of the shared embind helpers declared in
// `parts/embind_common.h`. The status builders capture the thread-local
// diagnostic surface that follows every C-ABI call, the translation
// helpers project a `fm_*` POD into an embind-friendly mirror, and the
// field readers pull caller-supplied JS records apart.

#include "wasm/parts/embind_common.h"

#include <emscripten/bind.h>
#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"

namespace formulon {
namespace wasm {
namespace parts {

JsStatus ok_status() {
  return JsStatus{true, 0, std::string(), std::string()};
}

JsStatus binding_error_status(int32_t code, const char* message) {
  JsStatus s;
  s.ok = false;
  s.status = code;
  s.message = message != nullptr ? message : "";
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
  const char* msg = fm_last_error_message();
  const char* ctx = fm_last_error_context();
  s.message = msg != nullptr ? msg : "";
  s.context = ctx != nullptr ? ctx : "";
  return s;
}

JsStatus status_from_rc(fm_status_t rc) {
  return rc == 0 ? ok_status() : error_status(rc);
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

std::vector<uint8_t> val_to_bytes(const emscripten::val& v) {
  const emscripten::val uint8_array = emscripten::val::global("Uint8Array");
  if (v.isNull() || v.isUndefined() || !v.instanceof (uint8_array)) {
    return {};
  }
  const std::size_t len = v["length"].as<std::size_t>();
  std::vector<uint8_t> out(len);
  if (len == 0) {
    return out;
  }
  // `typed_memory_view` gives JS a view of the vector's WASM allocation;
  // Uint8Array#set copies the whole input in one JS operation. This avoids
  // one embind boundary crossing per byte for workbook-sized inputs.
  emscripten::val(emscripten::typed_memory_view(len, out.data())).call<void>("set", v);
  return out;
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

bool js_has(const emscripten::val& v, const char* key) {
  const emscripten::val f = v[key];
  return !f.isUndefined() && !f.isNull();
}

uint32_t js_pull_u32(const emscripten::val& v, const char* key, uint32_t dflt) {
  const emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return dflt;
  }
  return f.as<uint32_t>();
}

int32_t js_pull_i32(const emscripten::val& v, const char* key, int32_t dflt) {
  const emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return dflt;
  }
  return f.as<int32_t>();
}

double js_pull_double(const emscripten::val& v, const char* key, double dflt) {
  const emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return dflt;
  }
  return f.as<double>();
}

bool js_pull_bool(const emscripten::val& v, const char* key, bool dflt) {
  const emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return dflt;
  }
  return f.as<bool>();
}

std::string js_pull_string(const emscripten::val& v, const char* key) {
  const emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return std::string();
  }
  return f.as<std::string>();
}

const char* js_pull_optional_string(const emscripten::val& v, const char* key, std::string& storage) {
  const emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return nullptr;
  }
  storage = f.as<std::string>();
  return storage.c_str();
}

uint32_t js_length(const emscripten::val& arr) {
  return arr["length"].as<uint32_t>();
}

fm_merge_range js_pull_range(const emscripten::val& v) {
  fm_merge_range m{};
  m.first_row = v["firstRow"].as<uint32_t>();
  m.last_row = v["lastRow"].as<uint32_t>();
  m.first_col = v["firstCol"].as<uint32_t>();
  m.last_col = v["lastCol"].as<uint32_t>();
  return m;
}

std::vector<fm_merge_range> js_pull_ranges(const emscripten::val& v, const char* key) {
  std::vector<fm_merge_range> out;
  if (!v.hasOwnProperty(key)) {
    return out;
  }
  const emscripten::val arr = v[key];
  if (!arr.isArray()) {
    return out;
  }
  const uint32_t n = js_length(arr);
  out.reserve(n);
  for (uint32_t i = 0; i < n; ++i) {
    out.push_back(js_pull_range(arr[i]));
  }
  return out;
}

std::vector<uint32_t> js_pull_u32_list(const emscripten::val& arr) {
  std::vector<uint32_t> out;
  if (arr.isUndefined() || arr.isNull()) {
    return out;
  }
  const uint32_t n = js_length(arr);
  out.reserve(n);
  for (uint32_t i = 0; i < n; ++i) {
    out.push_back(arr[i].as<uint32_t>());
  }
  return out;
}

bool js_pull_u32_array(const emscripten::val& arr, uint32_t* out, uint32_t n) {
  if (!arr.isArray() || js_length(arr) != n) {
    return false;
  }
  for (uint32_t i = 0; i < n; ++i) {
    out[i] = arr[i].as<uint32_t>();
  }
  return true;
}

fm_color_spec js_pull_color_spec(const emscripten::val& v, const char* key) {
  fm_color_spec spec{};
  const emscripten::val f = v[key];
  if (f.isUndefined() || f.isNull()) {
    return spec;
  }
  spec.kind = js_pull_u8(f, "kind", 0);
  spec.rgb = js_pull_u32(f, "rgb", 0U);
  spec.theme = js_pull_u32(f, "theme", 0U);
  spec.tint = js_pull_double(f, "tint", 0.0);
  spec.indexed = js_pull_u32(f, "indexed", 0U);
  return spec;
}

fm_border_side js_pull_border_side(const emscripten::val& v) {
  fm_border_side s{};
  if (v.isUndefined() || v.isNull()) {
    return s;
  }
  s.style = js_pull_u8(v, "style", 0);
  s.color_argb = js_pull_u32(v, "colorArgb", 0U);
  s.color = js_pull_color_spec(v, "color");
  return s;
}

void js_set_cstr(emscripten::val& o, const char* key, const char* s) {
  o.set(key, s != nullptr ? std::string(s) : std::string());
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
