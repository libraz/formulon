//
// Out-of-line implementation of the shared embind helpers declared in
// `parts/embind_common.h`. The status builders capture the thread-local
// diagnostic surface that follows every C-ABI call, and the translation
// helpers project a `fm_*` POD into an embind-friendly mirror.

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

JsStatus error_status(int32_t code) {
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
  item.set("fieldName", cell.field_name != nullptr ? std::string(cell.field_name) : std::string());
  item.set("numberFormat", cell.number_format != nullptr ? std::string(cell.number_format) : std::string());
  return item;
}

std::vector<uint8_t> val_to_bytes(const emscripten::val& v) {
  if (v.isNull() || v.isUndefined()) {
    return {};
  }
  // `length` works for both Uint8Array and plain arrays; `byteLength`
  // is the canonical Uint8Array property. Prefer `length`.
  const auto len_val = v["length"];
  if (len_val.isUndefined()) {
    return {};
  }
  const std::size_t len = len_val.as<std::size_t>();
  std::vector<uint8_t> out(len);
  if (len == 0) {
    return out;
  }
  // `emscripten::val::vecFromJSArray` copies element-by-element via
  // `as<T>`. For a typed Uint8Array this is correct and avoids relying
  // on heap-pointer aliasing tricks that depend on internal layout.
  for (std::size_t i = 0; i < len; ++i) {
    out[i] = v[i].as<uint8_t>();
  }
  return out;
}

emscripten::val bytes_to_val(const uint8_t* data, std::size_t len) {
  emscripten::val u8 = emscripten::val::global("Uint8Array").new_(len);
  for (std::size_t i = 0; i < len; ++i) {
    u8.set(i, data[i]);
  }
  return u8;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
