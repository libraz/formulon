//
// JsWorkbook styles surface: per-cell xfIndex get/set, font / fill /
// border / numFmt / xf record getters and adders, named cell-style and
// cellStyleXfs accessors, plus the matching count accessors. The four
// add* paths share the `js_pull_*` helpers in `parts/embind_common.h`
// so the embind glue is emitted once per field type rather than per
// call site.

#include <emscripten/val.h>

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

// ---- Per-cell xf index get/set -----------------------------------------

emscripten::val JsWorkbook::getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  uint32_t xf = 0;
  fm_status_t rc = fm_cell_get_xf_index(handle_, sheet, row, col, &xf);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("xfIndex", xf);
  return o;
}

JsStatus JsWorkbook::setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_cell_set_xf_index(handle_, sheet, row, col, xf_index);
  return status_from_rc(rc);
}

// ---- Style record getters ----------------------------------------------

emscripten::val JsWorkbook::getCellXf(uint32_t xf_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_cell_xf_ex xf{};
  fm_status_t rc = fm_styles_get_cell_xf_ex(handle_, xf_index, &xf);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("fontIndex", xf.base.font_index);
  o.set("fillIndex", xf.base.fill_index);
  o.set("borderIndex", xf.base.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.base.num_fmt_id));
  o.set("horizontalAlign", static_cast<uint32_t>(xf.base.horizontal_align));
  o.set("verticalAlign", static_cast<uint32_t>(xf.base.vertical_align));
  o.set("wrapText", xf.base.wrap_text != 0);
  o.set("justifyLastLine", xf.justify_last_line != 0);
  o.set("xfId", xf.xf_id);
  return o;
}

emscripten::val JsWorkbook::getFont(uint32_t font_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_font_record_ex f{};
  fm_status_t rc = fm_styles_get_font_ex(handle_, font_index, &f);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("name", std::string(f.base.name != nullptr ? f.base.name : ""));
  o.set("size", f.base.size);
  o.set("colorArgb", f.base.color_argb);
  o.set("bold", f.base.bold != 0);
  o.set("italic", f.base.italic != 0);
  o.set("strike", f.base.strike != 0);
  o.set("underline", static_cast<uint32_t>(f.base.underline));
  o.set("vertAlign", static_cast<uint32_t>(f.vert_align));
  return o;
}

emscripten::val JsWorkbook::getFill(uint32_t fill_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_fill_record f{};
  fm_status_t rc = fm_styles_get_fill(handle_, fill_index, &f);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("pattern", static_cast<uint32_t>(f.pattern));
  o.set("fgArgb", f.fg_argb);
  o.set("bgArgb", f.bg_argb);
  return o;
}

emscripten::val JsWorkbook::getBorder(uint32_t border_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_border_record b{};
  fm_status_t rc = fm_styles_get_border(handle_, border_index, &b);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  auto side_obj = [](const fm_border_side& s) {
    emscripten::val v = emscripten::val::object();
    v.set("style", static_cast<uint32_t>(s.style));
    v.set("colorArgb", s.color_argb);
    return v;
  };
  o.set("status", ok_status());
  o.set("left", side_obj(b.left));
  o.set("right", side_obj(b.right));
  o.set("top", side_obj(b.top));
  o.set("bottom", side_obj(b.bottom));
  o.set("diagonal", side_obj(b.diagonal));
  o.set("diagonalUp", b.diagonal_up != 0);
  o.set("diagonalDown", b.diagonal_down != 0);
  return o;
}

emscripten::val JsWorkbook::getNumFmt(uint32_t num_fmt_id) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  const char* s = nullptr;
  fm_status_t rc = fm_styles_get_num_fmt_string(handle_, static_cast<uint16_t>(num_fmt_id), &s);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("numFmtId", num_fmt_id);
  o.set("formatCode", std::string(s != nullptr ? s : ""));
  return o;
}

emscripten::val JsWorkbook::getDxf(uint32_t dxf_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_dxf_record d{};
  fm_status_t rc = fm_styles_get_dxf(handle_, dxf_index, &d);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  if (d.font_engaged != 0) {
    emscripten::val font = emscripten::val::object();
    font.set("name", std::string(d.font.name != nullptr ? d.font.name : ""));
    font.set("size", d.font.size);
    font.set("colorArgb", d.font.color_argb);
    font.set("bold", d.font.bold != 0);
    font.set("italic", d.font.italic != 0);
    font.set("strike", d.font.strike != 0);
    font.set("underline", static_cast<uint32_t>(d.font.underline));
    o.set("font", font);
  }
  if (d.fill_engaged != 0) {
    emscripten::val fill = emscripten::val::object();
    fill.set("pattern", static_cast<uint32_t>(d.fill.pattern));
    fill.set("fgArgb", d.fill.fg_argb);
    fill.set("bgArgb", d.fill.bg_argb);
    o.set("fill", fill);
  }
  if (d.border_engaged != 0) {
    auto side_obj = [](const fm_border_side& s) {
      emscripten::val v = emscripten::val::object();
      v.set("style", static_cast<uint32_t>(s.style));
      v.set("colorArgb", s.color_argb);
      return v;
    };
    emscripten::val border = emscripten::val::object();
    border.set("left", side_obj(d.border.left));
    border.set("right", side_obj(d.border.right));
    border.set("top", side_obj(d.border.top));
    border.set("bottom", side_obj(d.border.bottom));
    border.set("diagonal", side_obj(d.border.diagonal));
    border.set("diagonalUp", d.border.diagonal_up != 0);
    border.set("diagonalDown", d.border.diagonal_down != 0);
    o.set("border", border);
  }
  if (d.num_fmt_engaged != 0) {
    emscripten::val num_fmt = emscripten::val::object();
    num_fmt.set("numFmtId", static_cast<uint32_t>(d.num_fmt_id));
    num_fmt.set("formatCode", std::string(d.num_fmt_code != nullptr ? d.num_fmt_code : ""));
    o.set("numFmt", num_fmt);
  }
  return o;
}

// ---- Style record adders -----------------------------------------------

JsAddStyleResult JsWorkbook::addFont(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  const std::string name = js_pull_string(record, "name");
  fm_font_record_ex fr{};
  fr.base.name = name.c_str();
  fr.base.size = js_pull_double(record, "size", 11.0);
  fr.base.bold = js_pull_bool(record, "bold", false) ? 1 : 0;
  fr.base.italic = js_pull_bool(record, "italic", false) ? 1 : 0;
  fr.base.strike = js_pull_bool(record, "strike", false) ? 1 : 0;
  fr.base.underline = js_pull_u8(record, "underline", 0U);
  fr.base.color_argb = js_pull_u32(record, "colorArgb", 0xFF000000U);
  fr.vert_align = js_pull_u8(record, "vertAlign", 0U);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_font_ex(handle_, fr, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

JsAddStyleResult JsWorkbook::addFill(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_fill_record fr{};
  fr.pattern = js_pull_u8(record, "pattern", 0U);
  fr.fg_argb = js_pull_u32(record, "fgArgb", 0U);
  fr.bg_argb = js_pull_u32(record, "bgArgb", 0U);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_fill(handle_, fr, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

JsAddStyleResult JsWorkbook::addBorder(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_border_record br{};
  br.left = js_pull_border_side(record["left"]);
  br.right = js_pull_border_side(record["right"]);
  br.top = js_pull_border_side(record["top"]);
  br.bottom = js_pull_border_side(record["bottom"]);
  br.diagonal = js_pull_border_side(record["diagonal"]);
  br.diagonal_up = js_pull_bool(record, "diagonalUp", false) ? 1 : 0;
  br.diagonal_down = js_pull_bool(record, "diagonalDown", false) ? 1 : 0;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_border(handle_, br, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

JsAddNumFmtResult JsWorkbook::addNumFmt(const std::string& format_code) {
  JsAddNumFmtResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  uint16_t id = 0;
  fm_status_t rc = fm_styles_add_num_fmt(handle_, format_code.c_str(), &id);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.numFmtId = id;
  return r;
}

JsAddStyleResult JsWorkbook::addXf(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_cell_xf_ex xf{};
  xf.base.font_index = js_pull_u32(record, "fontIndex", 0U);
  xf.base.fill_index = js_pull_u32(record, "fillIndex", 0U);
  xf.base.border_index = js_pull_u32(record, "borderIndex", 0U);
  xf.base.num_fmt_id = js_pull_u16(record, "numFmtId", 0U);
  xf.base.horizontal_align = js_pull_u8(record, "horizontalAlign", 0U);
  xf.base.vertical_align = js_pull_u8(record, "verticalAlign", 0U);
  xf.base.wrap_text = js_pull_bool(record, "wrapText", false) ? 1 : 0;
  xf.justify_last_line = js_pull_bool(record, "justifyLastLine", false) ? 1 : 0;
  xf.xf_id = js_pull_u32(record, "xfId", 0U);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_cell_xf_ex(handle_, xf, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

JsAddStyleResult JsWorkbook::addDxf(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }

  std::string font_name;
  std::string num_fmt_code;
  fm_dxf_record dxf{};

  emscripten::val font = record["font"];
  if (!font.isUndefined() && !font.isNull()) {
    dxf.font_engaged = 1;
    font_name = js_pull_string(font, "name");
    dxf.font.name = font_name.c_str();
    dxf.font.size = js_pull_double(font, "size", 11.0);
    dxf.font.bold = js_pull_bool(font, "bold", false) ? 1 : 0;
    dxf.font.italic = js_pull_bool(font, "italic", false) ? 1 : 0;
    dxf.font.strike = js_pull_bool(font, "strike", false) ? 1 : 0;
    dxf.font.underline = js_pull_u8(font, "underline", 0U);
    dxf.font.color_argb = js_pull_u32(font, "colorArgb", 0xFF000000U);
  }

  emscripten::val fill = record["fill"];
  if (!fill.isUndefined() && !fill.isNull()) {
    dxf.fill_engaged = 1;
    dxf.fill.pattern = js_pull_u8(fill, "pattern", 0U);
    dxf.fill.fg_argb = js_pull_u32(fill, "fgArgb", 0U);
    dxf.fill.bg_argb = js_pull_u32(fill, "bgArgb", 0U);
  }

  emscripten::val border = record["border"];
  if (!border.isUndefined() && !border.isNull()) {
    dxf.border_engaged = 1;
    dxf.border.left = js_pull_border_side(border["left"]);
    dxf.border.right = js_pull_border_side(border["right"]);
    dxf.border.top = js_pull_border_side(border["top"]);
    dxf.border.bottom = js_pull_border_side(border["bottom"]);
    dxf.border.diagonal = js_pull_border_side(border["diagonal"]);
    dxf.border.diagonal_up = js_pull_bool(border, "diagonalUp", false) ? 1 : 0;
    dxf.border.diagonal_down = js_pull_bool(border, "diagonalDown", false) ? 1 : 0;
  }

  emscripten::val num_fmt = record["numFmt"];
  if (!num_fmt.isUndefined() && !num_fmt.isNull()) {
    dxf.num_fmt_engaged = 1;
    dxf.num_fmt_id = js_pull_u16(num_fmt, "numFmtId", 0U);
    num_fmt_code = js_pull_string(num_fmt, "formatCode");
    dxf.num_fmt_code = num_fmt_code.c_str();
  }

  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_dxf(handle_, dxf, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

// ---- Style count accessors ---------------------------------------------
//
// `fontCount` / `fillCount` / `borderCount` / `xfCount` are now emitted
// by the binding codegen (see `src/wasm/generated/styles_counts.cpp`).

uint32_t JsWorkbook::dxfCount() const {
  if (handle_ == nullptr) {
    return 0;
  }
  uint32_t n = 0;
  if (fm_styles_get_dxf_count(handle_, &n) != 0) {
    return 0;
  }
  return n;
}
// `cellStyleCount` / `cellStyleXfCount` stay here because they have no
// N-API counterpart and are therefore not part of the cross-binding
// manifest.

uint32_t JsWorkbook::cellStyleCount() const {
  if (handle_ == nullptr) {
    return 0U;
  }
  uint32_t n = 0;
  if (fm_styles_get_cell_style_count(handle_, &n) != 0) {
    return 0U;
  }
  return n;
}

uint32_t JsWorkbook::cellStyleXfCount() const {
  if (handle_ == nullptr) {
    return 0U;
  }
  uint32_t n = 0;
  if (fm_styles_get_cell_style_xf_count(handle_, &n) != 0) {
    return 0U;
  }
  return n;
}

emscripten::val JsWorkbook::getCellStyle(uint32_t index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_cell_style_record_t cs{};
  fm_status_t rc = fm_styles_get_cell_style(handle_, index, &cs);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("name", std::string(cs.name != nullptr ? cs.name : ""));
  o.set("xfId", cs.xf_id);
  o.set("builtinId", cs.builtin_id);
  o.set("iLevel", cs.i_level);
  o.set("hidden", cs.hidden != 0);
  o.set("customBuiltin", cs.custom_builtin != 0);
  return o;
}

emscripten::val JsWorkbook::getCellStyleXf(uint32_t index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_cell_xf_ex xf{};
  fm_status_t rc = fm_styles_get_cell_style_xf_ex(handle_, index, &xf);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("fontIndex", xf.base.font_index);
  o.set("fillIndex", xf.base.fill_index);
  o.set("borderIndex", xf.base.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.base.num_fmt_id));
  o.set("horizontalAlign", static_cast<uint32_t>(xf.base.horizontal_align));
  o.set("verticalAlign", static_cast<uint32_t>(xf.base.vertical_align));
  o.set("wrapText", xf.base.wrap_text != 0);
  o.set("justifyLastLine", xf.justify_last_line != 0);
  return o;
}

JsAddStyleResult JsWorkbook::addCellStyleXf(emscripten::val record) {
  JsAddStyleResult out;
  if (handle_ == nullptr) {
    out.status = error_status(7000);
    return out;
  }
  fm_cell_xf_ex xf{};
  xf.base.font_index = js_pull_u32(record, "fontIndex", 0U);
  xf.base.fill_index = js_pull_u32(record, "fillIndex", 0U);
  xf.base.border_index = js_pull_u32(record, "borderIndex", 0U);
  xf.base.num_fmt_id = js_pull_u16(record, "numFmtId", 0U);
  xf.base.horizontal_align = js_pull_u8(record, "horizontalAlign", 0U);
  xf.base.vertical_align = js_pull_u8(record, "verticalAlign", 0U);
  xf.base.wrap_text = js_pull_bool(record, "wrapText", false) ? 1 : 0;
  xf.justify_last_line = js_pull_bool(record, "justifyLastLine", false) ? 1 : 0;
  uint32_t index = 0;
  const fm_status_t rc = fm_styles_add_cell_style_xf_ex(handle_, xf, &index);
  out.status = rc == 0 ? ok_status() : error_status(rc);
  out.index = index;
  return out;
}

JsStatus JsWorkbook::setCellStyle(const std::string& name, uint32_t xfId, uint32_t builtinId) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_styles_set_cell_style(handle_, name.c_str(), xfId, builtinId));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
