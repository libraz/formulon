// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
  fm_cell_xf xf{};
  fm_status_t rc = fm_styles_get_cell_xf(handle_, xf_index, &xf);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("fontIndex", xf.font_index);
  o.set("fillIndex", xf.fill_index);
  o.set("borderIndex", xf.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.num_fmt_id));
  o.set("horizontalAlign", static_cast<uint32_t>(xf.horizontal_align));
  o.set("verticalAlign", static_cast<uint32_t>(xf.vertical_align));
  o.set("wrapText", xf.wrap_text != 0);
  return o;
}

emscripten::val JsWorkbook::getFont(uint32_t font_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_font_record f{};
  fm_status_t rc = fm_styles_get_font(handle_, font_index, &f);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("name", std::string(f.name != nullptr ? f.name : ""));
  o.set("size", f.size);
  o.set("colorArgb", f.color_argb);
  o.set("bold", f.bold != 0);
  o.set("italic", f.italic != 0);
  o.set("strike", f.strike != 0);
  o.set("underline", static_cast<uint32_t>(f.underline));
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

// ---- Style record adders -----------------------------------------------

JsAddStyleResult JsWorkbook::addFont(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  const std::string name = js_pull_string(record, "name");
  fm_font_record fr{};
  fr.name = name.c_str();
  fr.size = js_pull_double(record, "size", 11.0);
  fr.bold = js_pull_bool(record, "bold", false) ? 1 : 0;
  fr.italic = js_pull_bool(record, "italic", false) ? 1 : 0;
  fr.strike = js_pull_bool(record, "strike", false) ? 1 : 0;
  fr.underline = js_pull_u8(record, "underline", 0U);
  fr.color_argb = js_pull_u32(record, "colorArgb", 0xFF000000U);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_font(handle_, fr, &idx);
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
  fm_cell_xf xf{};
  xf.font_index = js_pull_u32(record, "fontIndex", 0U);
  xf.fill_index = js_pull_u32(record, "fillIndex", 0U);
  xf.border_index = js_pull_u32(record, "borderIndex", 0U);
  xf.num_fmt_id = js_pull_u16(record, "numFmtId", 0U);
  xf.horizontal_align = js_pull_u8(record, "horizontalAlign", 0U);
  xf.vertical_align = js_pull_u8(record, "verticalAlign", 0U);
  xf.wrap_text = js_pull_bool(record, "wrapText", false) ? 1 : 0;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_cell_xf(handle_, xf, &idx);
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
  fm_cell_xf xf{};
  fm_status_t rc = fm_styles_get_cell_style_xf(handle_, index, &xf);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("fontIndex", xf.font_index);
  o.set("fillIndex", xf.fill_index);
  o.set("borderIndex", xf.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.num_fmt_id));
  o.set("horizontalAlign", static_cast<uint32_t>(xf.horizontal_align));
  o.set("verticalAlign", static_cast<uint32_t>(xf.vertical_align));
  o.set("wrapText", xf.wrap_text != 0);
  return o;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
