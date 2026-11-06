//
// JsWorkbook named cell-style accessors (list / get / set / remove) and
// table-style accessors (list / get / set / remove, workbook defaults).

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

emscripten::val cell_style_to_val(const fm_cell_style_record_t& cs) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", cs.name);
  o.set("xfId", cs.xf_id);
  o.set("builtinId", cs.builtin_id);
  o.set("iLevel", cs.i_level);
  o.set("hidden", cs.hidden != 0);
  o.set("customBuiltin", cs.custom_builtin != 0);
  return o;
}

emscripten::val table_style_to_val(const fm_table_style_t& style) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", style.name);
  o.set("pivot", style.pivot != 0);
  o.set("table", style.table != 0);
  emscripten::val elements = emscripten::val::array();
  for (uint32_t i = 0; i < style.element_count; ++i) {
    const fm_table_style_element_t& element = style.elements[i];
    emscripten::val e = emscripten::val::object();
    js_set_cstr(e, "type", element.type);
    e.set("size", element.size);
    if (element.has_dxf_id != 0) {
      e.set("dxfId", element.dxf_id);
    }
    elements.call<void>("push", e);
  }
  o.set("elements", elements);
  return o;
}

// Style tables are workbook-scoped; the sheet argument is ignored.
fm_status_t count_cell_styles(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_cell_style_count(handle, out);
}

fm_status_t cell_style_item(const fm_workbook_t* handle, size_t, size_t index, emscripten::val* out) {
  fm_cell_style_record_t cs{};
  const fm_status_t rc = fm_workbook_get_cell_style_at(handle, index, &cs);
  if (rc == 0) {
    *out = cell_style_to_val(cs);
  }
  return rc;
}

fm_status_t count_table_styles(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_table_style_count(handle, out);
}

fm_status_t table_style_item(const fm_workbook_t* handle, size_t, size_t index, emscripten::val* out) {
  fm_table_style_t style{};
  const fm_status_t rc = fm_workbook_get_table_style_at(handle, index, &style);
  if (rc == 0) {
    *out = table_style_to_val(style);
  }
  return rc;
}

}  // namespace

emscripten::val JsWorkbook::listCellStyles() const {
  emscripten::val arr = emscripten::val::array();
  if (!succeeded(js_collect_list(handle_, 0U, &count_cell_styles, &cell_style_item, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

emscripten::val JsWorkbook::getCellStyle(const std::string& name) const {
  fm_cell_style_record_t cs{};
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_cell_style(handle_, name.c_str(), &present, &cs) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? cell_style_to_val(cs) : emscripten::val::null());
}

void JsWorkbook::setCellStyle(emscripten::val record) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setCellStyle");
  const std::string name = reader.string(record, "name", "cellStyle.name");
  fm_cell_style_record_t cs{};
  cs.name = name.c_str();
  cs.xf_id = reader.u32(record, "xfId", 0U, "cellStyle.xfId");
  cs.builtin_id = reader.u32(record, "builtinId", FM_CELL_STYLE_BUILTIN_ID_NONE, "cellStyle.builtinId");
  cs.i_level = reader.u32(record, "iLevel", 0U, "cellStyle.iLevel");
  cs.hidden = reader.boolean(record, "hidden", false, "cellStyle.hidden") ? 1 : 0;
  cs.custom_builtin = reader.boolean(record, "customBuiltin", false, "cellStyle.customBuiltin") ? 1 : 0;
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_workbook_set_cell_style(handle_, &cs));
}

void JsWorkbook::removeCellStyle(const std::string& name) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_cell_style(handle_, name.c_str()) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::listTableStyles() const {
  emscripten::val arr = emscripten::val::array();
  if (!succeeded(js_collect_list(handle_, 0U, &count_table_styles, &table_style_item, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

emscripten::val JsWorkbook::getTableStyle(const std::string& name) const {
  fm_table_style_t style{};
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_table_style(handle_, name.c_str(), &present, &style) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? table_style_to_val(style) : emscripten::val::null());
}

void JsWorkbook::setTableStyle(emscripten::val style) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setTableStyle");
  const std::string name = reader.string(style, "name", "style.name");
  const bool pivot = reader.boolean(style, "pivot", true, "style.pivot");
  const bool table = reader.boolean(style, "table", true, "style.table");
  const emscripten::val elements = reader.value(style, "elements", "style.elements");
  std::vector<std::string> types;
  std::vector<fm_table_style_element_t> records;
  if (reader.ok() && js_value_present(elements)) {
    if (!reader.is_array(elements, "style.elements")) {
      record_type_error("setTableStyle: `elements` must be an array");
      return;
    }
    const uint32_t count = reader.length(elements, "style.elements.length");
    types.reserve(count);
    records.reserve(count);
    for (uint32_t i = 0; i < count && reader.ok(); ++i) {
      const emscripten::val element = reader.array_element(elements, i, "style.elements[]");
      types.push_back(reader.string(element, "type", "style.elements[].type"));
      fm_table_style_element_t record{};
      record.size = reader.u32(element, "size", 1U, "style.elements[].size");
      const emscripten::val dxf = reader.value(element, "dxfId", "style.elements[].dxfId");
      if (reader.ok() && js_value_present(dxf)) {
        record.has_dxf_id = 1;
        record.dxf_id = reader.u32_value(dxf, 0U, "style.elements[].dxfId");
      }
      records.push_back(record);
    }
  }
  if (!reader.ok()) {
    return;
  }
  // `types` has stopped growing, so its `c_str()` views stay put.
  for (std::size_t i = 0; i < records.size(); ++i) {
    records[i].type = types[i].c_str();
  }
  fm_table_style_t record{};
  record.name = name.c_str();
  record.pivot = pivot ? 1 : 0;
  record.table = table ? 1 : 0;
  record.elements = records.empty() ? nullptr : records.data();
  record.element_count = static_cast<uint32_t>(records.size());
  succeeded(fm_workbook_set_table_style(handle_, &record));
}

void JsWorkbook::removeTableStyle(const std::string& name) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_table_style(handle_, name.c_str()) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getTableStyleDefaults() const {
  const char* table = nullptr;
  const char* pivot = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_table_style_defaults(handle_, &table, &pivot)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "tableStyle", table);
  js_set_cstr(o, "pivotStyle", pivot);
  return o;
}

void JsWorkbook::setTableStyleDefaults(emscripten::val defaults) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setTableStyleDefaults");
  std::string table;
  std::string pivot;
  const char* table_ptr = reader.optional_string(defaults, "tableStyle", table, "defaults.tableStyle");
  const char* pivot_ptr = reader.optional_string(defaults, "pivotStyle", pivot, "defaults.pivotStyle");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_workbook_set_table_style_defaults(handle_, table_ptr, pivot_ptr));
}
}  // namespace parts
}  // namespace wasm
}  // namespace formulon
