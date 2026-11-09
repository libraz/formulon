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

emscripten::val cell_style_to_val(const fm_cell_style_t& cs) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", cs.name);
  o.set("styleXfIndex", cs.style_xf_index);
  js_set_present_or_null(o, "builtinId", cs.has_builtin_id, cs.builtin_id);
  o.set("outlineLevel", cs.outline_level);
  o.set("hidden", cs.hidden != 0);
  o.set("customBuiltin", cs.custom_builtin != 0);
  return o;
}

emscripten::val table_style_to_val(const fm_table_style_t& style) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", style.name);
  o.set("pivot", style.pivot == FM_TRISTATE_TRUE);
  o.set("table", style.table == FM_TRISTATE_TRUE);
  emscripten::val elements = emscripten::val::array();
  for (uint32_t i = 0; i < style.element_count; ++i) {
    const fm_table_style_element_t& element = style.elements[i];
    emscripten::val e = emscripten::val::object();
    e.set("type", element.type);
    e.set("size", element.size);
    e.set("dxfIndex", element.has_dxf_index != 0 ? emscripten::val(element.dxf_index) : emscripten::val::null());
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
  fm_cell_style_t cs{};
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
  fm_table_style_t style = FM_STRUCT_INIT(fm_table_style_t);
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
  fm_cell_style_t cs{};
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
  JsNarrowNumericReader reader("Workbook.setCellStyle");
  const std::string name = reader.string(record, "name", "name");
  fm_cell_style_t cs{};
  cs.name = name.c_str();
  cs.style_xf_index = reader.u32(record, "styleXfIndex", 0U, "styleXfIndex");
  const emscripten::val builtin_id = reader.value(record, "builtinId", "builtinId");
  if (reader.ok() && js_value_present(builtin_id)) {
    cs.has_builtin_id = 1;
    cs.builtin_id = reader.u32_value(builtin_id, 0U, "builtinId");
  }
  cs.outline_level = reader.u32(record, "outlineLevel", 0U, "outlineLevel");
  cs.hidden = reader.boolean(record, "hidden", false, "hidden") ? 1 : 0;
  cs.custom_builtin = reader.boolean(record, "customBuiltin", false, "customBuiltin") ? 1 : 0;
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
  fm_table_style_t style = FM_STRUCT_INIT(fm_table_style_t);
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
  JsNarrowNumericReader reader("Workbook.setTableStyle");
  const std::string name = reader.string(style, "name", "style.name");
  // An omitted or `null` flag is the default (usable).
  const auto usable = [&](const char* key, const char* field) -> int32_t {
    const emscripten::val v = reader.value(style, key, field);
    if (!reader.ok() || !js_value_present(v)) {
      return FM_TRISTATE_DEFAULT;
    }
    return reader.boolean_value(v, true) ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  };
  const int32_t pivot = usable("pivot", "style.pivot");
  const int32_t table = usable("table", "style.table");
  const emscripten::val elements = reader.value(style, "elements", "style.elements");
  std::vector<fm_table_style_element_t> records;
  if (reader.ok() && js_value_present(elements)) {
    if (!reader.is_array(elements, "style.elements")) {
      record_type_error("Workbook.setTableStyle elements must be an array");
      return;
    }
    const uint32_t count = reader.length(elements, "style.elements.length");
    records.reserve(count);
    for (uint32_t i = 0; i < count && reader.ok(); ++i) {
      const emscripten::val element = reader.array_element(elements, i, "style.elements[]");
      fm_table_style_element_t record{};
      record.type = static_cast<int32_t>(reader.required_u32(element, "type", "style.elements[].type"));
      record.size = reader.u32(element, "size", 1U, "style.elements[].size");
      const emscripten::val dxf = reader.value(element, "dxfIndex", "style.elements[].dxfIndex");
      if (reader.ok() && js_value_present(dxf)) {
        record.has_dxf_index = 1;
        record.dxf_index = reader.u32_value(dxf, 0U, "style.elements[].dxfIndex");
      }
      records.push_back(record);
    }
  }
  if (!reader.ok()) {
    return;
  }
  fm_table_style_t record = FM_STRUCT_INIT(fm_table_style_t);
  record.name = name.c_str();
  record.pivot = pivot;
  record.table = table;
  record.elements = records.empty() ? nullptr : records.data();
  record.element_count = static_cast<uint32_t>(records.size());
  succeeded(fm_workbook_set_table_style(handle_, &record));
}

void JsWorkbook::removeTableStyle(const std::string& name) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_table_style(handle_, name.c_str()) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getTableStyleDefaults() const {
  fm_table_style_defaults_t defaults = FM_STRUCT_INIT(fm_table_style_defaults_t);
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_table_style_defaults(handle_, &defaults)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr_or_null_fields(
      o, {{"tableStyleName", defaults.table_style_name}, {"pivotStyleName", defaults.pivot_style_name}});
  return o;
}

void JsWorkbook::setTableStyleDefaults(emscripten::val defaults) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Workbook.updateTableStyleDefaults");
  std::string table;
  std::string pivot;
  fm_table_style_defaults_patch_t patch = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  js_patch_string(reader, defaults, "tableStyleName", "patch.tableStyleName", true, table, &patch.table_style_name_op,
                  &patch.table_style_name);
  js_patch_string(reader, defaults, "pivotStyleName", "patch.pivotStyleName", true, pivot, &patch.pivot_style_name_op,
                  &patch.pivot_style_name);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_workbook_update_table_style_defaults(handle_, &patch));
}
}  // namespace parts
}  // namespace wasm
}  // namespace formulon
