// Named cell-style and custom table / PivotTable style bindings.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/styles_internal.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

using detail::CellXfResultToJs;
using detail::StyleCountResult;

// ---- Named cell styles ----------------------------------------------

namespace {

// The JS `CellStyleRecord` record of `cs`.
Napi::Object CellStyleToJs(Napi::Env env, const fm_cell_style_record_t& cs) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("name", JsString(env, cs.name));
  out.Set("xfId", Napi::Number::New(env, cs.xf_id));
  out.Set("builtinId", Napi::Number::New(env, cs.builtin_id));
  out.Set("iLevel", Napi::Number::New(env, cs.i_level));
  out.Set("hidden", Napi::Boolean::New(env, cs.hidden != 0));
  out.Set("customBuiltin", Napi::Boolean::New(env, cs.custom_builtin != 0));
  return out;
}

// The JS `TableStyle` record of `style`.
Napi::Object TableStyleToJs(Napi::Env env, const fm_table_style_t& style) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("name", JsString(env, style.name));
  out.Set("pivot", JsBool(env, style.pivot));
  out.Set("table", JsBool(env, style.table));
  Napi::Array elements = Napi::Array::New(env, style.element_count);
  for (uint32_t i = 0; i < style.element_count; ++i) {
    const fm_table_style_element_t& element = style.elements[i];
    Napi::Object e = Napi::Object::New(env);
    e.Set("type", JsString(env, element.type));
    e.Set("size", Napi::Number::New(env, element.size));
    if (element.has_dxf_id != 0) {
      e.Set("dxfId", Napi::Number::New(env, element.dxf_id));
    }
    elements.Set(i, e);
  }
  out.Set("elements", elements);
  return out;
}

// Lists a workbook-scoped style table through its count and `_at` getters.
template <typename Record, typename Convert>
Napi::Value StyleList(Napi::Env env, const fm_workbook_t* handle, fm_status_t (*count)(const fm_workbook_t*, size_t*),
                      fm_status_t (*at)(const fm_workbook_t*, size_t, Record*), Convert convert) {
  Napi::Array arr = Napi::Array::New(env);
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  size_t n = 0;
  fm_status_t rc = count(handle, &n);
  for (size_t i = 0; rc == 0 && i < n; ++i) {
    Record record{};
    rc = at(handle, i, &record);
    if (rc == 0) {
      arr.Set(static_cast<uint32_t>(i), convert(env, record));
    }
  }
  return ValueOrThrow(env, rc, arr);
}

}  // namespace

Napi::Value Workbook::ListCellStyles(const Napi::CallbackInfo& info) {
  return StyleList(info.Env(), handle_, &fm_workbook_get_cell_style_count, &fm_workbook_get_cell_style_at,
                   &CellStyleToJs);
}

Napi::Value Workbook::CellStyleXfCount(const Napi::CallbackInfo& info) {
  return StyleCountResult(info, handle_, &fm_workbook_get_cell_style_xf_count);
}

Napi::Value Workbook::GetCellStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string name = ArgString(info, 0);
  fm_cell_style_record_t cs{};
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_cell_style(handle_, name.c_str(), &present, &cs);
  return OptionalOrNull(env, rc, present, rc == 0 && present != 0 ? Napi::Value(CellStyleToJs(env, cs)) : env.Null());
}

Napi::Value Workbook::GetCellStyleXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t index = ArgU32(info, 0);
  fm_cell_xf_t xf{};
  const fm_status_t rc = fm_workbook_get_cell_style_xf(handle_, index, &xf);
  return rc == 0 ? Napi::Value(CellXfResultToJs(env, xf)) : ThrowStatus(env, rc);
}

Napi::Value Workbook::SetCellStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object record = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  std::string name;
  reader.String(record, "name", &name);
  fm_cell_style_record_t cs{};
  cs.name = name.c_str();
  cs.xf_id = reader.U32(record, "xfId", 0U);
  cs.builtin_id = reader.U32(record, "builtinId", FM_CELL_STYLE_BUILTIN_ID_NONE);
  cs.i_level = reader.U32(record, "iLevel", 0U);
  cs.hidden = reader.Bool(record, "hidden", false) ? 1 : 0;
  cs.custom_builtin = reader.Bool(record, "customBuiltin", false) ? 1 : 0;
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_workbook_set_cell_style(handle_, &cs), env.Undefined());
}

Napi::Value Workbook::RemoveCellStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string name = ArgString(info, 0);
  return ValueOrThrow(env, fm_workbook_remove_cell_style(handle_, name.c_str()), env.Undefined());
}

// ---- Custom table / PivotTable styles --------------------------------

Napi::Value Workbook::ListTableStyles(const Napi::CallbackInfo& info) {
  return StyleList(info.Env(), handle_, &fm_workbook_get_table_style_count, &fm_workbook_get_table_style_at,
                   &TableStyleToJs);
}

Napi::Value Workbook::GetTableStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string name = ArgString(info, 0);
  fm_table_style_t style{};
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_table_style(handle_, name.c_str(), &present, &style);
  return OptionalOrNull(env, rc, present,
                        rc == 0 && present != 0 ? Napi::Value(TableStyleToJs(env, style)) : env.Null());
}

Napi::Value Workbook::SetTableStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const Napi::Object record = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  std::string name;
  reader.String(record, "name", &name);
  const bool pivot = reader.Bool(record, "pivot", true);
  const bool table = reader.Bool(record, "table", true);
  Napi::Array elements;
  const bool has_elements = reader.Array(record, "elements", &elements);
  std::vector<std::string> types;
  std::vector<fm_table_style_element_t> records;
  for (uint32_t i = 0; has_elements && reader.ok() && i < elements.Length(); ++i) {
    Napi::Value value;
    Napi::Object element;
    if (!reader.ArrayElement(elements, i, &value) || !reader.Object(value, "elements[]", &element)) {
      continue;
    }
    std::string type;
    reader.String(element, "type", &type);
    types.push_back(std::move(type));
    fm_table_style_element_t out{};
    out.size = reader.U32(element, "size", 1U);
    bool has_dxf = false;
    out.dxf_id = reader.U32(element, "dxfId", 0U, &has_dxf);
    out.has_dxf_id = has_dxf ? 1 : 0;
    records.push_back(out);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  // `types` has stopped growing, so its `c_str()` views stay put.
  for (std::size_t i = 0; i < records.size(); ++i) {
    records[i].type = types[i].c_str();
  }
  fm_table_style_t style{};
  style.name = name.c_str();
  style.pivot = pivot ? 1 : 0;
  style.table = table ? 1 : 0;
  style.elements = records.empty() ? nullptr : records.data();
  style.element_count = static_cast<uint32_t>(records.size());
  return ValueOrThrow(env, fm_workbook_set_table_style(handle_, &style), env.Undefined());
}

Napi::Value Workbook::RemoveTableStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string name = ArgString(info, 0);
  return ValueOrThrow(env, fm_workbook_remove_table_style(handle_, name.c_str()), env.Undefined());
}

Napi::Value Workbook::GetTableStyleDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const char* table = nullptr;
  const char* pivot = nullptr;
  const fm_status_t rc = fm_workbook_get_table_style_defaults(handle_, &table, &pivot);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("tableStyle", JsStringOrNull(env, table));
  out.Set("pivotStyle", JsStringOrNull(env, pivot));
  return out;
}

Napi::Value Workbook::SetTableStyleDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const Napi::Object defaults = ArgObject(info, 0);
  // An omitted name keeps the default (NULL); `null` removes it ("").
  CheckedSpecReader reader(env);
  std::string table;
  std::string pivot;
  const auto name = [&](const char* key, std::string* store) -> const char* {
    const Napi::Value v = reader.Get(defaults, key);
    switch (ClassifyPatchKey(reader, v, key, true)) {
      case PatchKey::kClear:
        return "";
      case PatchKey::kSet:
        return reader.String(v, key, store) ? store->c_str() : nullptr;
      case PatchKey::kKeep:
        break;
    }
    return nullptr;
  };
  const char* table_ptr = name("tableStyle", &table);
  const char* pivot_ptr = name("pivotStyle", &pivot);
  if (!reader.ok()) {
    return env.Undefined();
  }
  const fm_status_t rc = fm_workbook_set_table_style_defaults(handle_, table_ptr, pivot_ptr);
  return ValueOrThrow(env, rc, env.Undefined());
}

}  // namespace formulon_node
