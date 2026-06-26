// Copyright 2026 libraz. Licensed under the MIT License.
//
// Dependency-graph trace (precedents / dependents), dynamic-array
// spill info, external-link enumeration, and the function-catalog
// metadata surface (metadata / names / localize / canonicalize). These
// are grouped here because they are non-mutating projections that share
// no state with the styles / sheet TUs.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

// Shared bridge for `precedents` / `dependents`: invokes the C ABI
// entry point, copies the result into a JS array of {sheet, row, col}
// objects, and frees the C-owned handle. Returns an empty array on any
// error so the JS side needs no separate failure path.
using TraceFn = fm_status_t (*)(const fm_workbook_t*, uint32_t, uint32_t, uint32_t, uint32_t, fm_cell_nodes_t**);

Napi::Array TraceToArray(Napi::Env env, const fm_workbook_t* handle, TraceFn fn, uint32_t sheet, uint32_t row,
                         uint32_t col, uint32_t depth) {
  Napi::Array arr = Napi::Array::New(env);
  if (handle == nullptr) {
    return arr;
  }
  fm_cell_nodes_t* nodes = nullptr;
  if (fn(handle, sheet, row, col, depth, &nodes) != 0) {
    return arr;
  }
  const std::size_t count = fm_cell_nodes_count(nodes);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_cell_node_t n{};
    if (fm_cell_nodes_at(nodes, i, &n) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("sheet", Napi::Number::New(env, n.sheet));
    item.Set("row", Napi::Number::New(env, n.row));
    item.Set("col", Napi::Number::New(env, n.col));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  fm_cell_nodes_destroy(nodes);
  return arr;
}

}  // namespace

// ---- Trace precedents / dependents ----------------------------------

Napi::Value Workbook::Precedents(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const uint32_t depth = ArgU32(info, 3);
  return TraceToArray(env, handle_, fm_workbook_precedents, sheet, row, col, depth);
}

Napi::Value Workbook::Dependents(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const uint32_t depth = ArgU32(info, 3);
  return TraceToArray(env, handle_, fm_workbook_dependents, sheet, row, col, depth);
}

// ---- Dynamic-array spill info ---------------------------------------

Napi::Value Workbook::SpillInfo(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("engaged", Napi::Boolean::New(env, false));
  out.Set("anchorRow", Napi::Number::New(env, 0));
  out.Set("anchorCol", Napi::Number::New(env, 0));
  out.Set("rows", Napi::Number::New(env, 0));
  out.Set("cols", Napi::Number::New(env, 0));
  if (handle_ == nullptr) {
    return out;
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_spill_info_t spill{};
  if (fm_workbook_spill_info(handle_, sheet, row, col, &spill) != 0) {
    return out;
  }
  out.Set("engaged", Napi::Boolean::New(env, spill.engaged != 0));
  out.Set("anchorRow", Napi::Number::New(env, spill.anchor_row));
  out.Set("anchorCol", Napi::Number::New(env, spill.anchor_col));
  out.Set("rows", Napi::Number::New(env, spill.rows));
  out.Set("cols", Napi::Number::New(env, spill.cols));
  return out;
}

// ---- External links -------------------------------------------------

Napi::Value Workbook::GetExternalLinks(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return arr;
  }
  uint32_t count = 0;
  if (fm_workbook_external_link_count(handle_, &count) != 0) {
    return arr;
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_external_link_record_t rec{};
    if (fm_workbook_external_link_at(handle_, i, &rec) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("index", Napi::Number::New(env, rec.index));
    item.Set("relId", Napi::String::New(env, rec.rel_id != nullptr ? rec.rel_id : ""));
    item.Set("partPath", Napi::String::New(env, rec.part_path != nullptr ? rec.part_path : ""));
    item.Set("target", Napi::String::New(env, rec.target != nullptr ? rec.target : ""));
    item.Set("targetExternal", Napi::Boolean::New(env, rec.target_external != 0));
    item.Set("kind", Napi::Number::New(env, rec.kind));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return arr;
}

// ---- Function catalog -----------------------------------------------

Napi::Value Workbook::FunctionMetadata(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  const std::string name = ArgString(info, 0);
  const uint32_t locale = ArgU32(info, 1);
  fm_function_metadata_t md{};
  fm_status_t rc = fm_function_metadata(name.c_str(), static_cast<fm_locale_t>(locale), &md);
  if (rc != 0) {
    out.Set("ok", Napi::Boolean::New(env, false));
    return out;
  }
  out.Set("ok", Napi::Boolean::New(env, true));
  out.Set("name", Napi::String::New(env, md.canonical_name != nullptr ? md.canonical_name : ""));
  out.Set("minArity", Napi::Number::New(env, md.min_arity));
  out.Set("maxArity", Napi::Number::New(env, md.max_arity));
  out.Set("availability", Napi::Number::New(env, static_cast<uint32_t>(md.availability)));
  if (md.signature_template != nullptr) {
    out.Set("signatureTemplate", Napi::String::New(env, md.signature_template));
  }
  if (md.description != nullptr) {
    out.Set("description", Napi::String::New(env, md.description));
  }
  return out;
}

Napi::Value Workbook::FunctionNames(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  const std::size_t n = fm_function_count();
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < n; ++i) {
    const char* name = nullptr;
    if (fm_function_name_at(i, &name) != 0 || name == nullptr) {
      continue;
    }
    arr.Set(static_cast<uint32_t>(emitted), Napi::String::New(env, name));
    ++emitted;
  }
  return arr;
}

Napi::Value Workbook::LocalizeFunctionName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::string canonical = ArgString(info, 0);
  const uint32_t locale = ArgU32(info, 1);
  const char* out = nullptr;
  if (fm_function_localize(canonical.c_str(), static_cast<fm_locale_t>(locale), &out) != 0 || out == nullptr) {
    return Napi::String::New(env, "");
  }
  return Napi::String::New(env, out);
}

Napi::Value Workbook::CanonicalizeFunctionName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::string localized = ArgString(info, 0);
  const uint32_t locale = ArgU32(info, 1);
  const char* out = nullptr;
  if (fm_function_canonicalize(localized.c_str(), static_cast<fm_locale_t>(locale), &out) != 0 || out == nullptr) {
    return Napi::String::New(env, "");
  }
  return Napi::String::New(env, out);
}

}  // namespace formulon_node
