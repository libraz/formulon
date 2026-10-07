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
// entry point, copies the result into a `ListResult` of {sheet, row, col}
// objects, and frees the C-owned handle.
using TraceFn = fm_status_t (*)(const fm_workbook_t*, uint32_t, uint32_t, uint32_t, uint32_t, fm_cell_nodes_t**);

Napi::Array TraceToArray(const Napi::CallbackInfo& info, const fm_workbook_t* handle, TraceFn fn) {
  Napi::Env env = info.Env();
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  const uint32_t row = Workbook::ArgU32(info, 1);
  const uint32_t col = Workbook::ArgU32(info, 2);
  const uint32_t depth = Workbook::ArgU32(info, 3);
  Napi::Array arr = Napi::Array::New(env);
  if (handle == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  fm_cell_nodes_t* nodes = nullptr;
  fm_status_t rc = fn(handle, sheet, row, col, depth, &nodes);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  const std::size_t count = fm_cell_nodes_count(nodes);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_cell_node_t n{};
    rc = fm_cell_nodes_at(nodes, i, &n);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("sheet", Napi::Number::New(env, n.sheet));
    item.Set("row", Napi::Number::New(env, n.row));
    item.Set("col", Napi::Number::New(env, n.col));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  fm_cell_nodes_destroy(nodes);
  return FinishListResult(env, arr, rc);
}

// Shared bridge for `localizeFunctionName` / `canonicalizeFunctionName`.
using FunctionNameMapFn = fm_status_t (*)(const char*, int32_t, const char**);

Napi::Value MapFunctionName(const Napi::CallbackInfo& info, FunctionNameMapFn fn) {
  const std::string name = Workbook::ArgString(info, 0);
  const std::int32_t locale = Workbook::ArgI32(info, 1);
  const char* out = nullptr;
  const fm_status_t rc = fn(name.c_str(), locale, &out);
  return MakeStringResult(info.Env(), rc, out);
}

}  // namespace

// ---- Trace precedents / dependents ----------------------------------

Napi::Value Workbook::Precedents(const Napi::CallbackInfo& info) {
  return TraceToArray(info, handle_, fm_workbook_precedents);
}

Napi::Value Workbook::Dependents(const Napi::CallbackInfo& info) {
  return TraceToArray(info, handle_, fm_workbook_dependents);
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
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_spill_info_t spill{};
  const fm_status_t rc = fm_workbook_spill_info(handle_, sheet, row, col, &spill);
  out.Set("status", MakeStatus(env, rc));
  if (rc != 0) {
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
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  uint32_t count = 0;
  fm_status_t rc = fm_workbook_external_link_count(handle_, &count);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_external_link_record_t rec{};
    rc = fm_workbook_external_link_at(handle_, i, &rec);
    if (rc != 0) {
      return FinishListResult(env, arr, rc);
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
  return FinishListResult(env, arr, 0);
}

// ---- Function catalog -----------------------------------------------

Napi::Value Workbook::FunctionMetadata(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  const std::string name = ArgString(info, 0);
  const std::int32_t locale = ArgI32(info, 1);
  fm_function_metadata_t md{};
  fm_status_t rc = fm_function_metadata(name.c_str(), locale, &md);
  if (rc != 0) {
    out.Set("ok", Napi::Boolean::New(env, false));
    return out;
  }
  out.Set("ok", Napi::Boolean::New(env, true));
  out.Set("name", Napi::String::New(env, md.canonical_name != nullptr ? md.canonical_name : ""));
  out.Set("minArity", Napi::Number::New(env, md.min_arity));
  // `0xFFFFFFFF` is the unbounded / unknown-arity sentinel; surface it as
  // `null` so JS callers do not mistake it for a concrete upper bound.
  if (md.max_arity == 0xFFFFFFFFu) {
    out.Set("maxArity", env.Null());
  } else {
    out.Set("maxArity", Napi::Number::New(env, md.max_arity));
  }
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
  fm_status_t rc = 0;
  for (std::size_t i = 0; i < n; ++i) {
    const char* name = nullptr;
    rc = fm_function_name_at(i, &name);
    if (rc != 0) {
      break;
    }
    arr.Set(static_cast<uint32_t>(i), Napi::String::New(env, name != nullptr ? name : ""));
  }
  return FinishListResult(env, arr, rc);
}

Napi::Value Workbook::LocalizeFunctionName(const Napi::CallbackInfo& info) {
  return MapFunctionName(info, &fm_function_localize);
}

Napi::Value Workbook::CanonicalizeFunctionName(const Napi::CallbackInfo& info) {
  return MapFunctionName(info, &fm_function_canonicalize);
}

}  // namespace formulon_node
