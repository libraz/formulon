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

// Shared bridge for the `(text, profileId)` -> text catalog calls
// (`localizeFunctionName`, `canonicalizeFunctionName`, `localizeFormula`,
// `canonicalizeFormula`).
using ProfileTextMapFn = fm_status_t (*)(const char*, const char*, const char**);

Napi::Value MapProfileText(const Napi::CallbackInfo& info, ProfileTextMapFn fn) {
  const std::string text = Workbook::ArgString(info, 0);
  const std::string profile_id = Workbook::ArgString(info, 1);
  const char* out = nullptr;
  const fm_status_t rc = fn(text.c_str(), profile_id.c_str(), &out);
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
  fm_function_metadata_t md{};
  fm_status_t rc = fm_function_metadata(name.c_str(), &md);
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
  return MapProfileText(info, &fm_function_localize);
}

Napi::Value Workbook::CanonicalizeFunctionName(const Napi::CallbackInfo& info) {
  return MapProfileText(info, &fm_function_canonicalize);
}

Napi::Value Workbook::LocalizeFormula(const Napi::CallbackInfo& info) {
  return MapProfileText(info, &fm_formula_localize);
}

Napi::Value Workbook::CanonicalizeFormula(const Napi::CallbackInfo& info) {
  return MapProfileText(info, &fm_formula_canonicalize);
}

Napi::Value Workbook::LocaleFacts(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("facts", env.Null());
  const std::string profile_id = ArgString(info, 0);
  fm_locale_facts_t f{};
  fm_status_t rc = fm_locale_facts(profile_id.c_str(), &f);
  if (rc != 0) {
    out.Set("status", MakeStatus(env, rc));
    return out;
  }
  static const char* const kDateOrders[] = {"mdy", "ymd", "dmy"};
  const auto str = [&env](const char* s) { return Napi::String::New(env, s != nullptr ? s : ""); };
  Napi::Object facts = Napi::Object::New(env);
  facts.Set("decimalSeparator", str(f.decimal_separator));
  facts.Set("groupSeparator", str(f.group_separator));
  facts.Set("listSeparator", str(f.list_separator));
  facts.Set("arrayColumnSeparator", str(f.array_column_separator));
  facts.Set("arrayRowSeparator", str(f.array_row_separator));
  facts.Set("trueName", str(f.true_name));
  facts.Set("falseName", str(f.false_name));
  facts.Set("dateOrder", Napi::String::New(env, kDateOrders[static_cast<std::size_t>(f.date_order)]));
  facts.Set("currencySymbol", str(f.currency_symbol));
  facts.Set("currencySuffix", Napi::Boolean::New(env, f.currency_suffix != 0));
  facts.Set("currencySpace", Napi::Boolean::New(env, f.currency_space != 0));
  facts.Set("currencyDefaultDecimals", Napi::Number::New(env, f.currency_default_decimals));
  facts.Set("measured", Napi::Boolean::New(env, f.measured != 0));
  Napi::Array errors = Napi::Array::New(env);
  const std::size_t n = fm_locale_error_name_count();
  for (std::size_t i = 0; i < n && rc == 0; ++i) {
    const char* canonical = nullptr;
    const char* localized = nullptr;
    std::int32_t measured = 0;
    rc = fm_locale_error_name(profile_id.c_str(), i, &canonical, &localized, &measured);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("canonical", str(canonical));
    item.Set("localized", str(localized));
    item.Set("measured", Napi::Boolean::New(env, measured != 0));
    errors.Set(static_cast<uint32_t>(i), item);
  }
  facts.Set("errorNames", errors);
  out.Set("status", MakeStatus(env, rc));
  if (rc == 0) {
    out.Set("facts", facts);
  }
  return out;
}

}  // namespace formulon_node
