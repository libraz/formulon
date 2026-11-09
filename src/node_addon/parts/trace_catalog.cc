// Dependency-graph trace (precedents / dependents), dynamic-array
// spill info, legacy array formulas, external-link enumeration, and the function-catalog
// metadata surface (metadata / names / localize / canonicalize). These
// are grouped here because they are non-mutating projections that share
// no state with the styles / sheet TUs.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/free_funcs.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

// Shared bridge for `precedents` / `dependents`: invokes the C ABI
// entry point, copies the result into an array of {sheet, row, col}
// objects, and frees the C-owned handle.
using TraceFn = fm_status_t (*)(const fm_workbook_t*, std::size_t, uint32_t, uint32_t, uint32_t, fm_cell_nodes_t**);

Napi::Value TraceToArray(const Napi::CallbackInfo& info, const fm_workbook_t* handle, TraceFn fn) {
  Napi::Env env = info.Env();
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  const uint32_t row = Workbook::ArgU32(info, 1);
  const uint32_t col = Workbook::ArgU32(info, 2);
  const uint32_t depth = Workbook::ArgU32(info, 3);
  Napi::Array arr = Napi::Array::New(env);
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_cell_nodes_t* nodes = nullptr;
  fm_status_t rc = fn(handle, sheet, row, col, depth, &nodes);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  std::size_t count = 0;
  rc = fm_cell_nodes_get_node_count(nodes, &count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_cell_node_info_t n = FM_STRUCT_INIT(fm_cell_node_info_t);
    rc = fm_cell_nodes_get_node_at(nodes, i, &n);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("sheet", Napi::Number::New(env, n.sheet_index));
    item.Set("row", Napi::Number::New(env, n.row));
    item.Set("col", Napi::Number::New(env, n.col));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  fm_cell_nodes_destroy(nodes);
  return ValueOrThrow(env, rc, arr);
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
  return ValueOrThrow(info.Env(), rc, JsString(info.Env(), out));
}

}  // namespace

// ---- Trace precedents / dependents ----------------------------------

Napi::Value Workbook::Precedents(const Napi::CallbackInfo& info) {
  return TraceToArray(info, handle_, fm_sheet_get_precedents);
}

Napi::Value Workbook::Dependents(const Napi::CallbackInfo& info) {
  return TraceToArray(info, handle_, fm_sheet_get_dependents);
}

// ---- Dynamic-array spill info ---------------------------------------

Napi::Value Workbook::GetSpill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_spill(handle_, sheet, row, col, &present, &spill);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, 0, env.Null());
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("rect", RangeToJs(env, spill.rect));
  out.Set("blocked", Napi::Boolean::New(env, spill.blocked != 0));
  return out;
}

// ---- Legacy array (CSE) formulas ------------------------------------

Napi::Value Workbook::GetArrayFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_array_formula_info_t array = FM_STRUCT_INIT(fm_array_formula_info_t);
  int32_t present = 0;
  const fm_status_t rc =
      fm_sheet_get_array_formula(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &present, &array);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, 0, env.Null());
  }
  Napi::Object value = Napi::Object::New(env);
  value.Set("rect", RangeToJs(env, array.rect));
  value.Set("formula", Napi::String::New(env, array.formula));
  return value;
}

Napi::Value Workbook::SetArrayFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRect(reader, ArgObject(info, 1), &rect)) {
    return env.Undefined();
  }
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string formula = ArgString(info, 2);
    return fm_sheet_set_array_formula(handle, ArgU32(info, 0), &rect, formula.c_str());
  });
}

// ---- External links -------------------------------------------------

Napi::Value Workbook::GetExternalLinkCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_get_external_link_count(handle_, &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

Napi::Value Workbook::GetExternalLink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_external_link_info_t rec = FM_STRUCT_INIT(fm_external_link_info_t);
  const fm_status_t rc = fm_workbook_get_external_link_at(handle_, ArgU32(info, 0), &rec);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object item = Napi::Object::New(env);
  item.Set("bookNumber", Napi::Number::New(env, rec.book_number));
  item.Set("relId", JsString(env, rec.rel_id));
  item.Set("partPath", JsString(env, rec.part_path));
  item.Set("target", JsStringOrNull(env, rec.target));
  item.Set("targetExternal", Napi::Boolean::New(env, rec.target_external != 0));
  item.Set("kind", Napi::Number::New(env, rec.kind));
  return item;
}

// ---- Function catalog -----------------------------------------------

Napi::Value FunctionMetadata(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::string name = Workbook::ArgString(info, 0);
  int32_t present = 0;
  fm_function_metadata_info_t md = FM_STRUCT_INIT(fm_function_metadata_info_t);
  fm_status_t rc = fm_get_function_metadata(name.c_str(), &present, &md);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("name", JsString(env, md.canonical_name));
  out.Set("minArity", Napi::Number::New(env, md.min_arity));
  out.Set("maxArity", JsNumberOrNull(env, md.has_max_arity, md.max_arity));
  out.Set("availability", Napi::Number::New(env, static_cast<uint32_t>(md.availability)));
  out.Set("signature", JsStringOrNull(env, md.signature_template));
  out.Set("description", JsStringOrNull(env, md.description));
  return OptionalOrNull(env, rc, present, out);
}

Napi::Value FunctionNames(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  const fm_status_t rc = CollectList<const char*>(
      arr, [](size_t* n) { return fm_get_function_count(n); },
      [](size_t i, const char** name) { return fm_get_function_name_at(i, name); },
      [&](const char* name) { return JsString(env, name); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value ListProfileIds(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  const fm_status_t rc = CollectList<const char*>(
      arr, [](size_t* n) { return fm_get_profile_id_count(n); },
      [](size_t i, const char** id) { return fm_get_profile_id_at(i, id); },
      [&](const char* id) { return JsString(env, id); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value LocalizeFunctionName(const Napi::CallbackInfo& info) {
  return MapProfileText(info, &fm_localize_function_name);
}

Napi::Value CanonicalizeFunctionName(const Napi::CallbackInfo& info) {
  return MapProfileText(info, &fm_canonicalize_function_name);
}

Napi::Value LocalizeFormula(const Napi::CallbackInfo& info) {
  return MapProfileText(info, &fm_localize_formula);
}

Napi::Value CanonicalizeFormula(const Napi::CallbackInfo& info) {
  return MapProfileText(info, &fm_canonicalize_formula);
}

namespace {

// Appends profile `profile_id`'s localized error names to `list`, in error-code order.
fm_status_t CollectLocaleErrorNames(Napi::Env env, const std::string& profile_id, Napi::Array& list) {
  std::size_t n = 0;
  fm_status_t rc = fm_get_locale_error_name_count(&n);
  for (std::size_t i = 0; i < n && rc == 0; ++i) {
    const char* canonical = nullptr;
    const char* localized = nullptr;
    std::int32_t measured = 0;
    rc = fm_get_locale_error_name_at(profile_id.c_str(), i, &canonical, &localized, &measured);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("canonical", JsString(env, canonical));
    item.Set("localized", JsString(env, localized));
    item.Set("measured", Napi::Boolean::New(env, measured != 0));
    list.Set(static_cast<uint32_t>(i), item);
  }
  return rc;
}

}  // namespace

Napi::Value LocaleFacts(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::string profile_id = Workbook::ArgString(info, 0);
  fm_locale_facts_info_t f = FM_STRUCT_INIT(fm_locale_facts_info_t);
  fm_status_t rc = fm_get_locale_facts(profile_id.c_str(), &f);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  const auto str = [&env](const char* s) { return JsString(env, s); };
  Napi::Object facts = Napi::Object::New(env);
  facts.Set("decimalSeparator", str(f.decimal_separator));
  facts.Set("groupSeparator", str(f.group_separator));
  facts.Set("listSeparator", str(f.list_separator));
  facts.Set("arrayColumnSeparator", str(f.array_column_separator));
  facts.Set("arrayRowSeparator", str(f.array_row_separator));
  facts.Set("trueName", str(f.true_name));
  facts.Set("falseName", str(f.false_name));
  facts.Set("dateOrder", Napi::Number::New(env, static_cast<int32_t>(f.date_order)));
  facts.Set("currencySymbol", str(f.currency_symbol));
  facts.Set("currencySuffix", Napi::Boolean::New(env, f.currency_suffix != 0));
  facts.Set("currencySpace", Napi::Boolean::New(env, f.currency_space != 0));
  facts.Set("currencyDefaultDecimals", Napi::Number::New(env, f.currency_default_decimals));
  facts.Set("measured", Napi::Boolean::New(env, f.measured != 0));
  facts.Set("collationMeasured", Napi::Boolean::New(env, f.collation_measured != 0));
  Napi::Array errors = Napi::Array::New(env);
  rc = CollectLocaleErrorNames(env, profile_id, errors);
  facts.Set("errorNames", errors);
  return ValueOrThrow(env, rc, facts);
}

Napi::Value CompareText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::string a = Workbook::ArgString(info, 0);
  const std::string b = Workbook::ArgString(info, 1);
  const auto strength = static_cast<fm_collation_strength_t>(Workbook::ArgU32(info, 2));
  const std::string profile_id = Workbook::ArgString(info, 3);
  int32_t order = 0;
  const fm_status_t rc = fm_compare_text(a.c_str(), b.c_str(), strength, profile_id.c_str(), &order);
  return ValueOrThrow(env, rc, Napi::Number::New(env, order));
}

Napi::Value CollationClasses(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array array;
  if (!ArrayArg(info, 0, &array)) {
    Napi::TypeError::New(env, "getCollationClasses expects (values:Value[], strength, profileId)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const uint32_t n = array.Length();
  const auto strength = static_cast<fm_collation_strength_t>(Workbook::ArgU32(info, 1));
  const std::string profile_id = Workbook::ArgString(info, 2);

  // Text payloads must keep their addresses until the C call returns.
  std::vector<std::string> texts(n);
  std::vector<fm_value_t> values(n);
  CheckedSpecReader reader(env);
  for (uint32_t i = 0; i < n; ++i) {
    Napi::Value element;
    if (!reader.ArrayElement(array, i, &element)) {
      return env.Undefined();
    }
    if (!JsToValue(env, element, &values[i], &texts[i])) {
      return env.Undefined();
    }
  }

  Napi::Uint32Array ids = Napi::Uint32Array::New(env, n);
  uint32_t class_count = 0;
  const fm_status_t rc = fm_get_collation_classes(values.data(), n, strength, profile_id.c_str(),
                                                  n == 0 ? nullptr : ids.Data(), &class_count);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("classIds", ids);
  out.Set("classCount", Napi::Number::New(env, class_count));
  return out;
}

}  // namespace formulon_node
