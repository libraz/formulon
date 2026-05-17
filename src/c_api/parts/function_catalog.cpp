// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// C ABI - function catalog metadata.
//
// Surfaces the registry's canonical-name + arity data through the C ABI
// so JS / Python autocomplete UIs can enumerate Formulon functions
// without reaching into the engine internals. `description` /
// `signature_template` are reserved for the locale metadata table
// (data/function_metadata_<locale>.cpp) and stay `NULL` until that
// table is wired up; the surface itself is shippable today.
//
// The sorted name list is computed on first use and cached for the
// process lifetime - order matters for UIs that expect deterministic
// enumeration, and the registry's `for_each_name` does not promise
// any order.

#include <algorithm>
#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/function_registry.h"
#include "utils/error.h"

using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

namespace {

struct FunctionAvailabilityEntry {
  std::string_view name;
  fm_function_availability_t availability;
};

fm_function_availability_t function_availability(std::string_view canonical_name) {
  static constexpr FunctionAvailabilityEntry kEntries[] = {
      {"CELL", FM_FUNCTION_ENVIRONMENT_BOUND},
      {"COPILOT", FM_FUNCTION_UNAVAILABLE_STUB},
      {"CUBEKPIMEMBER", FM_FUNCTION_UNAVAILABLE_STUB},
      {"CUBEMEMBER", FM_FUNCTION_UNAVAILABLE_STUB},
      {"CUBEMEMBERPROPERTY", FM_FUNCTION_UNAVAILABLE_STUB},
      {"CUBERANKEDMEMBER", FM_FUNCTION_UNAVAILABLE_STUB},
      {"CUBESET", FM_FUNCTION_UNAVAILABLE_STUB},
      {"CUBESETCOUNT", FM_FUNCTION_UNAVAILABLE_STUB},
      {"CUBEVALUE", FM_FUNCTION_UNAVAILABLE_STUB},
      {"DETECTLANGUAGE", FM_FUNCTION_UNAVAILABLE_STUB},
      {"IMAGE", FM_FUNCTION_UNAVAILABLE_STUB},
      {"INFO", FM_FUNCTION_ENVIRONMENT_BOUND},
      {"PY", FM_FUNCTION_UNAVAILABLE_STUB},
      {"RTD", FM_FUNCTION_UNAVAILABLE_STUB},
      {"STOCKHISTORY", FM_FUNCTION_UNAVAILABLE_STUB},
      {"TRANSLATE", FM_FUNCTION_UNAVAILABLE_STUB},
      {"WEBSERVICE", FM_FUNCTION_UNAVAILABLE_STUB},
  };
  for (const auto& entry : kEntries) {
    if (entry.name == canonical_name) {
      return entry.availability;
    }
  }
  return FM_FUNCTION_IMPLEMENTED;
}

const std::vector<std::string>& sorted_function_names() {
  static const std::vector<std::string> names = []() {
    std::vector<std::string> out;
    formulon::eval::default_registry().for_each_name(
        [](std::string_view name, void* ctx) { static_cast<std::vector<std::string>*>(ctx)->emplace_back(name); },
        &out);
    std::sort(out.begin(), out.end());
    return out;
  }();
  return names;
}

}  // namespace

extern "C" fm_status_t fm_function_metadata(const char* name, fm_locale_t /*locale*/, fm_function_metadata_t* out) {
  clear_last_error();
  if (name == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_function_metadata: NULL argument");
  }
  const auto& reg = formulon::eval::default_registry();
  const auto* def = reg.lookup(std::string_view(name));
  if (def == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_metadata: unknown function",
                             std::string("name=") + name);
  }
  *out = fm_function_metadata_t{};
  // `canonical_name` is held in static storage by the registry; surface
  // a borrowed pointer so callers can hold the result indefinitely.
  out->canonical_name = def->canonical_name.data();
  out->min_arity = def->min_arity;
  out->max_arity = def->max_arity;
  out->availability = function_availability(def->canonical_name);
  // Locale metadata table (description / signature_template) is not
  // yet populated - the public ABI documents these as nullable.
  out->signature_template = nullptr;
  out->description = nullptr;
  return 0;
}

extern "C" std::size_t fm_function_count(void) {
  return sorted_function_names().size();
}

extern "C" fm_status_t fm_function_name_at(std::size_t idx, const char** out_name) {
  clear_last_error();
  if (out_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_function_name_at: out_name is NULL");
  }
  const auto& names = sorted_function_names();
  if (idx >= names.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_name_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(names.size()));
  }
  *out_name = names[idx].c_str();
  return 0;
}

extern "C" fm_status_t fm_function_localize(const char* canonical_name, fm_locale_t /*locale*/,
                                            const char** out_localized) {
  clear_last_error();
  if (canonical_name == nullptr || out_localized == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_function_localize: NULL argument");
  }
  const auto* def = formulon::eval::default_registry().lookup(std::string_view(canonical_name));
  if (def == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_localize: unknown function",
                             std::string("canonical_name=") + canonical_name);
  }
  // Locale alias table not yet populated - fall through to the canonical
  // name. Once `data/function_names_<locale>.csv` lands, this lookup
  // will branch on `locale` and consult the alias table first.
  *out_localized = def->canonical_name.data();
  return 0;
}

extern "C" fm_status_t fm_function_canonicalize(const char* localized_name, fm_locale_t /*locale*/,
                                                const char** out_canonical) {
  clear_last_error();
  if (localized_name == nullptr || out_canonical == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_function_canonicalize: NULL argument");
  }
  // Alias table not yet populated - fall through to a case-insensitive
  // canonical-name match.
  const auto* def = formulon::eval::default_registry().lookup(std::string_view(localized_name));
  if (def == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_function_canonicalize: unknown function",
                             std::string("localized_name=") + localized_name);
  }
  *out_canonical = def->canonical_name.data();
  return 0;
}
