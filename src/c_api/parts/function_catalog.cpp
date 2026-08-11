//
// C ABI - function catalog metadata.
//
// Surfaces the engine's canonical-name + arity data through the C ABI
// so JS / Python autocomplete UIs can enumerate Formulon functions
// without reaching into the engine internals. `description` /
// `signature_template` are reserved for the locale metadata table
// (data/function_metadata_<locale>.cpp) and stay `NULL` until that
// table is wired up; the surface itself is shippable today.
//
// A runtime-recognised function name comes from one of three sources:
// the eager `FunctionRegistry`, the tree walker's lazy-dispatch table
// (`eval::lazy_form_names()` — IF / XLOOKUP / SUMIFS / VLOOKUP / ...),
// or the parser's special forms (`eval::parser_special_form_names()` —
// LET / LAMBDA). All three are enumerated and de-duplicated so the
// catalog matches what the evaluator actually accepts. Lazy and
// special forms carry no `FunctionDef`, so they have no arity data;
// they report `min_arity = 0` and the unbounded `max_arity` sentinel.
//
// The sorted name list is computed on first use and cached for the
// process lifetime - order matters for UIs that expect deterministic
// enumeration, and the underlying iteration hooks do not promise any
// order.

#include <algorithm>
#include <cctype>
#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/function_registry.h"
#include "eval/special_forms_catalog.h"
#include "eval/tree_walker.h"
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

// Case-insensitive ASCII equality. Canonical catalog names are already
// UPPERCASE, so this only has to fold the incoming query.
bool ascii_iequals(std::string_view a, std::string_view b) {
  if (a.size() != b.size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.size(); ++i) {
    if (std::toupper(static_cast<unsigned char>(a[i])) != std::toupper(static_cast<unsigned char>(b[i]))) {
      return false;
    }
  }
  return true;
}

// Returns the static canonical-name pointer for `name` if it belongs to the
// lazy-dispatch or parser special-form tables (matched case-insensitively),
// or nullptr otherwise. The returned pointer has program lifetime.
const char* lookup_lazy_or_special(std::string_view name) {
  for (const char* const* p = formulon::eval::lazy_form_names(); p != nullptr && *p != nullptr; ++p) {
    if (ascii_iequals(name, *p)) {
      return *p;
    }
  }
  for (const char* const* p = formulon::eval::parser_special_form_names(); p != nullptr && *p != nullptr; ++p) {
    if (ascii_iequals(name, *p)) {
      return *p;
    }
  }
  return nullptr;
}

// Resolves `name` to its canonical static-storage name pointer across the
// same three sources the enumeration draws from: the eager
// `FunctionRegistry` first, then the lazy-dispatch and parser
// special-form tables. Returns nullptr when the name is recognised by
// none. Used so canonicalize / localize accept every enumerated function
// (XLOOKUP / SUMIFS / LET / LAMBDA / ...), not just the eager subset.
const char* resolve_canonical_name(std::string_view name) {
  if (const auto* def = formulon::eval::default_registry().lookup(name); def != nullptr) {
    return def->canonical_name.data();
  }
  return lookup_lazy_or_special(name);
}

const std::vector<std::string>& sorted_function_names() {
  static const std::vector<std::string> names = []() {
    std::vector<std::string> out;
    formulon::eval::default_registry().for_each_name(
        [](std::string_view name, void* ctx) { static_cast<std::vector<std::string>*>(ctx)->emplace_back(name); },
        &out);
    for (const char* const* p = formulon::eval::lazy_form_names(); p != nullptr && *p != nullptr; ++p) {
      out.emplace_back(*p);
    }
    for (const char* const* p = formulon::eval::parser_special_form_names(); p != nullptr && *p != nullptr; ++p) {
      out.emplace_back(*p);
    }
    std::sort(out.begin(), out.end());
    // A name can appear in more than one source (e.g. registry + lazy); keep a
    // single entry so the enumeration count matches the recognised set.
    out.erase(std::unique(out.begin(), out.end()), out.end());
    return out;
  }();
  return names;
}

}  // namespace

extern "C" fm_status_t fm_function_metadata(const char* name, std::int32_t locale, fm_function_metadata_t* out) {
  clear_last_error();
  if (name == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_function_metadata: NULL argument");
  }
  *out = fm_function_metadata_t{};
  if (locale < FM_LOCALE_EN_US || locale > FM_LOCALE_JA_JP) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_metadata: invalid locale",
                             "locale=" + std::to_string(locale));
  }
  const auto& reg = formulon::eval::default_registry();
  const auto* def = reg.lookup(std::string_view(name));
  if (def == nullptr) {
    // Not in the eager registry: it may still be a lazy-dispatch form
    // (XLOOKUP / SUMIFS / VLOOKUP / ...) or a parser special form (LET /
    // LAMBDA). Those have no `FunctionDef`, so arity is unknown and is
    // reported as `min_arity = 0` with the unbounded `max_arity` sentinel.
    const char* canonical = lookup_lazy_or_special(std::string_view(name));
    if (canonical == nullptr) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_metadata: unknown function",
                               std::string("name=") + name);
    }
    *out = fm_function_metadata_t{};
    out->canonical_name = canonical;
    out->min_arity = 0;
    out->max_arity = formulon::eval::kVariadic;
    out->availability = function_availability(std::string_view(canonical));
    out->signature_template = nullptr;
    out->description = nullptr;
    return 0;
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

extern "C" fm_status_t fm_function_localize(const char* canonical_name, std::int32_t locale,
                                            const char** out_localized) {
  clear_last_error();
  if (canonical_name == nullptr || out_localized == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_function_localize: NULL argument");
  }
  *out_localized = nullptr;
  if (locale < FM_LOCALE_EN_US || locale > FM_LOCALE_JA_JP) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_localize: invalid locale",
                             "locale=" + std::to_string(locale));
  }
  const char* canonical = resolve_canonical_name(std::string_view(canonical_name));
  if (canonical == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_localize: unknown function",
                             std::string("canonical_name=") + canonical_name);
  }
  // Locale alias table not yet populated - fall through to the canonical
  // name. Once `data/function_names_<locale>.csv` lands, this lookup
  // will branch on `locale` and consult the alias table first.
  *out_localized = canonical;
  return 0;
}

extern "C" fm_status_t fm_function_canonicalize(const char* localized_name, std::int32_t locale,
                                                const char** out_canonical) {
  clear_last_error();
  if (localized_name == nullptr || out_canonical == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_function_canonicalize: NULL argument");
  }
  *out_canonical = nullptr;
  if (locale < FM_LOCALE_EN_US || locale > FM_LOCALE_JA_JP) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_function_canonicalize: invalid locale",
                             "locale=" + std::to_string(locale));
  }
  // Alias table not yet populated - fall through to a case-insensitive
  // canonical-name match across the registry, lazy, and special-form
  // tables so every enumerated function canonicalizes.
  const char* canonical = resolve_canonical_name(std::string_view(localized_name));
  if (canonical == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_function_canonicalize: unknown function",
                             std::string("localized_name=") + localized_name);
  }
  *out_canonical = canonical;
  return 0;
}
