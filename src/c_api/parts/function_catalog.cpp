//
// C ABI - function catalog metadata.
//
// Surfaces the engine's canonical-name + arity data through the C ABI
// so JS / Python autocomplete UIs can enumerate Formulon functions
// without reaching into the engine internals.
//
// `description` / `signature_template` are always `NULL`. That is the
// design, not an unfinished table: the engine ships no human-readable
// function text, and a host merges its own document over this structural
// result at display time (`docs/function-metadata-schema.md`). Anything
// that would populate these fields from inside the engine contradicts
// that contract. Localize / canonicalize read the function-name table
// FORMULATEXT uses (`eval/locale_function_names.h`) through the profile's
// locale: de-DE, fr-FR, es-ES, es-MX, pt-BR, ru-RU, it-IT and nl-NL localize
// the catalog, ja-JP three functions and zh-TW one; en-US, zh-CN, ko-KR and
// th-TH show the canonical name.
//
// A runtime-recognised function name comes from one of three sources:
// the eager `FunctionRegistry`, the tree walker's lazy-dispatch table
// (`eval::lazy_form_names()` — IF / XLOOKUP / SUMIFS / VLOOKUP / ...),
// or the parser's special forms (`eval::parser_special_form_names()` —
// LET / LAMBDA). All three are enumerated and de-duplicated so the
// catalog matches what the evaluator actually accepts; a single name
// resolves through `eval::resolve_builtin_function_name`, and its arity
// through `eval::builtin_arity`, the bounds formula entry enforces.
//
// The sorted name list is computed on first use and cached for the
// process lifetime - order matters for UIs that expect deterministic
// enumeration, and the underlying iteration hooks do not promise any
// order.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "eval/builtin_names.h"
#include "eval/formula_localize.h"
#include "eval/function_registry.h"
#include "eval/special_forms_catalog.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "utils/error.h"
#include "utils/index_sort.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_profile_id;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;

namespace {

struct FunctionAvailabilityEntry {
  std::string_view name;
  fm_function_availability_t availability;
};

fm_function_availability_t function_availability(std::string_view canonical_name) {
  static constexpr FunctionAvailabilityEntry kEntries[] = {
      {"CELL", FM_FUNCTION_AVAILABILITY_ENVIRONMENT_BOUND},
      {"COPILOT", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"CUBEKPIMEMBER", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"CUBEMEMBER", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"CUBEMEMBERPROPERTY", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"CUBERANKEDMEMBER", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"CUBESET", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"CUBESETCOUNT", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"CUBEVALUE", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"DETECTLANGUAGE", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"IMAGE", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"INFO", FM_FUNCTION_AVAILABILITY_ENVIRONMENT_BOUND},
      {"PY", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"RTD", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"STOCKHISTORY", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"TRANSLATE", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
      {"WEBSERVICE", FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB},
  };
  for (const auto& entry : kEntries) {
    if (entry.name == canonical_name) {
      return entry.availability;
    }
  }
  return FM_FUNCTION_AVAILABILITY_IMPLEMENTED;
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
    formulon::sort_by_index(out, [](const std::string& lhs, const std::string& rhs) { return lhs < rhs; });
    // A name can appear in more than one source (e.g. registry + lazy); keep a
    // single entry so the enumeration count matches the recognised set.
    out.erase(std::unique(out.begin(), out.end()), out.end());
    return out;
  }();
  return names;
}

}  // namespace

extern "C" fm_status_t fm_get_function_metadata(const char* name, int32_t* out_present, fm_function_metadata_t* out) {
  clear_last_error();
  if (name == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_function_metadata_t{};
  out->struct_size = struct_size;
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  const char* canonical = formulon::eval::resolve_builtin_function_name(std::string_view(name));
  if (canonical == nullptr) {
    return 0;
  }
  *out_present = 1;
  // `canonical_name` is static storage; callers may hold the result indefinitely.
  out->canonical_name = canonical;
  formulon::eval::builtin_arity(canonical, &out->min_arity, &out->max_arity);
  out->availability = function_availability(std::string_view(canonical));
  // Display text is host-supplied, never engine-owned: see the file header.
  out->signature_template = nullptr;
  out->description = nullptr;
  return 0;
}

extern "C" fm_status_t fm_get_function_count(std::size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  *out_count = sorted_function_names().size();
  return 0;
}

extern "C" fm_status_t fm_get_function_name_at(std::size_t idx, const char** out_name) {
  clear_last_error();
  if (out_name == nullptr) {
    return null_argument("out_name");
  }
  const auto& names = sorted_function_names();
  if (auto rc = check_index(idx, names.size(), "index"); rc != 0) {
    return rc;
  }
  *out_name = names[idx].c_str();
  return 0;
}

namespace {

// Shared body of `fm_localize_function_name` / `fm_canonicalize_function_name`.
// `localize` maps canonical -> localized; otherwise a localized spelling is
// matched first and the canonical name is accepted as a fallback.
fm_status_t resolve_function_name(const char* name_field, const char* name, const char* profile_id, bool localize,
                                  const char** out) {
  *out = nullptr;
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, &profile); rc != 0) {
    return rc;
  }
  const formulon::ExcelLocale locale = profile.locale;
  std::string_view lookup(name);
  if (!localize) {
    if (const char* canonical = formulon::eval::canonical_function_name(locale, lookup); canonical != nullptr) {
      lookup = canonical;
    }
  }
  const char* canonical = formulon::eval::resolve_builtin_function_name(lookup);
  if (canonical == nullptr) {
    return invalid_argument("unknown function", std::string(name_field) + "=" + name);
  }
  const char* localized = localize ? formulon::eval::localized_function_name(locale, canonical) : nullptr;
  *out = localized != nullptr ? localized : canonical;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_localize_function_name(const char* canonical_name, const char* profile_id,
                                                 const char** out_localized) {
  clear_last_error();
  if (canonical_name == nullptr || profile_id == nullptr || out_localized == nullptr) {
    return null_argument();
  }
  return resolve_function_name("canonical_name", canonical_name, profile_id,
                               /*localize=*/true, out_localized);
}

extern "C" fm_status_t fm_canonicalize_function_name(const char* localized_name, const char* profile_id,
                                                     const char** out_canonical) {
  clear_last_error();
  if (localized_name == nullptr || profile_id == nullptr || out_canonical == nullptr) {
    return null_argument();
  }
  return resolve_function_name("localized_name", localized_name, profile_id,
                               /*localize=*/false, out_canonical);
}
