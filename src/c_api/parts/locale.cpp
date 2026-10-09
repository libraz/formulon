//
// C ABI - locale facts and formula text conversion.
//
// Every entry takes a profile id string (`parse_excel_profile_id`), so one
// identifier selects the locale and the measured / estimated status. Facts
// are read from `LocaleFacts`; the string views behind them are literals, so
// they are NUL-terminated and live for the process.

#include <array>
#include <cstddef>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/formula_localize.h"
#include "excel_locale.h"
#include "excel_profile.h"
#include "utils/error.h"
#include "value.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_profile_id;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

namespace {

// "x\0" pairs for every byte value, so a one-character separator has a
// NUL-terminated process-static spelling.
constexpr std::array<char, 512> make_one_char_table() {
  std::array<char, 512> table{};
  for (std::size_t i = 0; i < 256; ++i) {
    table[2 * i] = static_cast<char>(i);
  }
  return table;
}

constexpr std::array<char, 512> kOneChar = make_one_char_table();

const char* one_char(char c) noexcept {
  return &kOneChar[2 * static_cast<std::size_t>(static_cast<unsigned char>(c))];
}

// Shared body of `fm_formula_localize` / `fm_formula_canonicalize`.
fm_status_t convert_formula(const char* api, const char* formula, const char* profile_id, bool localize,
                            const char** out_formula) {
  clear_last_error();
  if (formula == nullptr || profile_id == nullptr || out_formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(api) + ": NULL argument").c_str());
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, api, &profile); rc != 0) {
    return rc;
  }
  thread_local std::string buffer;
  buffer = localize ? formulon::eval::localize_formula_text(formula, profile)
                    : formulon::eval::canonicalize_formula_text(formula, profile);
  *out_formula = buffer.c_str();
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_formula_localize(const char* formula, const char* profile_id, const char** out_formula) {
  return convert_formula("fm_formula_localize", formula, profile_id, /*localize=*/true, out_formula);
}

extern "C" fm_status_t fm_formula_canonicalize(const char* formula, const char* profile_id, const char** out_formula) {
  return convert_formula("fm_formula_canonicalize", formula, profile_id, /*localize=*/false, out_formula);
}

extern "C" fm_status_t fm_locale_facts(const char* profile_id, fm_locale_facts_t* out) {
  clear_last_error();
  if (profile_id == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_locale_facts: NULL argument");
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, "fm_locale_facts", &profile); rc != 0) {
    return rc;
  }
  const formulon::LocaleFacts& facts = formulon::locale_facts(profile);
  out->decimal_separator = one_char(facts.decimal_separator);
  out->group_separator = one_char(facts.group_separator);
  out->list_separator = one_char(facts.list_separator);
  out->array_column_separator = one_char(facts.array_column_separator);
  out->array_row_separator = one_char(facts.array_row_separator);
  out->true_name = facts.true_name.data();
  out->false_name = facts.false_name.data();
  out->date_order = static_cast<int32_t>(facts.date_order);
  out->currency_symbol = facts.currency.symbol.data();
  out->currency_suffix = facts.currency.suffix ? 1 : 0;
  out->currency_space = facts.currency.space ? 1 : 0;
  out->currency_default_decimals = static_cast<int32_t>(facts.currency.default_decimals);
  out->measured = formulon::is_measured_profile(profile) ? 1 : 0;
  return 0;
}

extern "C" size_t fm_locale_error_name_count(void) {
  return formulon::kErrorNameCount;
}

extern "C" fm_status_t fm_locale_error_name(const char* profile_id, size_t error_index, const char** out_canonical,
                                            const char** out_localized, int32_t* out_measured) {
  clear_last_error();
  if (profile_id == nullptr || out_canonical == nullptr || out_localized == nullptr || out_measured == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_locale_error_name: NULL argument");
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, "fm_locale_error_name", &profile); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(error_index, formulon::kErrorNameCount, "fm_locale_error_name", "error_index"); rc != 0) {
    return rc;
  }
  *out_canonical = formulon::kErrorTable[error_index].display_name;
  *out_localized = formulon::locale_facts(profile).error_names[error_index].data();
  *out_measured = formulon::is_measured_profile(profile) && formulon::error_name_measured(error_index) ? 1 : 0;
  return 0;
}
