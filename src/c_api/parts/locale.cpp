//
// C ABI - locale facts and formula text conversion.
//
// Every entry takes a profile id string (`parse_excel_profile_id`), so one
// identifier selects the locale and the Excel-backed / estimated status. Facts
// are read from `LocaleFacts`; the string views behind them are literals, so
// they are NUL-terminated and live for the process.

#include <array>
#include <cstddef>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "eval/builtin_names.h"
#include "eval/formula_localize.h"
#include "excel_locale.h"
#include "excel_profile.h"
#include "parser/ast.h"
#include "parser/parse_status.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/formula_text.h"
#include "utils/utf8_length.h"
#include "value.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_profile_id;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::formula_parse_failed;
using formulon::c_api::parts::null_argument;

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

// Parses `text` (stored spelling) and, on a syntax error, publishes the failure
// placed by `place` over the text the caller passed. A leading `=` is part of
// the offsets, as the parser counts from the body after it.
template <typename PlaceFn>
fm_status_t check_parses(std::string_view text, PlaceFn place) {
  const std::string_view body = formulon::strip_formula_prefix(text);
  if (body.empty() && text.empty()) {
    return 0;
  }
  formulon::Arena arena;
  formulon::parser::ParseFailure failure;
  if (formulon::eval::parse_formula_entry_with_diagnostics(body, arena, &failure) != nullptr) {
    return 0;
  }
  const std::uint32_t prefix = formulon::utf16_units_in(text.substr(0, text.size() - body.size()));
  failure.range.start += prefix;
  failure.range.end += prefix;
  if (failure.range.line == 1U) {
    failure.range.column += prefix;
  }
  return formula_parse_failed(text, place(failure.range, failure), "field=formula");
}

// Shared body of `fm_localize_formula` / `fm_canonicalize_formula`.
fm_status_t convert_formula(const char* formula, const char* profile_id, bool localize, const char** out_formula) {
  clear_last_error();
  if (formula == nullptr || profile_id == nullptr || out_formula == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({formula}); rc != 0) {
    return rc;
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, &profile); rc != 0) {
    return rc;
  }
  thread_local std::string buffer;
  if (localize) {
    // The input is the canonical text, so the failure is already in the caller's coordinates.
    if (auto rc = check_parses(formula,
                               [](const formulon::parser::TextRange&, formulon::parser::ParseFailure f) { return f; });
        rc != 0) {
      return rc;
    }
    buffer = formulon::eval::localize_formula_text(formula, profile);
  } else {
    const formulon::eval::MappedFormulaText mapped = formulon::eval::canonicalize_formula_text_mapped(formula, profile);
    const std::string_view source(formula);
    if (auto rc = check_parses(mapped.text,
                               [&](const formulon::parser::TextRange& range, formulon::parser::ParseFailure f) {
                                 f.range = mapped.source_range(source, range);
                                 return f;
                               });
        rc != 0) {
      return rc;
    }
    buffer = mapped.text;
  }
  *out_formula = buffer.c_str();
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_localize_formula(const char* formula, const char* profile_id, const char** out_formula) {
  return convert_formula(formula, profile_id, /*localize=*/true, out_formula);
}

extern "C" fm_status_t fm_canonicalize_formula(const char* formula, const char* profile_id, const char** out_formula) {
  return convert_formula(formula, profile_id, /*localize=*/false, out_formula);
}

extern "C" fm_status_t fm_get_locale_facts(const char* profile_id, fm_locale_facts_info_t* out) {
  clear_last_error();
  if (profile_id == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, &profile); rc != 0) {
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
  out->collation_measured = profile.host == formulon::ExcelHost::kMac365 ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_get_locale_error_name_count(size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  *out_count = formulon::kErrorNameCount;
  return 0;
}

extern "C" fm_status_t fm_get_locale_error_name_at(const char* profile_id, size_t index, const char** out_canonical,
                                                   const char** out_localized, int32_t* out_measured) {
  clear_last_error();
  if (profile_id == nullptr || out_canonical == nullptr || out_localized == nullptr || out_measured == nullptr) {
    return null_argument();
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, &profile); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(index, formulon::kErrorNameCount, "index"); rc != 0) {
    return rc;
  }
  *out_canonical = formulon::kErrorTable[index].display_name;
  *out_localized = formulon::locale_facts(profile).error_names[index].data();
  *out_measured = formulon::is_measured_profile(profile) && formulon::error_name_measured(index) ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_get_profile_id_count(size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  *out_count = formulon::detail::kExcelProfileIds.size();
  return 0;
}

extern "C" fm_status_t fm_get_profile_id_at(size_t index, const char** out_profile_id) {
  clear_last_error();
  if (out_profile_id == nullptr) {
    return null_argument("out_profile_id");
  }
  if (auto rc = check_index(index, formulon::detail::kExcelProfileIds.size(), "index"); rc != 0) {
    return rc;
  }
  *out_profile_id = formulon::detail::kExcelProfileIds[index].id;
  return 0;
}
