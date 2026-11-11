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

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_enum_domain;
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
// placed by `place` over `caller_text`, the text the caller passed. A leading
// `=` is part of the offsets, as the parser counts from the body after it.
template <typename PlaceFn>
fm_status_t check_parses(std::string_view text, std::string_view caller_text, PlaceFn place) {
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
  return formula_parse_failed(caller_text, place(failure.range, failure), "field=formula");
}

// Reads `spec` (NULL = A1, no base) into `*out` after the checks of `fm_workbook_parse_reference`.
fm_status_t read_notation(const fm_reference_notation_spec_t* spec, formulon::eval::FormulaNotation* out) {
  const fm_reference_notation_spec_t plain = FM_STRUCT_INIT(fm_reference_notation_spec_t);
  if (spec == nullptr) {
    spec = &plain;
  } else if (auto rc = check_record_size(*spec); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec->r1c1, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "r1c1"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec->has_base, 0, 1, "has_base"); rc != 0) {
    return rc;
  }
  if (spec->has_base == 1) {
    if (auto rc = check_cell_coord(spec->base_row, spec->base_col); rc != 0) {
      return rc;
    }
  }
  out->r1c1 = spec->r1c1 == FM_TRISTATE_TRUE;
  out->has_base = spec->has_base == 1;
  out->base_row = spec->base_row;
  out->base_col = spec->base_col;
  return 0;
}

// Shared body of `fm_localize_formula` / `fm_canonicalize_formula`.
fm_status_t convert_formula(const char* formula, const char* profile_id, const fm_reference_notation_spec_t* spec,
                            bool localize, const char** out_formula) {
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
  formulon::eval::FormulaNotation notation;
  if (auto rc = read_notation(spec, &notation); rc != 0) {
    return rc;
  }
  const std::string_view source(formula);
  thread_local std::string buffer;
  // Both rewrites read the caller's text, so their failures are already in its coordinates.
  auto mapped = localize ? formulon::eval::localize_formula_text_r1c1(source, profile, notation)
                         : formulon::eval::canonicalize_formula_text_r1c1(source, profile, notation);
  if (!mapped) {
    return formula_parse_failed(source, mapped.error(), "field=formula");
  }
  const formulon::eval::MappedFormulaText& text = mapped.value();
  // The A1 canonicalize rewrite does not parse its result.
  if (!localize && !notation.r1c1) {
    if (auto rc = check_parses(text.text, source,
                               [&](const formulon::parser::TextRange& range, formulon::parser::ParseFailure f) {
                                 f.range = text.source_range(source, range);
                                 return f;
                               });
        rc != 0) {
      return rc;
    }
  }
  buffer = text.text;
  *out_formula = buffer.c_str();
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_localize_formula(const char* formula, const char* profile_id,
                                           const fm_reference_notation_spec_t* spec, const char** out_formula) {
  return convert_formula(formula, profile_id, spec, /*localize=*/true, out_formula);
}

extern "C" fm_status_t fm_canonicalize_formula(const char* formula, const char* profile_id,
                                               const fm_reference_notation_spec_t* spec, const char** out_formula) {
  return convert_formula(formula, profile_id, spec, /*localize=*/false, out_formula);
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
