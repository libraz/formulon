//
// C ABI - reference text: parsing with the grammar INDIRECT reads under the
// workbook's profile, and A1 formatting with the stored-formula quoting rule.

#include "parser/reference.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "eval/a1_parse.h"
#include "eval/eval_profile_scope.h"
#include "parser/parse_error.h"
#include "parser/parse_status.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/error.h"
#include "utils/utf8_length.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::format_rect_a1;
using formulon::c_api::parts::formula_parse_failed;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::store_cstr;

extern "C" fm_status_t fm_workbook_parse_reference(const fm_workbook_t* wb, const char* text,
                                                   const fm_parse_reference_spec_t* spec,
                                                   fm_parsed_reference_info_t* out) {
  clear_last_error();
  if (wb == nullptr || text == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({text}); rc != 0) {
    return rc;
  }
  fm_parse_reference_spec_t plain = FM_STRUCT_INIT(fm_parse_reference_spec_t);
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
  const bool r1c1 = spec->r1c1 == FM_TRISTATE_TRUE;
  const std::int32_t has_base = spec->has_base;
  const std::uint32_t base_row = spec->base_row;
  const std::uint32_t base_col = spec->base_col;
  if (has_base == 1) {
    if (auto rc = check_cell_coord(base_row, base_col); rc != 0) {
      return rc;
    }
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  out->struct_size = struct_size;
  namespace refs = formulon::eval::refs_internal;
  refs::A1Parse parsed;
  {
    const formulon::eval::ScopedEvalProfile profile_scope(wb->workbook().excel_profile());
    if (r1c1) {
      refs::R1C1Base base;
      base.present = has_base == 1;
      base.row = base_row;
      base.col = base_col;
      parsed = refs::parse_r1c1_ref(text, base);
    } else {
      parsed = refs::parse_a1_ref(text);
    }
  }
  if (!parsed.valid || formulon::utf16_units_in(parsed.sheet) > formulon::sheet_names::kMaxNameUnits) {
    // The grammar reports no position, so the failure covers the whole text.
    formulon::parser::ParseFailure failure;
    failure.code = formulon::parser::ParseErrorCode::InvalidReference;
    failure.range.end = formulon::utf16_units_in(text);
    failure.message = formulon::parser::default_message(failure.code);
    return formula_parse_failed(text, failure, std::string());
  }
  const std::uint32_t row2 = parsed.is_range ? parsed.row2 : parsed.row;
  const std::uint32_t col2 = parsed.is_range ? parsed.col2 : parsed.col;
  out->rect = fm_cell_rect_t{std::min(parsed.row, row2), std::min(parsed.col, col2), std::max(parsed.row, row2),
                             std::max(parsed.col, col2)};
  if (!parsed.sheet.empty()) {
    // The parser's view points at thread-local scratch the next parse reuses.
    out->sheet_name = store_cstr(fresh_read_scratch(wb), parsed.sheet);
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_format_reference(const fm_workbook_t* wb, const char* sheet_name,
                                                    const fm_cell_rect_t* range, int32_t absolute,
                                                    const char** out_text) {
  clear_last_error();
  if (wb == nullptr || range == nullptr || out_text == nullptr) {
    return null_argument();
  }
  const std::uint32_t row0 = range->first_row;
  const std::uint32_t col0 = range->first_col;
  const std::uint32_t row1 = range->last_row;
  const std::uint32_t col1 = range->last_col;
  if (auto rc = check_utf8({sheet_name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row0, col0); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row1, col1); rc != 0) {
    return rc;
  }
  if (row0 > row1 || col0 > col1) {
    return invalid_argument("range first_row and first_col must not exceed last_row and last_col");
  }
  if (auto rc = check_enum_domain(absolute, 0, 1, "absolute"); rc != 0) {
    return rc;
  }
  *out_text = publish_scratch(
      wb,
      format_rect_a1(sheet_name != nullptr ? std::string_view(sheet_name) : std::string_view(), *range, absolute == 1));
  return 0;
}
