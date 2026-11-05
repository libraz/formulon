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
#include "eval/a1_parse.h"
#include "eval/eval_profile_scope.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/error.h"
#include "utils/utf8_length.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::store_cstr;

namespace {

// One endpoint of a whole-column / whole-row pair: the text before the `:`
// `format_a1` writes between the two copies of the axis.
std::string whole_axis_endpoint(const formulon::parser::Reference& ref) {
  std::string text = formulon::parser::format_a1(ref);
  text.resize(text.rfind(':'));
  return text;
}

}  // namespace

extern "C" fm_status_t fm_workbook_parse_reference(const fm_workbook_t* wb, const char* text, int32_t r1c1,
                                                   int32_t has_base, uint32_t base_row, uint32_t base_col,
                                                   fm_parsed_reference_t* out) {
  static constexpr const char* kApi = "fm_workbook_parse_reference";
  clear_last_error();
  if (wb == nullptr || text == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  if (auto rc = check_utf8(kApi, {text}); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(r1c1, 1, kApi, "r1c1"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(has_base, 1, kApi, "has_base"); rc != 0) {
    return rc;
  }
  if (has_base == 1) {
    if (auto rc = check_cell_coord(base_row, base_col, kApi); rc != 0) {
      return rc;
    }
  }
  *out = fm_parsed_reference_t{};
  namespace refs = formulon::eval::refs_internal;
  refs::A1Parse parsed;
  {
    const formulon::eval::ScopedEvalProfile profile_scope(wb->workbook().excel_profile());
    if (r1c1 == 1) {
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
    return 0;
  }
  const std::uint32_t row2 = parsed.is_range ? parsed.row2 : parsed.row;
  const std::uint32_t col2 = parsed.is_range ? parsed.col2 : parsed.col;
  out->valid = 1;
  out->rect = fm_cell_rect_t{std::min(parsed.row, row2), std::min(parsed.col, col2), std::max(parsed.row, row2),
                             std::max(parsed.col, col2)};
  if (!parsed.sheet.empty()) {
    // The parser's view points at thread-local scratch the next parse reuses.
    out->has_sheet = 1;
    out->sheet_name = store_cstr(fresh_read_scratch(wb), parsed.sheet);
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_format_reference(const fm_workbook_t* wb, const char* sheet_name,
                                                    const fm_cell_rect_t* range, int32_t absolute,
                                                    const char** out_text) {
  static constexpr const char* kApi = "fm_workbook_format_reference";
  clear_last_error();
  if (wb == nullptr || range == nullptr || out_text == nullptr) {
    return null_argument(kApi);
  }
  const std::uint32_t row0 = range->first_row;
  const std::uint32_t col0 = range->first_col;
  const std::uint32_t row1 = range->last_row;
  const std::uint32_t col1 = range->last_col;
  if (auto rc = check_utf8(kApi, {sheet_name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row0, col0, kApi); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row1, col1, kApi); rc != 0) {
    return rc;
  }
  if (row0 > row1 || col0 > col1) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(kApi) + ": inverted rectangle").c_str());
  }
  if (auto rc = check_enum_domain(absolute, 1, kApi, "absolute"); rc != 0) {
    return rc;
  }
  formulon::parser::Reference first;
  first.sheet = sheet_name != nullptr ? std::string_view(sheet_name) : std::string_view();
  first.row = row0;
  first.col = col0;
  first.row_abs = absolute == 1;
  first.col_abs = absolute == 1;
  formulon::parser::Reference second = first;
  second.sheet = std::string_view();
  second.row = row1;
  second.col = col1;
  std::string text;
  if (row0 == 0U && row1 == formulon::Sheet::kMaxRows - 1U) {
    first.is_full_col = true;
    second.is_full_col = true;
    text = whole_axis_endpoint(first) + ":" + whole_axis_endpoint(second);
  } else if (col0 == 0U && col1 == formulon::Sheet::kMaxCols - 1U) {
    first.is_full_row = true;
    second.is_full_row = true;
    text = whole_axis_endpoint(first) + ":" + whole_axis_endpoint(second);
  } else if (row0 == row1 && col0 == col1) {
    text = formulon::parser::format_a1(first);
  } else {
    text = formulon::parser::format_a1(first) + ":" + formulon::parser::format_a1(second);
  }
  *out_text = publish_scratch(wb, std::move(text));
  return 0;
}
