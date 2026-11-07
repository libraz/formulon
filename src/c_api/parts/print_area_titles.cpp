//
// C ABI - print area and print titles.
//
// Both settings live in workbook-scope defined names (`_xlnm.Print_Area` /
// `_xlnm.Print_Titles`) rather than in the worksheet part, so they are read
// and written through the defined-name table. The helpers here render the
// sheet-qualified A1 text those names carry.
//
// @size-budget: 5 KB

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "parser/ast_format.h"
#include "parser/reference.h"
#include "print/print_area.h"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_column_span;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_row_span;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_last_error;

namespace {

using formulon::print::kPrintAreaName;
using formulon::print::kPrintTitlesName;

/// Returns the sheet-qualified prefix (`Sheet1!` / `'集計 2026'!`) for a
/// defined-name formula, quoting per Excel's rules.
std::string sheet_qualifier(std::string_view sheet_name) {
  std::string out;
  formulon::parser::append_sheet_name(sheet_name, /*force_quote=*/false, out);
  out.push_back('!');
  return out;
}

}  // namespace

/* -------------------------------------------------------------------------- */
/* Print area and titles                                                      */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_set_print_area(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* ranges,
                                               size_t range_count) {
  clear_last_error();
  if (range_count != 0U && ranges == nullptr) {
    return null_argument("ranges");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  std::string formula;
  if (range_count != 0U) {
    if (auto rc = formulon::c_api::parts::check_rect_list(ranges, range_count); rc != 0) {
      return rc;
    }
    const std::string qualifier = sheet_qualifier(wb->workbook().sheet(sheet_index).name());
    for (std::size_t i = 0; i < range_count; ++i) {
      if (!formula.empty()) {
        formula.push_back(',');
      }
      formula.append(qualifier);
      formula.append(formulon::c_api::parts::format_rect_a1({}, ranges[i], /*absolute=*/true));
    }
  }
  auto result = wb->workbook().set_defined_name_scoped(std::string(kPrintAreaName), std::move(formula),
                                                       static_cast<std::int32_t>(sheet_index));
  if (!result) {
    return set_last_error(result.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_print_area(const fm_workbook_t* wb, size_t sheet_index,
                                               const fm_cell_rect_t** out_ranges, size_t* out_range_count) {
  clear_last_error();
  if (out_ranges == nullptr || out_range_count == nullptr) {
    return null_argument("out_ranges");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto areas = formulon::print::resolve_print_area(wb->workbook(), static_cast<std::uint32_t>(sheet_index));
  if (!areas) {
    return set_last_error(areas.error());
  }
  std::vector<formulon::MergeRange> rects;
  rects.reserve(areas.value().size());
  for (const formulon::print::CellRange& range : areas.value()) {
    rects.push_back({range.first_row, range.first_col, range.last_row, range.last_col});
  }
  fresh_read_scratch(wb);
  *out_ranges = formulon::c_api::parts::publish_rects(wb, rects);
  *out_range_count = rects.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_set_print_titles(fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_print_titles_t* titles) {
  clear_last_error();
  if (titles != nullptr) {
    if (auto rc = check_record_size(*titles); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const bool has_rows = titles != nullptr && titles->has_rows != 0;
  const bool has_cols = titles != nullptr && titles->has_cols != 0;
  if (has_rows) {
    if (auto rc = check_row_span(titles->first_row, titles->last_row); rc != 0) {
      return rc;
    }
  }
  if (has_cols) {
    if (auto rc = check_column_span(titles->first_col, titles->last_col); rc != 0) {
      return rc;
    }
  }
  std::string formula;
  if (has_rows || has_cols) {
    const std::string qualifier = sheet_qualifier(wb->workbook().sheet(sheet_index).name());
    // Excel lists rows before columns; the resolver accepts either order,
    // but matching Excel keeps a diffed worksheet quiet.
    if (has_rows) {
      formula.append(qualifier);
      formula.append("$" + std::to_string(static_cast<std::uint64_t>(titles->first_row) + 1U) + ":$" +
                     std::to_string(static_cast<std::uint64_t>(titles->last_row) + 1U));
    }
    if (has_cols) {
      if (!formula.empty()) {
        formula.push_back(',');
      }
      formula.append(qualifier);
      formula.push_back('$');
      formulon::a1::append_column_letters(formula, titles->first_col);
      formula.append(":$");
      formulon::a1::append_column_letters(formula, titles->last_col);
    }
  }
  auto result = wb->workbook().set_defined_name_scoped(std::string(kPrintTitlesName), std::move(formula),
                                                       static_cast<std::int32_t>(sheet_index));
  if (!result) {
    return set_last_error(result.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_print_titles(const fm_workbook_t* wb, size_t sheet_index, fm_print_titles_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto titles = formulon::print::resolve_print_titles(wb->workbook(), static_cast<std::uint32_t>(sheet_index));
  if (!titles) {
    return set_last_error(titles.error());
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_print_titles_t{};
  out->struct_size = struct_size;
  if (titles.value().repeat_rows.has_value()) {
    out->has_rows = 1;
    out->first_row = titles.value().repeat_rows->first;
    out->last_row = titles.value().repeat_rows->second;
  }
  if (titles.value().repeat_cols.has_value()) {
    out->has_cols = 1;
    out->first_col = titles.value().repeat_cols->first;
    out->last_col = titles.value().repeat_cols->second;
  }
  return 0;
}
