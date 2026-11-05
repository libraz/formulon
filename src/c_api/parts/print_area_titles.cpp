//
// C ABI - print area and print titles.
//
// Both settings live in workbook-scope defined names (`_xlnm.Print_Area` /
// `_xlnm.Print_Titles`) rather than in the worksheet part, so they are read
// and written through the defined-name table. The helpers here render and
// parse the sheet-qualified A1 text those names carry.
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
#include "parser/ast_format.h"
#include "parser/reference.h"
#include "print/print_area.h"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;

namespace {

using formulon::Sheet;

// OOXML built-in defined names for the two print settings Excel stores
// outside the worksheet part.
constexpr std::string_view kPrintAreaName = "_xlnm.Print_Area";
constexpr std::string_view kPrintTitlesName = "_xlnm.Print_Titles";

/// Returns the sheet-qualified prefix (`Sheet1!` / `'集計 2026'!`) for a
/// defined-name formula, quoting per Excel's rules.
std::string sheet_qualifier(std::string_view sheet_name) {
  std::string out;
  formulon::parser::append_sheet_name(sheet_name, /*force_quote=*/false, out);
  out.push_back('!');
  return out;
}

/// Rewrites one A1 token with every column and row anchored.
///
/// Works on the authored text rather than a parsed rectangle so a
/// whole-axis area keeps its shape: `A:D` becomes `$A:$D`, not the
/// `$A$1:$D$1048576` a round trip through `CellRange` would produce.
/// Existing anchors are not doubled.
std::string absolutise_a1_token(std::string_view token) {
  // A run is a maximal stretch of one character class. `A1` is two runs
  // (the column and the row), each needing its own anchor; `AB` and `12`
  // are one run each.
  enum class Run { kNone, kLetters, kDigits };

  std::string out;
  out.reserve(token.size() + 4U);
  Run run = Run::kNone;
  bool anchor_pending = false;
  for (const char chr : token) {
    if (chr == '$') {
      anchor_pending = true;
      run = Run::kNone;
      out.push_back(chr);
      continue;
    }
    Run kind = Run::kNone;
    if ((chr >= 'A' && chr <= 'Z') || (chr >= 'a' && chr <= 'z')) {
      kind = Run::kLetters;
    } else if (chr >= '0' && chr <= '9') {
      kind = Run::kDigits;
    }
    if (kind == Run::kNone) {
      run = Run::kNone;
      anchor_pending = false;
      out.push_back(chr);
      continue;
    }
    if (kind != run) {
      if (!anchor_pending) {
        out.push_back('$');
      }
      run = kind;
    }
    anchor_pending = false;
    out.push_back(chr);
  }
  return out;
}

/// A span endpoint reader: `formulon::a1::parse_uint` or `parse_column_letters`.
using SpanEndParser = bool (*)(std::string_view, std::size_t*, std::uint32_t*);

/// True when `token` is a whole-row (`1:2`) or whole-column (`A:A`) span
/// whose endpoints `parse_end` reads within `1..max`.
bool is_span(std::string_view token, SpanEndParser parse_end, std::uint32_t max) {
  const std::size_t colon = token.find(':');
  if (colon == std::string_view::npos) {
    return false;
  }
  const auto is_end = [&](std::string_view part) {
    std::size_t pos = 0;
    std::uint32_t value = 0;
    if (!part.empty() && part.front() == '$') {
      part.remove_prefix(1);
    }
    return parse_end(part, &pos, &value) && pos == part.size() && value >= 1U && value <= max;
  };
  return is_end(token.substr(0, colon)) && is_end(token.substr(colon + 1));
}

}  // namespace

/* -------------------------------------------------------------------------- */
/* Print area and titles                                                      */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_set_print_area(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* ranges,
                                               size_t range_count) {
  static constexpr const char* kFn = "fm_sheet_set_print_area";
  clear_last_error();
  if (range_count != 0U && ranges == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=ranges");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  std::string formula;
  if (range_count != 0U) {
    if (auto rc = formulon::c_api::parts::check_rect_list(ranges, range_count, kFn); rc != 0) {
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
  static constexpr const char* kFn = "fm_sheet_get_print_area";
  clear_last_error();
  if (out_ranges == nullptr || out_range_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=out_ranges");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
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

extern "C" fm_status_t fm_sheet_set_print_titles(fm_workbook_t* wb, size_t sheet_index, const char* repeat_rows,
                                                 const char* repeat_cols) {
  static constexpr const char* kFn = "fm_sheet_set_print_titles";
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8(kFn, {repeat_rows, repeat_cols}); rc != 0) {
    return rc;
  }
  const std::string_view rows(repeat_rows == nullptr ? "" : repeat_rows);
  const std::string_view cols(repeat_cols == nullptr ? "" : repeat_cols);
  std::string formula;
  if (!rows.empty() || !cols.empty()) {
    const std::string qualifier = sheet_qualifier(wb->workbook().sheet(sheet_index).name());
    // Excel lists rows before columns; the resolver accepts either order,
    // but matching Excel keeps a diffed worksheet quiet.
    if (!rows.empty()) {
      if (!is_span(rows, &formulon::a1::parse_uint, Sheet::kMaxRows)) {
        return set_binding_error(formulon::FormulonErrorCode::kPrintInvalidArea, kFn,
                                 "repeat_rows=" + std::string(rows));
      }
      formula.append(qualifier);
      formula.append(absolutise_a1_token(rows));
    }
    if (!cols.empty()) {
      if (!is_span(cols, &formulon::a1::parse_column_letters, Sheet::kMaxCols)) {
        return set_binding_error(formulon::FormulonErrorCode::kPrintInvalidArea, kFn,
                                 "repeat_cols=" + std::string(cols));
      }
      if (!formula.empty()) {
        formula.push_back(',');
      }
      formula.append(qualifier);
      formula.append(absolutise_a1_token(cols));
    }
  }
  auto result = wb->workbook().set_defined_name_scoped(std::string(kPrintTitlesName), std::move(formula),
                                                       static_cast<std::int32_t>(sheet_index));
  if (!result) {
    return set_last_error(result.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_print_titles(const fm_workbook_t* wb, size_t sheet_index,
                                                 const char** out_repeat_rows, const char** out_repeat_cols) {
  static constexpr const char* kFn = "fm_sheet_get_print_titles";
  clear_last_error();
  if (out_repeat_rows == nullptr || out_repeat_cols == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn,
                             "arg=out_repeat_rows|out_repeat_cols");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  auto titles = formulon::print::resolve_print_titles(wb->workbook(), static_cast<std::uint32_t>(sheet_index));
  if (!titles) {
    return set_last_error(titles.error());
  }
  std::string rows;
  if (titles.value().repeat_rows.has_value()) {
    const auto& span = *titles.value().repeat_rows;
    rows.append(std::to_string(static_cast<std::uint64_t>(span.first) + 1U));
    rows.push_back(':');
    rows.append(std::to_string(static_cast<std::uint64_t>(span.second) + 1U));
  }
  std::string cols;
  if (titles.value().repeat_cols.has_value()) {
    const auto& span = *titles.value().repeat_cols;
    formulon::a1::append_column_letters(cols, span.first);
    cols.push_back(':');
    formulon::a1::append_column_letters(cols, span.second);
  }
  // Both pointers are handed out from one refresh: clearing between the two
  // pushes would dangle the first.
  TextStore& store = fresh_read_scratch(wb);
  *out_repeat_rows = store_cstr(store, rows);
  *out_repeat_cols = store_cstr(store, cols);
  return 0;
}
