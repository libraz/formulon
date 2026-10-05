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
#include "parser/reference.h"
#include "print/print_area.h"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

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
  if (!formulon::parser::sheet_name_needs_quoting(sheet_name)) {
    out.append(sheet_name);
    out.push_back('!');
    return out;
  }
  out.push_back('\'');
  for (const char c : sheet_name) {
    if (c == '\'') {
      out.push_back('\'');
    }
    out.push_back(c);
  }
  out.append("'!");
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

/// Splits `text` on commas, trimming ASCII whitespace around each token.
/// The caller-facing range syntax is unqualified, so no sheet name (and
/// therefore no quoted comma) can appear.
std::vector<std::string_view> split_areas(std::string_view text) {
  std::vector<std::string_view> out;
  std::size_t start = 0;
  while (start <= text.size()) {
    std::size_t comma = text.find(',', start);
    if (comma == std::string_view::npos) {
      comma = text.size();
    }
    std::string_view token = text.substr(start, comma - start);
    while (!token.empty() && (token.front() == ' ' || token.front() == '\t')) {
      token.remove_prefix(1);
    }
    while (!token.empty() && (token.back() == ' ' || token.back() == '\t')) {
      token.remove_suffix(1);
    }
    out.push_back(token);
    start = comma + 1;
  }
  return out;
}

/// Renders a resolved rectangle as an unqualified, anchor-free A1 range.
std::string format_range(const formulon::print::CellRange& range) {
  std::string out;
  formulon::a1::append_column_letters(out, range.first_col);
  out.append(std::to_string(static_cast<std::uint64_t>(range.first_row) + 1U));
  out.push_back(':');
  formulon::a1::append_column_letters(out, range.last_col);
  out.append(std::to_string(static_cast<std::uint64_t>(range.last_row) + 1U));
  return out;
}

/// True when `token` is a whole-row span (`1:2`), writing the 1-based
/// endpoints back out.
bool parse_row_span(std::string_view token, std::uint32_t* first, std::uint32_t* last) {
  const std::size_t colon = token.find(':');
  if (colon == std::string_view::npos) {
    return false;
  }
  const auto to_row = [](std::string_view part, std::uint32_t* out) {
    std::size_t pos = 0;
    if (!part.empty() && part.front() == '$') {
      part.remove_prefix(1);
    }
    return formulon::a1::parse_uint(part, &pos, out) && pos == part.size() && *out >= 1U && *out <= Sheet::kMaxRows;
  };
  return to_row(token.substr(0, colon), first) && to_row(token.substr(colon + 1), last);
}

/// True when `token` is a whole-column span (`A:A`), writing the 1-based
/// endpoints back out.
bool parse_col_span(std::string_view token, std::uint32_t* first, std::uint32_t* last) {
  const std::size_t colon = token.find(':');
  if (colon == std::string_view::npos) {
    return false;
  }
  const auto to_col = [](std::string_view part, std::uint32_t* out) {
    std::size_t pos = 0;
    if (!part.empty() && part.front() == '$') {
      part.remove_prefix(1);
    }
    return formulon::a1::parse_column_letters(part, &pos, out) && pos == part.size() && *out >= 1U &&
           *out <= Sheet::kMaxCols;
  };
  return to_col(token.substr(0, colon), first) && to_col(token.substr(colon + 1), last);
}

}  // namespace

/* -------------------------------------------------------------------------- */
/* Print area and titles                                                      */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_set_print_area(fm_workbook_t* wb, size_t sheet_index, const char* ranges_a1) {
  static constexpr const char* kFn = "fm_sheet_set_print_area";
  clear_last_error();
  if (ranges_a1 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=ranges_a1");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  const std::string_view input(ranges_a1);
  std::string formula;
  if (!input.empty()) {
    const std::string qualifier = sheet_qualifier(wb->workbook().sheet(sheet_index).name());
    const std::vector<std::string_view> tokens = split_areas(input);
    if (!formulon::c_api::parts::check_range_count(static_cast<std::uint32_t>(tokens.size()), kFn)) {
      return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
    }
    for (const std::string_view token : tokens) {
      // Round-trip the token through the resolver's own grammar so a
      // malformed area is rejected here rather than surfacing later as a
      // `kPrintInvalidArea` from `resolve_print_area`.
      formulon::print::CellRange probe;
      if (token.empty() || !formulon::print::parse_area_token(token, &probe)) {
        return set_binding_error(formulon::FormulonErrorCode::kPrintInvalidArea, kFn, "token=" + std::string(token));
      }
      if (!formula.empty()) {
        formula.push_back(',');
      }
      formula.append(qualifier);
      formula.append(absolutise_a1_token(token));
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
                                               const char** out_ranges_a1) {
  static constexpr const char* kFn = "fm_sheet_get_print_area";
  clear_last_error();
  if (out_ranges_a1 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=out_ranges_a1");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  auto areas = formulon::print::resolve_print_area(wb->workbook(), static_cast<std::uint32_t>(sheet_index));
  if (!areas) {
    return set_last_error(areas.error());
  }
  std::string rendered;
  for (const formulon::print::CellRange& range : areas.value()) {
    if (!rendered.empty()) {
      rendered.push_back(',');
    }
    rendered.append(format_range(range));
  }
  fm_workbook_t* mutable_wb = const_cast<fm_workbook_t*>(wb);
  mutable_wb->read_scratch.clear();
  mutable_wb->read_scratch.emplace_back(std::move(rendered));
  *out_ranges_a1 = mutable_wb->read_scratch.back().c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_set_print_titles(fm_workbook_t* wb, size_t sheet_index, const char* repeat_rows,
                                                 const char* repeat_cols) {
  static constexpr const char* kFn = "fm_sheet_set_print_titles";
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
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
      std::uint32_t first = 0;
      std::uint32_t last = 0;
      if (!parse_row_span(rows, &first, &last)) {
        return set_binding_error(formulon::FormulonErrorCode::kPrintInvalidArea, kFn,
                                 "repeat_rows=" + std::string(rows));
      }
      formula.append(qualifier);
      formula.append(absolutise_a1_token(rows));
    }
    if (!cols.empty()) {
      std::uint32_t first = 0;
      std::uint32_t last = 0;
      if (!parse_col_span(cols, &first, &last)) {
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
  fm_workbook_t* mutable_wb = const_cast<fm_workbook_t*>(wb);
  mutable_wb->read_scratch.clear();
  mutable_wb->read_scratch.emplace_back(std::move(rows));
  *out_repeat_rows = mutable_wb->read_scratch.back().c_str();
  mutable_wb->read_scratch.emplace_back(std::move(cols));
  *out_repeat_cols = mutable_wb->read_scratch.back().c_str();
  return 0;
}
