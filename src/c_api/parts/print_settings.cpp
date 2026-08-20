//
// C ABI - worksheet print settings (authoring side).
//
// `Sheet::print_settings()` already models everything Excel's Page Setup
// dialog writes: five raw XML fragments the writer re-emits verbatim, plus
// structured `PageSetup` / `PageMargins` / break views the paginator reads.
// Only the read side reached the ABI, so a caller could paginate a workbook
// it could not configure. This part supplies the mutators.
//
// Three surfaces, narrowing as they go:
//   - raw XML per element, for a caller that already has an Excel-authored
//     fragment to splice in;
//   - typed patch structs, which touch only the attributes the caller
//     engaged and leave every other one (including ones this engine does
//     not model) where the file put them;
//   - named helpers for the two settings that are not worksheet elements at
//     all - the print area and print titles live in workbook-scope defined
//     names, and `fitToPage` hides inside `<sheetPr>`.
//
// The invariant every entry point here upholds: the raw XML is the writer's
// source of truth, and any mutation of it re-derives the structured views
// through the same parser the OOXML reader uses
// (`io/ooxml/print_settings_parse.h`). That is what makes "set it, then
// paginate" observe the new setting instead of the loaded one.
//
// @size-budget: 17 KB

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/xml_fragment.h"
#include "io/a1_ref.h"
#include "io/ooxml/print_settings_parse.h"
#include "io/xml_utils.h"
#include "parser/reference.h"
#include "print/print_area.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/double_format.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fragment_within_size_limit;
using formulon::c_api::parts::FragmentValidation;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::validate_single_element_fragment;

namespace {

using formulon::ManualBreak;
using formulon::Orientation;
using formulon::PageMargins;
using formulon::PageSetup;
using formulon::Sheet;
using formulon::SheetPrintSettings;
using formulon::io::raw_xml;
using formulon::io::ooxml::refresh_structured_views;

// OOXML built-in defined names for the two print settings Excel stores
// outside the worksheet part.
constexpr std::string_view kPrintAreaName = "_xlnm.Print_Area";
constexpr std::string_view kPrintTitlesName = "_xlnm.Print_Titles";

// `<pageSetup scale>` is a percentage Excel accepts only inside this band;
// the UI spinner enforces the same range.
constexpr std::uint32_t kMinPrintScalePercent = 10U;
constexpr std::uint32_t kMaxPrintScalePercent = 400U;

// Perpendicular span Excel writes for a break inserted across the whole
// sheet: a row break spans every column, a column break every row.
constexpr std::uint32_t kWholeSheetColSpanMax = Sheet::kMaxCols - 1U;
constexpr std::uint32_t kWholeSheetRowSpanMax = Sheet::kMaxRows - 1U;

/// The five print-settings members reachable through the raw-XML surface,
/// named by their OOXML element so one implementation serves all of them.
struct FragmentSlot {
  std::string SheetPrintSettings::*member;
  std::string_view element;
  std::size_t limit_bytes;
};

constexpr FragmentSlot kPageSetupSlot{&SheetPrintSettings::page_setup_xml, "pageSetup",
                                      formulon::kMaxPrintFragmentBytes};
constexpr FragmentSlot kPageMarginsSlot{&SheetPrintSettings::page_margins_xml, "pageMargins",
                                        formulon::kMaxPrintFragmentBytes};
constexpr FragmentSlot kPrintOptionsSlot{&SheetPrintSettings::print_options_xml, "printOptions",
                                         formulon::kMaxPrintFragmentBytes};
constexpr FragmentSlot kHeaderFooterSlot{&SheetPrintSettings::header_footer_xml, "headerFooter",
                                         formulon::kMaxHeaderFooterFragmentBytes};
constexpr FragmentSlot kSheetPrSlot{&SheetPrintSettings::sheet_pr_xml, "sheetPr", formulon::kMaxSheetPrFragmentBytes};

/// Shared body of the five raw-XML getters.
fm_status_t get_fragment(const fm_workbook_t* wb, std::size_t sheet_index, const FragmentSlot& slot, const char* fn,
                         const char** out_xml) {
  clear_last_error();
  if (out_xml == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn, "arg=out_xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, fn); rc != 0) {
    return rc;
  }
  fm_workbook_t* mutable_wb = const_cast<fm_workbook_t*>(wb);
  mutable_wb->read_scratch.clear();
  mutable_wb->read_scratch.emplace_back(wb->workbook().sheet(sheet_index).print_settings().*slot.member);
  *out_xml = mutable_wb->read_scratch.back().c_str();
  return 0;
}

/// Shared body of the five raw-XML setters: size gate, structural check,
/// store, re-derive.
fm_status_t set_fragment(fm_workbook_t* wb, std::size_t sheet_index, const FragmentSlot& slot, const char* fn,
                         const char* xml) {
  clear_last_error();
  if (xml == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn, "arg=xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, fn); rc != 0) {
    return rc;
  }
  const std::string_view fragment(xml);
  if (!fragment.empty()) {
    // Size before structure: the parser is the part a hostile fragment
    // would exercise, so it must not see one the engine would reject
    // anyway.
    if (!fragment_within_size_limit(fragment, slot.limit_bytes)) {
      return set_binding_error(formulon::FormulonErrorCode::kPreconditionFailed, fn,
                               "element=" + std::string(slot.element) + " bytes=" + std::to_string(fragment.size()) +
                                   " limit=" + std::to_string(slot.limit_bytes));
    }
    const FragmentValidation validation = validate_single_element_fragment(fragment, slot.element);
    if (!validation.valid) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, fn,
                               "element=" + std::string(slot.element) + " " + validation.context);
    }
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  print.*slot.member = std::string(fragment);
  refresh_structured_views(slot.element, print.*slot.member, print);
  return 0;
}

/// Loads `fragment` into `doc` and returns its root, creating an empty
/// `<element_name/>` when the fragment is absent or unparseable.
///
/// Only the typed patch setters call this, and they only ever see a
/// fragment the raw setter already validated (or one the reader captured),
/// so the unparseable branch is a safety net rather than a live path.
pugi::xml_node load_or_create_root(pugi::xml_document& doc, std::string_view fragment, const char* element_name) {
  if (!fragment.empty() &&
      doc.load_buffer(fragment.data(), fragment.size(), pugi::parse_default, pugi::encoding_utf8)) {
    if (pugi::xml_node root = doc.first_child()) {
      return root;
    }
  }
  doc.reset();
  return doc.append_child(element_name);
}

/// Upserts `name="value"` on `node`.
void set_attr(pugi::xml_node node, const char* name, std::string_view value) {
  pugi::xml_attribute attr = node.attribute(name);
  if (!attr) {
    attr = node.append_attribute(name);
  }
  attr.set_value(std::string(value).c_str());
}

void set_attr_u32(pugi::xml_node node, const char* name, std::uint32_t value) {
  set_attr(node, name, std::to_string(value));
}

/// Upserts a boolean attribute in the `"true"` / `"false"` lexical form
/// Excel writes for `<printOptions>` and `<headerFooter>`.
void set_attr_bool(pugi::xml_node node, const char* name, bool value) {
  set_attr(node, name, value ? "true" : "false");
}

void set_attr_double(pugi::xml_node node, const char* name, double value) {
  std::string text;
  formulon::format_double(text, value);
  set_attr(node, name, text);
}

/// Writes a `<headerFooter>` section child, creating it in schema order or
/// removing it when `text` is empty.
///
/// ECMA-376 fixes the child order, and `append_child` would put a
/// newly-created `<oddHeader>` after an existing `<oddFooter>`. The section
/// list is walked backwards from the target to find the nearest existing
/// predecessor so a new child lands where the schema expects it.
void set_header_footer_section(pugi::xml_node root, std::size_t section_index, std::string_view text) {
  static constexpr std::string_view kSections[] = {"oddHeader",  "oddFooter",   "evenHeader",
                                                   "evenFooter", "firstHeader", "firstFooter"};
  const std::string name(kSections[section_index]);
  pugi::xml_node node = root.child(name.c_str());
  if (text.empty()) {
    if (node) {
      root.remove_child(node);
    }
    return;
  }
  if (!node) {
    pugi::xml_node after;
    for (std::size_t i = section_index; i > 0; --i) {
      const std::string previous(kSections[i - 1U]);
      if (pugi::xml_node candidate = root.child(previous.c_str())) {
        after = candidate;
        break;
      }
    }
    node = after ? root.insert_child_after(pugi::node_element, after) : root.prepend_child(pugi::node_element);
    node.set_name(name.c_str());
  }
  // pugixml escapes PCDATA on output, so the decoded text handed in here
  // reaches the file as `&amp;C...` - which is exactly the form Excel
  // writes for a header formatting code.
  node.text().set(std::string(text).c_str());
}

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
    return formulon::io::parse_uint(part, &pos, out) && pos == part.size() && *out >= 1U && *out <= Sheet::kMaxRows;
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
    return formulon::io::parse_column_letters(part, &pos, out) && pos == part.size() && *out >= 1U &&
           *out <= Sheet::kMaxCols;
  };
  return to_col(token.substr(0, colon), first) && to_col(token.substr(colon + 1), last);
}

/// Upserts a manual break into an axis vector, keeping it sorted by index.
fm_status_t upsert_break(std::vector<ManualBreak>& breaks, std::uint32_t index, std::uint32_t span_max, bool manual,
                         const char* fn) {
  const auto pos = std::lower_bound(breaks.begin(), breaks.end(), index,
                                    [](const ManualBreak& entry, std::uint32_t key) { return entry.id < key; });
  if (pos != breaks.end() && pos->id == index) {
    pos->min = 0U;
    pos->max = span_max;
    pos->manual = manual;
    return 0;
  }
  if (breaks.size() >= formulon::kMaxManualBreaksPerAxis) {
    return set_binding_error(
        formulon::FormulonErrorCode::kPreconditionFailed, fn,
        "breaks=" + std::to_string(breaks.size()) + " limit=" + std::to_string(formulon::kMaxManualBreaksPerAxis));
  }
  ManualBreak entry;
  entry.id = index;
  entry.min = 0U;
  entry.max = span_max;
  entry.manual = manual;
  breaks.insert(pos, entry);
  return 0;
}

void erase_break(std::vector<ManualBreak>& breaks, std::uint32_t index) {
  const auto pos = std::lower_bound(breaks.begin(), breaks.end(), index,
                                    [](const ManualBreak& entry, std::uint32_t key) { return entry.id < key; });
  if (pos != breaks.end() && pos->id == index) {
    breaks.erase(pos);
  }
}

// Counts one axis. The `size_t` return has no room for a status, so an
// argument error is reported as `0` plus a populated thread-local
// diagnostic; `clear_last_error` above is what makes the empty message a
// reliable "this zero means the axis is empty" signal. Callers that
// enumerate by count-then-`break_at` must consult the diagnostic, since
// the `break_at` loop that would otherwise surface the status never runs
// on a zero count.
std::size_t break_count(const fm_workbook_t* wb, std::size_t sheet_index, bool rows, const char* fn) {
  clear_last_error();
  if (check_sheet_index(wb, sheet_index, fn) != 0) {
    return 0;
  }
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  return rows ? print.manual_row_breaks.size() : print.manual_col_breaks.size();
}

fm_status_t break_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index, bool rows, const char* fn,
                     fm_page_break_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn, "arg=out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, fn); rc != 0) {
    return rc;
  }
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  const std::vector<ManualBreak>& breaks = rows ? print.manual_row_breaks : print.manual_col_breaks;
  if (index >= breaks.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, fn,
                             "index=" + std::to_string(index) + " count=" + std::to_string(breaks.size()));
  }
  const ManualBreak& entry = breaks[index];
  out->id = entry.id;
  out->min = entry.min;
  out->max = entry.max;
  out->manual = entry.manual ? 1 : 0;
  return 0;
}

}  // namespace

/* -------------------------------------------------------------------------- */
/* Raw XML                                                                    */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_get_page_setup_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml) {
  return get_fragment(wb, sheet_index, kPageSetupSlot, "fm_sheet_get_page_setup_xml", out_xml);
}

extern "C" fm_status_t fm_sheet_set_page_setup_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  static constexpr const char* kFn = "fm_sheet_set_page_setup_xml";
  clear_last_error();
  if (xml == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  // A `<pageSetup r:id>` names the sheet's printerSettings part. Storing
  // one on a sheet that has no such part leaves a relationship reference
  // the writer cannot resolve, and Excel offers to repair the file. Reject
  // rather than strip: a fragment lifted off another sheet is a caller
  // mistake worth surfacing, and silently rewriting it would hide it.
  const std::string_view fragment(xml);
  if (fragment.find("r:id") != std::string_view::npos &&
      wb->workbook().sheet(sheet_index).print_settings().printer_settings_path.empty()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, kFn, "printer_settings_rid_without_part");
  }
  return set_fragment(wb, sheet_index, kPageSetupSlot, kFn, xml);
}

extern "C" fm_status_t fm_sheet_get_page_margins_xml(const fm_workbook_t* wb, size_t sheet_index,
                                                     const char** out_xml) {
  return get_fragment(wb, sheet_index, kPageMarginsSlot, "fm_sheet_get_page_margins_xml", out_xml);
}

extern "C" fm_status_t fm_sheet_set_page_margins_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  return set_fragment(wb, sheet_index, kPageMarginsSlot, "fm_sheet_set_page_margins_xml", xml);
}

extern "C" fm_status_t fm_sheet_get_print_options_xml(const fm_workbook_t* wb, size_t sheet_index,
                                                      const char** out_xml) {
  return get_fragment(wb, sheet_index, kPrintOptionsSlot, "fm_sheet_get_print_options_xml", out_xml);
}

extern "C" fm_status_t fm_sheet_set_print_options_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  return set_fragment(wb, sheet_index, kPrintOptionsSlot, "fm_sheet_set_print_options_xml", xml);
}

extern "C" fm_status_t fm_sheet_get_header_footer_xml(const fm_workbook_t* wb, size_t sheet_index,
                                                      const char** out_xml) {
  return get_fragment(wb, sheet_index, kHeaderFooterSlot, "fm_sheet_get_header_footer_xml", out_xml);
}

extern "C" fm_status_t fm_sheet_set_header_footer_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  return set_fragment(wb, sheet_index, kHeaderFooterSlot, "fm_sheet_set_header_footer_xml", xml);
}

extern "C" fm_status_t fm_sheet_get_sheet_pr_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml) {
  return get_fragment(wb, sheet_index, kSheetPrSlot, "fm_sheet_get_sheet_pr_xml", out_xml);
}

extern "C" fm_status_t fm_sheet_set_sheet_pr_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  return set_fragment(wb, sheet_index, kSheetPrSlot, "fm_sheet_set_sheet_pr_xml", xml);
}

extern "C" fm_status_t fm_sheet_set_fit_to_page(fm_workbook_t* wb, size_t sheet_index, int32_t enabled) {
  static constexpr const char* kFn = "fm_sheet_set_fit_to_page";
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.sheet_pr_xml, "sheetPr");
  pugi::xml_node page_setup_pr = root.child("pageSetUpPr");
  if (enabled == 0) {
    // Clear the attribute but keep both elements: `<pageSetUpPr>` also
    // carries `autoPageBreaks`, and `<sheetPr>` carries tab colour and the
    // VBA code name.
    if (page_setup_pr) {
      page_setup_pr.remove_attribute("fitToPage");
    }
  } else {
    if (!page_setup_pr) {
      // `<pageSetUpPr>` is the last child in `<sheetPr>`'s content model,
      // after `<tabColor>` and `<outlinePr>`.
      page_setup_pr = root.append_child("pageSetUpPr");
    }
    set_attr_bool(page_setup_pr, "fitToPage", true);
  }
  print.sheet_pr_xml = raw_xml(root);
  print.page_setup.fit_to_page = enabled != 0;
  return 0;
}

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
      return static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
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

/* -------------------------------------------------------------------------- */
/* Manual page breaks                                                         */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_add_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t manual) {
  static constexpr const char* kFn = "fm_sheet_add_row_break";
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  if (row >= Sheet::kMaxRows) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, kFn, "row=" + std::to_string(row));
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  return upsert_break(print.manual_row_breaks, row, kWholeSheetColSpanMax, manual != 0, kFn);
}

extern "C" fm_status_t fm_sheet_add_col_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col, int32_t manual) {
  static constexpr const char* kFn = "fm_sheet_add_col_break";
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  if (col >= Sheet::kMaxCols) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, kFn, "col=" + std::to_string(col));
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  return upsert_break(print.manual_col_breaks, col, kWholeSheetRowSpanMax, manual != 0, kFn);
}

extern "C" fm_status_t fm_sheet_remove_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_remove_row_break"); rc != 0) {
    return rc;
  }
  erase_break(wb->workbook().sheet(sheet_index).mutable_print_settings().manual_row_breaks, row);
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_col_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_remove_col_break"); rc != 0) {
    return rc;
  }
  erase_break(wb->workbook().sheet(sheet_index).mutable_print_settings().manual_col_breaks, col);
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_breaks(fm_workbook_t* wb, size_t sheet_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_clear_breaks"); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  print.manual_row_breaks.clear();
  print.manual_col_breaks.clear();
  return 0;
}

extern "C" size_t fm_sheet_row_break_count(const fm_workbook_t* wb, size_t sheet_index) {
  return break_count(wb, sheet_index, /*rows=*/true, "fm_sheet_row_break_count");
}

extern "C" fm_status_t fm_sheet_row_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                             fm_page_break_t* out) {
  return break_at(wb, sheet_index, index, /*rows=*/true, "fm_sheet_row_break_at", out);
}

extern "C" size_t fm_sheet_col_break_count(const fm_workbook_t* wb, size_t sheet_index) {
  return break_count(wb, sheet_index, /*rows=*/false, "fm_sheet_col_break_count");
}

extern "C" fm_status_t fm_sheet_col_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                             fm_page_break_t* out) {
  return break_at(wb, sheet_index, index, /*rows=*/false, "fm_sheet_col_break_at", out);
}

/* -------------------------------------------------------------------------- */
/* Typed patch setters                                                        */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_set_page_setup(fm_workbook_t* wb, size_t sheet_index, const fm_page_setup_t* setup) {
  static constexpr const char* kFn = "fm_sheet_set_page_setup";
  clear_last_error();
  if (setup == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=setup");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  if (setup->orientation_engaged != 0 && setup->orientation > FM_ORIENTATION_LANDSCAPE) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, kFn,
                             "orientation=" + std::to_string(setup->orientation));
  }
  if (setup->scale_engaged != 0 && (setup->scale < kMinPrintScalePercent || setup->scale > kMaxPrintScalePercent)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, kFn,
                             "scale=" + std::to_string(setup->scale));
  }

  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.page_setup_xml, "pageSetup");
  if (setup->orientation_engaged != 0) {
    if (setup->orientation == FM_ORIENTATION_DEFAULT) {
      root.remove_attribute("orientation");
    } else {
      set_attr(root, "orientation", setup->orientation == FM_ORIENTATION_PORTRAIT ? "portrait" : "landscape");
    }
  }
  if (setup->paper_size_engaged != 0) {
    set_attr_u32(root, "paperSize", setup->paper_size);
  }
  if (setup->scale_engaged != 0) {
    set_attr_u32(root, "scale", setup->scale);
  }
  if (setup->fit_to_width_engaged != 0) {
    set_attr_u32(root, "fitToWidth", setup->fit_to_width);
  }
  if (setup->fit_to_height_engaged != 0) {
    set_attr_u32(root, "fitToHeight", setup->fit_to_height);
  }
  print.page_setup_xml = raw_xml(root);
  refresh_structured_views("pageSetup", print.page_setup_xml, print);

  // `fitToPage` is not a `<pageSetup>` attribute; it rides along on this
  // struct because callers think of it as part of one page-setup decision.
  if (setup->fit_to_page_engaged != 0) {
    return fm_sheet_set_fit_to_page(wb, sheet_index, setup->fit_to_page);
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_page_margins(fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_page_margins_t* margins) {
  static constexpr const char* kFn = "fm_sheet_set_page_margins";
  clear_last_error();
  if (margins == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=margins");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  struct MarginField {
    const char* attr;
    int32_t engaged;
    double value;
  };
  // Only the engaged fields are stored. ECMA-376 requires all six on the
  // saved element, but completing it here would report every margin as
  // engaged on the next read and erase the caller's ability to tell a
  // stated margin from a defaulted one. The OOXML writer fills the gap
  // when it emits the element instead.
  const MarginField fields[] = {
      {"left", margins->left_engaged, margins->left},       {"right", margins->right_engaged, margins->right},
      {"top", margins->top_engaged, margins->top},          {"bottom", margins->bottom_engaged, margins->bottom},
      {"header", margins->header_engaged, margins->header}, {"footer", margins->footer_engaged, margins->footer},
  };
  for (const MarginField& field : fields) {
    if (field.engaged == 0) {
      continue;
    }
    // The paginator subtracts these from the paper; a negative, infinite or
    // NaN one produces a printable body that is not a rectangle.
    if (!std::isfinite(field.value) || field.value < 0.0) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, kFn,
                               std::string(field.attr) + "=invalid");
    }
  }

  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.page_margins_xml, "pageMargins");
  for (const MarginField& field : fields) {
    if (field.engaged != 0) {
      set_attr_double(root, field.attr, field.value);
    }
  }
  print.page_margins_xml = raw_xml(root);
  refresh_structured_views("pageMargins", print.page_margins_xml, print);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_print_options(fm_workbook_t* wb, size_t sheet_index,
                                                  const fm_print_options_t* options) {
  static constexpr const char* kFn = "fm_sheet_set_print_options";
  clear_last_error();
  if (options == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=options");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.print_options_xml, "printOptions");
  if (options->grid_lines_engaged != 0) {
    set_attr_bool(root, "gridLines", options->grid_lines != 0);
  }
  if (options->headings_engaged != 0) {
    set_attr_bool(root, "headings", options->headings != 0);
  }
  if (options->horizontal_centered_engaged != 0) {
    set_attr_bool(root, "horizontalCentered", options->horizontal_centered != 0);
  }
  if (options->vertical_centered_engaged != 0) {
    set_attr_bool(root, "verticalCentered", options->vertical_centered != 0);
  }
  print.print_options_xml = raw_xml(root);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_header_footer(fm_workbook_t* wb, size_t sheet_index, const fm_header_footer_t* hf) {
  static constexpr const char* kFn = "fm_sheet_set_header_footer";
  clear_last_error();
  if (hf == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=hf");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.header_footer_xml, "headerFooter");
  if (hf->different_odd_even_engaged != 0) {
    set_attr_bool(root, "differentOddEven", hf->different_odd_even != 0);
  }
  if (hf->different_first_engaged != 0) {
    set_attr_bool(root, "differentFirst", hf->different_first != 0);
  }
  if (hf->scale_with_doc_engaged != 0) {
    set_attr_bool(root, "scaleWithDoc", hf->scale_with_doc != 0);
  }
  if (hf->align_with_margins_engaged != 0) {
    set_attr_bool(root, "alignWithMargins", hf->align_with_margins != 0);
  }
  const char* const sections[] = {hf->odd_header,  hf->odd_footer,   hf->even_header,
                                  hf->even_footer, hf->first_header, hf->first_footer};
  for (std::size_t i = 0; i < std::size(sections); ++i) {
    if (sections[i] == nullptr) {
      continue;
    }
    set_header_footer_section(root, i, sections[i]);
  }
  print.header_footer_xml = raw_xml(root);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_page_setup(const fm_workbook_t* wb, size_t sheet_index, fm_page_setup_t* out) {
  static constexpr const char* kFn = "fm_sheet_get_page_setup";
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  const PageSetup& view = print.page_setup;
  *out = fm_page_setup_t{};
  switch (view.orientation) {
    case Orientation::kPortrait:
      out->orientation = FM_ORIENTATION_PORTRAIT;
      break;
    case Orientation::kLandscape:
      out->orientation = FM_ORIENTATION_LANDSCAPE;
      break;
    case Orientation::kDefault:
      out->orientation = FM_ORIENTATION_DEFAULT;
      break;
  }
  out->paper_size = view.paper_size;
  out->scale = view.scale;
  out->fit_to_width = view.fit_to_width;
  out->fit_to_height = view.fit_to_height;
  out->fit_to_page = view.fit_to_page ? 1 : 0;

  // The `_engaged` flags report presence, which only the XML knows: the
  // structured view cannot distinguish `scale="100"` from an absent
  // attribute that defaults to 100.
  pugi::xml_document doc;
  if (!print.page_setup_xml.empty() && doc.load_buffer(print.page_setup_xml.data(), print.page_setup_xml.size(),
                                                       pugi::parse_default, pugi::encoding_utf8)) {
    const pugi::xml_node root = doc.first_child();
    out->orientation_engaged = root.attribute("orientation") ? 1 : 0;
    out->paper_size_engaged = root.attribute("paperSize") ? 1 : 0;
    out->scale_engaged = root.attribute("scale") ? 1 : 0;
    out->fit_to_width_engaged = root.attribute("fitToWidth") ? 1 : 0;
    out->fit_to_height_engaged = root.attribute("fitToHeight") ? 1 : 0;
  }
  pugi::xml_document sheet_pr_doc;
  if (!print.sheet_pr_xml.empty() && sheet_pr_doc.load_buffer(print.sheet_pr_xml.data(), print.sheet_pr_xml.size(),
                                                              pugi::parse_default, pugi::encoding_utf8)) {
    const pugi::xml_node page_setup_pr = sheet_pr_doc.first_child().child("pageSetUpPr");
    out->fit_to_page_engaged = (page_setup_pr && page_setup_pr.attribute("fitToPage")) ? 1 : 0;
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_page_margins(const fm_workbook_t* wb, size_t sheet_index, fm_page_margins_t* out) {
  static constexpr const char* kFn = "fm_sheet_get_page_margins";
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn, "arg=out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  const PageMargins& view = print.page_margins;
  *out = fm_page_margins_t{};
  out->left = view.left;
  out->right = view.right;
  out->top = view.top;
  out->bottom = view.bottom;
  out->header = view.header;
  out->footer = view.footer;

  pugi::xml_document doc;
  if (!print.page_margins_xml.empty() && doc.load_buffer(print.page_margins_xml.data(), print.page_margins_xml.size(),
                                                         pugi::parse_default, pugi::encoding_utf8)) {
    const pugi::xml_node root = doc.first_child();
    out->left_engaged = root.attribute("left") ? 1 : 0;
    out->right_engaged = root.attribute("right") ? 1 : 0;
    out->top_engaged = root.attribute("top") ? 1 : 0;
    out->bottom_engaged = root.attribute("bottom") ? 1 : 0;
    out->header_engaged = root.attribute("header") ? 1 : 0;
    out->footer_engaged = root.attribute("footer") ? 1 : 0;
  }
  return 0;
}
