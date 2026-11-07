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
// @size-budget: 12 KB

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "c_api/parts/xml_fragment.h"
#include "io/ooxml/print_settings_parse.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/double_format.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "workbook.h"

using formulon::c_api::parts::check_column_index;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_row_index;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fragment_within_size_limit;
using formulon::c_api::parts::FragmentValidation;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;
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
fm_status_t get_fragment(const fm_workbook_t* wb, std::size_t sheet_index, const FragmentSlot& slot,
                         const char** out_xml) {
  clear_last_error();
  if (out_xml == nullptr) {
    return null_argument("out_xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_xml = publish_scratch(wb, wb->workbook().sheet(sheet_index).print_settings().*slot.member);
  return 0;
}

/// Shared body of the five raw-XML setters: size gate, structural check,
/// store, re-derive.
fm_status_t set_fragment(fm_workbook_t* wb, std::size_t sheet_index, const FragmentSlot& slot, const char* xml) {
  clear_last_error();
  if (xml == nullptr) {
    return null_argument("xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const std::string_view fragment(xml);
  if (!fragment.empty()) {
    // Size before structure: the parser is the part a hostile fragment
    // would exercise, so it must not see one the engine would reject
    // anyway.
    if (!fragment_within_size_limit(fragment, slot.limit_bytes)) {
      return set_binding_error(formulon::FormulonErrorCode::kPreconditionFailed, "XML fragment exceeds the size limit",
                               "element=" + std::string(slot.element) + " bytes=" + std::to_string(fragment.size()) +
                                   " limit=" + std::to_string(slot.limit_bytes));
    }
    if (auto rc = check_utf8(fragment, "element", slot.element); rc != 0) {
      return rc;
    }
    const FragmentValidation validation = validate_single_element_fragment(fragment, slot.element);
    if (!validation.valid) {
      return invalid_argument("invalid XML fragment",
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

/// One `(attribute, engaged, value)` triple of a partial boolean update.
struct BoolAttr {
  const char* attr;
  std::int32_t engaged;
  std::int32_t value;
};

/// Applies every engaged entry of `attrs` through `set_attr_bool`.
void set_engaged_bools(pugi::xml_node node, const BoolAttr* attrs, std::size_t count) {
  for (std::size_t i = 0; i < count; ++i) {
    if (attrs[i].engaged != 0) {
      set_attr_bool(node, attrs[i].attr, attrs[i].value != 0);
    }
  }
}

/// One `(attribute, engaged flag)` pair filled from attribute presence.
struct EngagedAttr {
  const char* attr;
  std::int32_t* engaged;
};

/// Sets each flag to whether `node` carries its attribute.
void read_engaged(pugi::xml_node node, const EngagedAttr* attrs, std::size_t count) {
  for (std::size_t i = 0; i < count; ++i) {
    *attrs[i].engaged = node.attribute(attrs[i].attr) ? 1 : 0;
  }
}

void set_attr_double(pugi::xml_node node, const char* name, double value) {
  std::string text;
  formulon::format_double(text, value);
  set_attr(node, name, text);
}

/// `<headerFooter>` section children in schema order, matching the section
/// pointers of `fm_header_footer_t`.
constexpr std::string_view kHeaderFooterSections[] = {"oddHeader",  "oddFooter",   "evenHeader",
                                                      "evenFooter", "firstHeader", "firstFooter"};

/// Writes a `<headerFooter>` section child, creating it in schema order or
/// removing it when `text` is empty.
///
/// ECMA-376 fixes the child order, and `append_child` would put a
/// newly-created `<oddHeader>` after an existing `<oddFooter>`. The section
/// list is walked backwards from the target to find the nearest existing
/// predecessor so a new child lands where the schema expects it.
void set_header_footer_section(pugi::xml_node root, std::size_t section_index, std::string_view text) {
  const std::string name(kHeaderFooterSections[section_index]);
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
      const std::string previous(kHeaderFooterSections[i - 1U]);
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

/// Sets or clears `<sheetPr><pageSetUpPr fitToPage>` without disturbing any
/// other `<sheetPr>` content.
void set_fit_to_page(SheetPrintSettings& print, bool enabled) {
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.sheet_pr_xml, "sheetPr");
  pugi::xml_node page_setup_pr = root.child("pageSetUpPr");
  if (!enabled) {
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
  print.page_setup.fit_to_page = enabled;
}

/// Sets or removes `<sheetPr><tabColor>`, leaving the other `<sheetPr>`
/// content alone. `<tabColor>` leads `<sheetPr>`'s content model.
void set_tab_color(SheetPrintSettings& print, const formulon::ColorSpec& color) {
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.sheet_pr_xml, "sheetPr");
  root.remove_child("tabColor");
  if (color.kind != formulon::ColorSpec::Kind::kNone) {
    pugi::xml_node node = root.prepend_child("tabColor");
    switch (color.kind) {
      case formulon::ColorSpec::Kind::kRgb: {
        // snprintf is unavailable to engine code on WASM.
        std::string hex(8U, '0');
        for (std::size_t i = 0; i < 8U; ++i) {
          hex[7U - i] = "0123456789ABCDEF"[(color.rgb >> (4U * i)) & 0xFU];
        }
        set_attr(node, "rgb", hex);
        break;
      }
      case formulon::ColorSpec::Kind::kTheme:
        set_attr_u32(node, "theme", color.theme);
        if (color.tint != 0.0) {
          set_attr_double(node, "tint", color.tint);
        }
        break;
      case formulon::ColorSpec::Kind::kIndexed:
        set_attr_u32(node, "indexed", color.indexed);
        break;
      default:
        set_attr(node, "auto", "1");
        break;
    }
  }
  print.sheet_pr_xml = (root.first_child() || root.first_attribute()) ? raw_xml(root) : std::string();
}

/// Loads a stored print fragment for a getter; false when it is absent.
bool load_fragment(pugi::xml_document& doc, const std::string& fragment) {
  return !fragment.empty() &&
         doc.load_buffer(fragment.data(), fragment.size(), pugi::parse_default, pugi::encoding_utf8);
}

/// Upserts a manual break into an axis vector, keeping it sorted by index.
fm_status_t upsert_break(std::vector<ManualBreak>& breaks, std::uint32_t index, std::uint32_t span_max, bool manual) {
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
        formulon::FormulonErrorCode::kPreconditionFailed, "too many manual page breaks",
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

/// Erases the break at `index`; NOT_FOUND when the axis has none there.
fm_status_t erase_break(std::vector<ManualBreak>& breaks, std::uint32_t index, const char* what) {
  const auto pos = std::lower_bound(breaks.begin(), breaks.end(), index,
                                    [](const ManualBreak& entry, std::uint32_t key) { return entry.id < key; });
  if (pos == breaks.end() || pos->id != index) {
    return not_found(what, "index=" + std::to_string(index));
  }
  breaks.erase(pos);
  return 0;
}

/// Shared body of `fm_sheet_add_row_break` / `fm_sheet_add_column_break`.
fm_status_t add_break(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t index, bool manual, bool rows) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = rows ? check_row_index(index) : check_column_index(index); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  return rows ? upsert_break(print.manual_row_breaks, index, kWholeSheetColSpanMax, manual)
              : upsert_break(print.manual_col_breaks, index, kWholeSheetRowSpanMax, manual);
}

// Counts one axis. The `size_t` return has no room for a status, so an
// argument error is reported as `0` plus a populated thread-local
// diagnostic; `clear_last_error` above is what makes the empty message a
// reliable "this zero means the axis is empty" signal. Callers that
// enumerate by count-then-`break_at` must consult the diagnostic, since
// the `break_at` loop that would otherwise surface the status never runs
// on a zero count.
fm_status_t break_count(const fm_workbook_t* wb, std::size_t sheet_index, bool rows, std::size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  *out_count = rows ? print.manual_row_breaks.size() : print.manual_col_breaks.size();
  return 0;
}

fm_status_t break_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index, bool rows,
                     fm_page_break_t* out) {
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
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  const std::vector<ManualBreak>& breaks = rows ? print.manual_row_breaks : print.manual_col_breaks;
  if (auto rc = check_index(index, breaks.size(), "index"); rc != 0) {
    return rc;
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
  return get_fragment(wb, sheet_index, kPageSetupSlot, out_xml);
}

extern "C" fm_status_t fm_sheet_set_page_setup_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  clear_last_error();
  if (xml == nullptr) {
    return null_argument("xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
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
    return invalid_argument("printer settings relationship has no part");
  }
  return set_fragment(wb, sheet_index, kPageSetupSlot, xml);
}

extern "C" fm_status_t fm_sheet_get_page_margins_xml(const fm_workbook_t* wb, size_t sheet_index,
                                                     const char** out_xml) {
  return get_fragment(wb, sheet_index, kPageMarginsSlot, out_xml);
}

extern "C" fm_status_t fm_sheet_set_page_margins_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  return set_fragment(wb, sheet_index, kPageMarginsSlot, xml);
}

extern "C" fm_status_t fm_sheet_get_print_options_xml(const fm_workbook_t* wb, size_t sheet_index,
                                                      const char** out_xml) {
  return get_fragment(wb, sheet_index, kPrintOptionsSlot, out_xml);
}

extern "C" fm_status_t fm_sheet_set_print_options_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  return set_fragment(wb, sheet_index, kPrintOptionsSlot, xml);
}

extern "C" fm_status_t fm_sheet_get_header_footer_xml(const fm_workbook_t* wb, size_t sheet_index,
                                                      const char** out_xml) {
  return get_fragment(wb, sheet_index, kHeaderFooterSlot, out_xml);
}

extern "C" fm_status_t fm_sheet_set_header_footer_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  return set_fragment(wb, sheet_index, kHeaderFooterSlot, xml);
}

extern "C" fm_status_t fm_sheet_get_sheet_pr_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml) {
  return get_fragment(wb, sheet_index, kSheetPrSlot, out_xml);
}

extern "C" fm_status_t fm_sheet_set_sheet_pr_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  return set_fragment(wb, sheet_index, kSheetPrSlot, xml);
}

/* -------------------------------------------------------------------------- */
/* Manual page breaks                                                         */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_add_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t manual) {
  clear_last_error();
  return add_break(wb, sheet_index, row, manual != 0, /*rows=*/true);
}

extern "C" fm_status_t fm_sheet_add_column_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col, int32_t manual) {
  clear_last_error();
  return add_break(wb, sheet_index, col, manual != 0, /*rows=*/false);
}

extern "C" fm_status_t fm_sheet_remove_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_row_index(row); rc != 0) {
    return rc;
  }
  return erase_break(wb->workbook().sheet(sheet_index).mutable_print_settings().manual_row_breaks, row,
                     "no row break at that row");
}

extern "C" fm_status_t fm_sheet_remove_column_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_column_index(col); rc != 0) {
    return rc;
  }
  return erase_break(wb->workbook().sheet(sheet_index).mutable_print_settings().manual_col_breaks, col,
                     "no column break at that column");
}

extern "C" fm_status_t fm_sheet_clear_breaks(fm_workbook_t* wb, size_t sheet_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  print.manual_row_breaks.clear();
  print.manual_col_breaks.clear();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  return break_count(wb, sheet_index, /*rows=*/true, out_count);
}

extern "C" fm_status_t fm_sheet_get_row_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                 fm_page_break_t* out) {
  return break_at(wb, sheet_index, index, /*rows=*/true, out);
}

extern "C" fm_status_t fm_sheet_get_column_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  return break_count(wb, sheet_index, /*rows=*/false, out_count);
}

extern "C" fm_status_t fm_sheet_get_column_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                    fm_page_break_t* out) {
  return break_at(wb, sheet_index, index, /*rows=*/false, out);
}

/* -------------------------------------------------------------------------- */
/* Typed patch setters                                                        */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_set_page_setup(fm_workbook_t* wb, size_t sheet_index, const fm_page_setup_t* setup) {
  clear_last_error();
  if (setup == nullptr) {
    return null_argument("setup");
  }
  if (auto rc = check_record_size(*setup); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (setup->orientation_engaged != 0) {
    if (auto rc = check_enum_domain(setup->orientation, 0, FM_PAGE_ORIENTATION_LANDSCAPE, "orientation"); rc != 0) {
      return rc;
    }
  }
  if (setup->scale_engaged != 0 && (setup->scale < kMinPrintScalePercent || setup->scale > kMaxPrintScalePercent)) {
    return invalid_argument("scale is out of range", "scale=" + std::to_string(setup->scale));
  }

  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.page_setup_xml, "pageSetup");
  if (setup->orientation_engaged != 0) {
    set_attr(root, "orientation", setup->orientation == FM_PAGE_ORIENTATION_LANDSCAPE ? "landscape" : "portrait");
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
    set_fit_to_page(print, setup->fit_to_page != 0);
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_page_margins(fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_page_margins_t* margins) {
  clear_last_error();
  if (margins == nullptr) {
    return null_argument("margins");
  }
  if (auto rc = check_record_size(*margins); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
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
      return invalid_argument("margin must be finite and non-negative", std::string(field.attr) + "=invalid");
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
  clear_last_error();
  if (options == nullptr) {
    return null_argument("options");
  }
  if (auto rc = check_record_size(*options); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.print_options_xml, "printOptions");
  const BoolAttr attrs[] = {
      {"gridLines", options->grid_lines_engaged, options->grid_lines},
      {"headings", options->headings_engaged, options->headings},
      {"horizontalCentered", options->horizontal_centered_engaged, options->horizontal_centered},
      {"verticalCentered", options->vertical_centered_engaged, options->vertical_centered},
  };
  set_engaged_bools(root, attrs, std::size(attrs));
  print.print_options_xml = raw_xml(root);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_header_footer(fm_workbook_t* wb, size_t sheet_index, const fm_header_footer_t* hf) {
  clear_last_error();
  if (hf == nullptr) {
    return null_argument("hf");
  }
  if (auto rc = check_record_size(*hf); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8(
          {hf->odd_header, hf->odd_footer, hf->even_header, hf->even_footer, hf->first_header, hf->first_footer});
      rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.header_footer_xml, "headerFooter");
  const BoolAttr attrs[] = {
      {"differentOddEven", hf->different_odd_even_engaged, hf->different_odd_even},
      {"differentFirst", hf->different_first_engaged, hf->different_first},
      {"scaleWithDoc", hf->scale_with_doc_engaged, hf->scale_with_doc},
      {"alignWithMargins", hf->align_with_margins_engaged, hf->align_with_margins},
  };
  set_engaged_bools(root, attrs, std::size(attrs));
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
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  const PageSetup& view = print.page_setup;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_page_setup_t{};
  out->struct_size = struct_size;
  // `default` and an absent attribute both print portrait.
  out->orientation =
      view.orientation == Orientation::kLandscape ? FM_PAGE_ORIENTATION_LANDSCAPE : FM_PAGE_ORIENTATION_PORTRAIT;
  out->paper_size = view.paper_size;
  out->scale = view.scale;
  out->fit_to_width = view.fit_to_width;
  out->fit_to_height = view.fit_to_height;
  out->fit_to_page = view.fit_to_page ? 1 : 0;

  // The `_engaged` flags report presence, which only the XML knows: the
  // structured view cannot distinguish `scale="100"` from an absent
  // attribute that defaults to 100.
  pugi::xml_document doc;
  if (load_fragment(doc, print.page_setup_xml)) {
    const pugi::xml_node root = doc.first_child();
    const EngagedAttr attrs[] = {
        {"orientation", &out->orientation_engaged},
        {"paperSize", &out->paper_size_engaged},
        {"scale", &out->scale_engaged},
        {"fitToWidth", &out->fit_to_width_engaged},
        {"fitToHeight", &out->fit_to_height_engaged},
    };
    read_engaged(root, attrs, std::size(attrs));
  }
  pugi::xml_document sheet_pr_doc;
  if (load_fragment(sheet_pr_doc, print.sheet_pr_xml)) {
    const pugi::xml_node page_setup_pr = sheet_pr_doc.first_child().child("pageSetUpPr");
    out->fit_to_page_engaged = (page_setup_pr && page_setup_pr.attribute("fitToPage")) ? 1 : 0;
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_page_margins(const fm_workbook_t* wb, size_t sheet_index, fm_page_margins_t* out) {
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
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  const PageMargins& view = print.page_margins;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_page_margins_t{};
  out->struct_size = struct_size;
  out->left = view.left;
  out->right = view.right;
  out->top = view.top;
  out->bottom = view.bottom;
  out->header = view.header;
  out->footer = view.footer;

  pugi::xml_document doc;
  if (load_fragment(doc, print.page_margins_xml)) {
    const pugi::xml_node root = doc.first_child();
    const EngagedAttr attrs[] = {
        {"left", &out->left_engaged},     {"right", &out->right_engaged},   {"top", &out->top_engaged},
        {"bottom", &out->bottom_engaged}, {"header", &out->header_engaged}, {"footer", &out->footer_engaged},
    };
    read_engaged(root, attrs, std::size(attrs));
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_print_options(const fm_workbook_t* wb, size_t sheet_index,
                                                  fm_print_options_t* out) {
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
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_print_options_t{};
  out->struct_size = struct_size;
  pugi::xml_document doc;
  if (!load_fragment(doc, wb->workbook().sheet(sheet_index).print_settings().print_options_xml)) {
    return 0;
  }
  const pugi::xml_node root = doc.first_child();
  out->grid_lines = root.attribute("gridLines").as_bool(false) ? 1 : 0;
  out->headings = root.attribute("headings").as_bool(false) ? 1 : 0;
  out->horizontal_centered = root.attribute("horizontalCentered").as_bool(false) ? 1 : 0;
  out->vertical_centered = root.attribute("verticalCentered").as_bool(false) ? 1 : 0;
  const EngagedAttr attrs[] = {
      {"gridLines", &out->grid_lines_engaged},
      {"headings", &out->headings_engaged},
      {"horizontalCentered", &out->horizontal_centered_engaged},
      {"verticalCentered", &out->vertical_centered_engaged},
  };
  read_engaged(root, attrs, std::size(attrs));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_header_footer(const fm_workbook_t* wb, size_t sheet_index,
                                                  fm_header_footer_t* out) {
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
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_header_footer_t{};
  out->struct_size = struct_size;
  // CT_HeaderFooter defaults: scaleWithDoc and alignWithMargins are true.
  out->scale_with_doc = 1;
  out->align_with_margins = 1;
  pugi::xml_document doc;
  if (!load_fragment(doc, wb->workbook().sheet(sheet_index).print_settings().header_footer_xml)) {
    return 0;
  }
  const pugi::xml_node root = doc.first_child();
  out->different_odd_even = root.attribute("differentOddEven").as_bool(false) ? 1 : 0;
  out->different_first = root.attribute("differentFirst").as_bool(false) ? 1 : 0;
  out->scale_with_doc = root.attribute("scaleWithDoc").as_bool(true) ? 1 : 0;
  out->align_with_margins = root.attribute("alignWithMargins").as_bool(true) ? 1 : 0;
  const EngagedAttr attrs[] = {
      {"differentOddEven", &out->different_odd_even_engaged},
      {"differentFirst", &out->different_first_engaged},
      {"scaleWithDoc", &out->scale_with_doc_engaged},
      {"alignWithMargins", &out->align_with_margins_engaged},
  };
  read_engaged(root, attrs, std::size(attrs));
  TextStore& store = fresh_read_scratch(wb);
  const char** const sections[] = {&out->odd_header,  &out->odd_footer,   &out->even_header,
                                   &out->even_footer, &out->first_header, &out->first_footer};
  for (std::size_t i = 0; i < std::size(sections); ++i) {
    const pugi::xml_node node = root.child(std::string(kHeaderFooterSections[i]).c_str());
    if (node) {
      *sections[i] = store_cstr(store, node.text().get());
    }
  }
  return 0;
}

/* -------------------------------------------------------------------------- */
/* Sheet tab colour                                                           */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_get_tab_color(const fm_workbook_t* wb, size_t sheet_index, fm_color_spec_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::ColorSpec color;
  pugi::xml_document doc;
  if (load_fragment(doc, wb->workbook().sheet(sheet_index).print_settings().sheet_pr_xml)) {
    const pugi::xml_node node = doc.first_child().child("tabColor");
    if (const pugi::xml_attribute rgb = node.attribute("rgb")) {
      color.kind = formulon::ColorSpec::Kind::kRgb;
      color.rgb = static_cast<std::uint32_t>(std::strtoul(rgb.value(), nullptr, 16));
      // A 6-digit value is RRGGBB with full alpha.
      if (std::string_view(rgb.value()).size() <= 6U) {
        color.rgb |= 0xFF000000U;
      }
    } else if (node.attribute("theme")) {
      color.kind = formulon::ColorSpec::Kind::kTheme;
      color.theme = node.attribute("theme").as_uint();
      color.tint = node.attribute("tint").as_double(0.0);
    } else if (node.attribute("indexed")) {
      color.kind = formulon::ColorSpec::Kind::kIndexed;
      color.indexed = node.attribute("indexed").as_uint();
    } else if (node.attribute("auto").as_bool(false)) {
      color.kind = formulon::ColorSpec::Kind::kAuto;
    }
  }
  *out = formulon::c_api::parts::color_spec_to_c(color);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_tab_color(fm_workbook_t* wb, size_t sheet_index, const fm_color_spec_t* color) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::ColorSpec spec;
  if (color != nullptr) {
    if (auto rc = check_enum_domain(color->kind, kFmColorNone, kFmColorAuto, "kind"); rc != 0) {
      return rc;
    }
    if (color->kind == kFmColorTheme && !std::isfinite(color->tint)) {
      return invalid_argument("theme tint must be finite", "field=tint");
    }
    spec = formulon::c_api::parts::color_spec_from_c(*color);
  }
  set_tab_color(wb->workbook().sheet(sheet_index).mutable_print_settings(), spec);
  return 0;
}
