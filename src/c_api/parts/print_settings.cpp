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
//   - typed patch structs, which touch only the attributes the patch
//     sets or clears and leave every other one (including ones this engine does
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
#include <initializer_list>
#include <iterator>
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
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
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
      return invalid_argument("xml fragment is not a valid element",
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

/// The status of a patch op outside 0..2, or a CLEAR on a field that cannot be cleared.
fm_status_t invalid_patch_op() {
  return invalid_argument("op is not a patch op");
}

/// True when every op of `ops` is KEEP or SET.
bool ops_without_clear(std::initializer_list<fm_patch_op_t> ops) {
  return std::all_of(ops.begin(), ops.end(), patch_op_ok_without_clear);
}

/// One `(attribute, op, value)` triple of a patched boolean attribute.
struct BoolAttr {
  const char* attr;
  fm_patch_op_t op;
  std::int32_t value;
};

/// One `(attribute, op, value)` triple of a patched unsigned attribute.
struct U32Attr {
  const char* attr;
  fm_patch_op_t op;
  std::uint32_t value;
};

/// Writes the `<pageSetup>` half of a page-setup patch and re-derives the view.
void write_page_setup(SheetPrintSettings& print, const fm_page_setup_patch_t& patch, const U32Attr* attrs,
                      std::size_t count) {
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.page_setup_xml, "pageSetup");
  if (patch.orientation_op == FM_PATCH_OP_SET) {
    set_attr(root, "orientation", patch.orientation == FM_PAGE_ORIENTATION_LANDSCAPE ? "landscape" : "portrait");
  } else if (patch.orientation_op == FM_PATCH_OP_CLEAR) {
    root.remove_attribute("orientation");
  }
  for (std::size_t i = 0; i < count; ++i) {
    if (attrs[i].op == FM_PATCH_OP_SET) {
      set_attr_u32(root, attrs[i].attr, attrs[i].value);
    }
  }
  print.page_setup_xml = raw_xml(root);
  refresh_structured_views("pageSetup", print.page_setup_xml, print);
}

/// True when no entry of `attrs` carries an op other than KEEP or SET.
bool bool_ops_ok(const BoolAttr* attrs, std::size_t count) {
  return std::all_of(attrs, attrs + count, [](const BoolAttr& a) { return patch_op_ok_without_clear(a.op); });
}

/// True when some entry of `attrs` is SET.
bool any_bool_set(const BoolAttr* attrs, std::size_t count) {
  return std::any_of(attrs, attrs + count, [](const BoolAttr& a) { return a.op == FM_PATCH_OP_SET; });
}

/// Applies every SET entry of `attrs` through `set_attr_bool`.
void set_bools(pugi::xml_node node, const BoolAttr* attrs, std::size_t count) {
  for (std::size_t i = 0; i < count; ++i) {
    if (attrs[i].op == FM_PATCH_OP_SET) {
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
  print.tab_color = color;
}

/// Loads a stored print fragment for a getter; false when it is absent.
bool load_fragment(pugi::xml_document& doc, const std::string& fragment) {
  return !fragment.empty() &&
         doc.load_buffer(fragment.data(), fragment.size(), pugi::parse_default, pugi::encoding_utf8);
}

/// Upserts a manual break into an axis vector, keeping it sorted by index.
fm_status_t upsert_break(std::vector<ManualBreak>& breaks, std::uint32_t index, std::uint32_t first, std::uint32_t last,
                         bool manual) {
  const auto pos = std::lower_bound(breaks.begin(), breaks.end(), index,
                                    [](const ManualBreak& entry, std::uint32_t key) { return entry.id < key; });
  if (pos != breaks.end() && pos->id == index) {
    pos->min = first;
    pos->max = last;
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
  entry.min = first;
  entry.max = last;
  entry.manual = manual;
  breaks.insert(pos, entry);
  return 0;
}

/// Erases the break at `index`; a no-op when the axis has none there.
void erase_break(std::vector<ManualBreak>& breaks, std::uint32_t index) {
  const auto pos = std::lower_bound(breaks.begin(), breaks.end(), index,
                                    [](const ManualBreak& entry, std::uint32_t key) { return entry.id < key; });
  if (pos != breaks.end() && pos->id == index) {
    breaks.erase(pos);
  }
}

/// Shared body of `fm_sheet_set_row_break` / `fm_sheet_set_column_break`.
/// A null record removes the break; `first == last == 0` takes the whole span.
fm_status_t set_break(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t index, bool rows,
                      const std::uint32_t* first_in, const std::uint32_t* last_in, const int32_t* manual_in) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = rows ? check_row_index(index) : check_column_index(index); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  std::vector<ManualBreak>& breaks = rows ? print.manual_row_breaks : print.manual_col_breaks;
  if (first_in == nullptr) {
    erase_break(breaks, index);
    return 0;
  }
  const std::uint32_t span_max = rows ? kWholeSheetColSpanMax : kWholeSheetRowSpanMax;
  std::uint32_t first = *first_in;
  std::uint32_t last = *last_in;
  if (first == 0U && last == 0U) {
    last = span_max;
  }
  if (first > last || last > span_max) {
    return invalid_argument("first and last must be within the break span",
                            "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  if (*manual_in < FM_TRISTATE_DEFAULT || *manual_in > FM_TRISTATE_TRUE) {
    return invalid_argument(("manual " + std::to_string(*manual_in) + " is not a tristate").c_str(),
                            "manual=" + std::to_string(*manual_in));
  }
  return upsert_break(breaks, index, first, last, *manual_in != FM_TRISTATE_FALSE);
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

template <typename Record>
fm_status_t break_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index, bool rows, Record* out,
                     std::uint32_t Record::*position) {
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
  out->*position = entry.id;
  out->first = entry.min;
  out->last = entry.max;
  out->manual = entry.manual ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
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
    return invalid_argument("xml references a printer settings part that does not exist");
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

extern "C" fm_status_t fm_sheet_set_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                              const fm_row_break_t* brk) {
  clear_last_error();
  if (brk != nullptr) {
    if (auto rc = check_record_size(*brk); rc != 0) {
      return rc;
    }
    return set_break(wb, sheet_index, row, /*rows=*/true, &brk->first, &brk->last, &brk->manual);
  }
  return set_break(wb, sheet_index, row, /*rows=*/true, nullptr, nullptr, nullptr);
}

extern "C" fm_status_t fm_sheet_set_column_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col,
                                                 const fm_column_break_t* brk) {
  clear_last_error();
  if (brk != nullptr) {
    if (auto rc = check_record_size(*brk); rc != 0) {
      return rc;
    }
    return set_break(wb, sheet_index, col, /*rows=*/false, &brk->first, &brk->last, &brk->manual);
  }
  return set_break(wb, sheet_index, col, /*rows=*/false, nullptr, nullptr, nullptr);
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
                                                 fm_row_break_t* out) {
  return break_at(wb, sheet_index, index, /*rows=*/true, out, &fm_row_break_t::row);
}

extern "C" fm_status_t fm_sheet_get_column_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  return break_count(wb, sheet_index, /*rows=*/false, out_count);
}

extern "C" fm_status_t fm_sheet_get_column_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                    fm_column_break_t* out) {
  return break_at(wb, sheet_index, index, /*rows=*/false, out, &fm_column_break_t::col);
}

/* -------------------------------------------------------------------------- */
/* Typed patch setters                                                        */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_update_page_setup(fm_workbook_t* wb, size_t sheet_index,
                                                  const fm_page_setup_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (!patch_op_ok(patch->orientation_op) ||
      !ops_without_clear({patch->paper_size_op, patch->scale_pct_op, patch->fit_to_width_op, patch->fit_to_height_op,
                          patch->fit_to_page_op})) {
    return invalid_patch_op();
  }
  if (patch->orientation_op == FM_PATCH_OP_SET) {
    if (auto rc = check_enum_domain(patch->orientation, 0, FM_PAGE_ORIENTATION_LANDSCAPE, "orientation"); rc != 0) {
      return rc;
    }
  }
  if (patch->scale_pct_op == FM_PATCH_OP_SET &&
      (patch->scale_pct < kMinPrintScalePercent || patch->scale_pct > kMaxPrintScalePercent)) {
    return invalid_argument(
        ("scale_pct must be in " + std::to_string(kMinPrintScalePercent) + ".." + std::to_string(kMaxPrintScalePercent))
            .c_str(),
        "scale_pct=" + std::to_string(patch->scale_pct));
  }

  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  const U32Attr attrs[] = {
      {"paperSize", patch->paper_size_op, patch->paper_size},
      {"scale", patch->scale_pct_op, patch->scale_pct},
      {"fitToWidth", patch->fit_to_width_op, patch->fit_to_width},
      {"fitToHeight", patch->fit_to_height_op, patch->fit_to_height},
  };
  const bool touches_page_setup =
      patch->orientation_op != FM_PATCH_OP_KEEP ||
      std::any_of(std::begin(attrs), std::end(attrs), [](const U32Attr& a) { return a.op != FM_PATCH_OP_KEEP; });
  if (touches_page_setup) {
    write_page_setup(print, *patch, attrs, std::size(attrs));
  }

  // `fitToPage` is not a `<pageSetup>` attribute; it rides along on this
  // patch because callers think of it as part of one page-setup decision.
  if (patch->fit_to_page_op == FM_PATCH_OP_SET) {
    set_fit_to_page(print, patch->fit_to_page != 0);
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_update_page_margins(fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_page_margins_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  struct MarginField {
    const char* attr;
    fm_patch_op_t op;
    double value;
  };
  // Only the SET fields are stored. ECMA-376 requires all six on the
  // saved element, but completing it here would report every margin as
  // engaged on the next read and erase the caller's ability to tell a
  // stated margin from a defaulted one. The OOXML writer fills the gap
  // when it emits the element instead.
  const MarginField fields[] = {
      {"left", patch->left_in_op, patch->left_in},       {"right", patch->right_in_op, patch->right_in},
      {"top", patch->top_in_op, patch->top_in},          {"bottom", patch->bottom_in_op, patch->bottom_in},
      {"header", patch->header_in_op, patch->header_in}, {"footer", patch->footer_in_op, patch->footer_in},
  };
  for (const MarginField& field : fields) {
    if (!patch_op_ok_without_clear(field.op)) {
      return invalid_patch_op();
    }
  }
  for (const MarginField& field : fields) {
    // The paginator subtracts these from the paper; a negative, infinite or
    // NaN one produces a printable body that is not a rectangle.
    if (field.op == FM_PATCH_OP_SET && (!std::isfinite(field.value) || field.value < 0.0)) {
      return invalid_argument((std::string(field.attr) + " margin must be finite and not negative").c_str(),
                              std::string(field.attr) + "=invalid");
    }
  }

  if (std::all_of(std::begin(fields), std::end(fields),
                  [](const MarginField& field) { return field.op == FM_PATCH_OP_KEEP; })) {
    return 0;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.page_margins_xml, "pageMargins");
  for (const MarginField& field : fields) {
    if (field.op == FM_PATCH_OP_SET) {
      set_attr_double(root, field.attr, field.value);
    }
  }
  print.page_margins_xml = raw_xml(root);
  refresh_structured_views("pageMargins", print.page_margins_xml, print);
  return 0;
}

extern "C" fm_status_t fm_sheet_update_print_options(fm_workbook_t* wb, size_t sheet_index,
                                                     const fm_print_options_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const BoolAttr attrs[] = {
      {"gridLines", patch->grid_lines_op, patch->grid_lines},
      {"headings", patch->headings_op, patch->headings},
      {"horizontalCentered", patch->horizontal_centered_op, patch->horizontal_centered},
      {"verticalCentered", patch->vertical_centered_op, patch->vertical_centered},
  };
  if (!bool_ops_ok(attrs, std::size(attrs))) {
    return invalid_patch_op();
  }
  if (!any_bool_set(attrs, std::size(attrs))) {
    return 0;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.print_options_xml, "printOptions");
  set_bools(root, attrs, std::size(attrs));
  print.print_options_xml = raw_xml(root);
  return 0;
}

extern "C" fm_status_t fm_sheet_update_header_footer(fm_workbook_t* wb, size_t sheet_index,
                                                     const fm_header_footer_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const BoolAttr attrs[] = {
      {"differentOddEven", patch->different_odd_even_op, patch->different_odd_even},
      {"differentFirst", patch->different_first_op, patch->different_first},
      {"scaleWithDoc", patch->scale_with_doc_op, patch->scale_with_doc},
      {"alignWithMargins", patch->align_with_margins_op, patch->align_with_margins},
  };
  if (!bool_ops_ok(attrs, std::size(attrs))) {
    return invalid_patch_op();
  }
  struct Section {
    fm_patch_op_t op;
    const char* text;
    const char* name;
  };
  // In `kHeaderFooterSections` order.
  const Section sections[] = {
      {patch->odd_header_op, patch->odd_header, "odd_header"},
      {patch->odd_footer_op, patch->odd_footer, "odd_footer"},
      {patch->even_header_op, patch->even_header, "even_header"},
      {patch->even_footer_op, patch->even_footer, "even_footer"},
      {patch->first_header_op, patch->first_header, "first_header"},
      {patch->first_footer_op, patch->first_footer, "first_footer"},
  };
  for (const Section& section : sections) {
    if (!patch_op_ok(section.op)) {
      return invalid_patch_op();
    }
    if (section.op == FM_PATCH_OP_SET && section.text == nullptr) {
      return null_field(section.name);
    }
  }
  for (const Section& section : sections) {
    if (section.op == FM_PATCH_OP_SET) {
      if (auto rc = check_utf8({section.text}); rc != 0) {
        return rc;
      }
    }
  }
  if (!any_bool_set(attrs, std::size(attrs)) &&
      std::all_of(std::begin(sections), std::end(sections),
                  [](const Section& section) { return section.op == FM_PATCH_OP_KEEP; })) {
    return 0;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  pugi::xml_document doc;
  pugi::xml_node root = load_or_create_root(doc, print.header_footer_xml, "headerFooter");
  set_bools(root, attrs, std::size(attrs));
  for (std::size_t i = 0; i < std::size(sections); ++i) {
    if (sections[i].op == FM_PATCH_OP_SET) {
      // An empty section removes the child, as CLEAR does.
      set_header_footer_section(root, i, sections[i].text);
    } else if (sections[i].op == FM_PATCH_OP_CLEAR) {
      set_header_footer_section(root, i, std::string_view());
    }
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
  out->scale_pct = view.scale;
  out->fit_to_width = view.fit_to_width;
  out->fit_to_height = view.fit_to_height;
  out->fit_to_page = view.fit_to_page ? 1 : 0;

  // The `has_*` flags report presence, which only the XML knows: the
  // structured view cannot distinguish `scale="100"` from an absent
  // attribute that defaults to 100.
  pugi::xml_document doc;
  if (load_fragment(doc, print.page_setup_xml)) {
    const pugi::xml_node root = doc.first_child();
    const EngagedAttr attrs[] = {
        {"orientation", &out->has_orientation},   {"paperSize", &out->has_paper_size},
        {"scale", &out->has_scale_pct},           {"fitToWidth", &out->has_fit_to_width},
        {"fitToHeight", &out->has_fit_to_height},
    };
    read_engaged(root, attrs, std::size(attrs));
  }
  pugi::xml_document sheet_pr_doc;
  if (load_fragment(sheet_pr_doc, print.sheet_pr_xml)) {
    const pugi::xml_node page_setup_pr = sheet_pr_doc.first_child().child("pageSetUpPr");
    out->has_fit_to_page = (page_setup_pr && page_setup_pr.attribute("fitToPage")) ? 1 : 0;
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
  out->left_in = view.left;
  out->right_in = view.right;
  out->top_in = view.top;
  out->bottom_in = view.bottom;
  out->header_in = view.header;
  out->footer_in = view.footer;

  pugi::xml_document doc;
  if (load_fragment(doc, print.page_margins_xml)) {
    const pugi::xml_node root = doc.first_child();
    const EngagedAttr attrs[] = {
        {"left", &out->has_left_in},     {"right", &out->has_right_in},   {"top", &out->has_top_in},
        {"bottom", &out->has_bottom_in}, {"header", &out->has_header_in}, {"footer", &out->has_footer_in},
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
      {"gridLines", &out->has_grid_lines},
      {"headings", &out->has_headings},
      {"horizontalCentered", &out->has_horizontal_centered},
      {"verticalCentered", &out->has_vertical_centered},
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
      {"differentOddEven", &out->has_different_odd_even},
      {"differentFirst", &out->has_different_first},
      {"scaleWithDoc", &out->has_scale_with_doc},
      {"alignWithMargins", &out->has_align_with_margins},
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

extern "C" fm_status_t fm_sheet_get_tab_color(const fm_workbook_t* wb, size_t sheet_index, fm_color_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const formulon::ColorSpec& color = wb->workbook().sheet(sheet_index).print_settings().tab_color;
  *out = formulon::c_api::parts::color_spec_to_c(color);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_tab_color(fm_workbook_t* wb, size_t sheet_index, const fm_color_t* color) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::ColorSpec spec;
  if (color != nullptr) {
    if (auto rc = check_enum_domain(color->kind, FM_COLOR_KIND_NONE, FM_COLOR_KIND_AUTO, "kind"); rc != 0) {
      return rc;
    }
    if (color->kind == FM_COLOR_KIND_THEME && !std::isfinite(color->tint)) {
      return invalid_argument("tint must be finite", "field=tint");
    }
    spec = formulon::c_api::parts::color_spec_from_c(*color);
  }
  set_tab_color(wb->workbook().sheet(sheet_index).mutable_print_settings(), spec);
  return 0;
}
