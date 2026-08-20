
#include "io/ooxml/print_settings_parse.h"

#include <algorithm>
#include <cstdint>
#include <string_view>
#include <vector>

#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "io/xsd_double.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/resource_budget.h"

namespace formulon {
namespace io {
namespace ooxml {

void apply_structured_page_setup(const pugi::xml_node& page_setup, PageSetup& out) {
  if (pugi::xml_attribute attr = page_setup.attribute("orientation"); attr) {
    const std::string_view value = attr.value();
    if (value == "portrait") {
      out.orientation = Orientation::kPortrait;
    } else if (value == "landscape") {
      out.orientation = Orientation::kLandscape;
    } else {
      out.orientation = Orientation::kDefault;
    }
  }
  if (pugi::xml_attribute attr = page_setup.attribute("paperSize"); attr) {
    out.paper_size = static_cast<std::uint32_t>(attr.as_uint(out.paper_size));
  }
  if (pugi::xml_attribute attr = page_setup.attribute("scale"); attr) {
    out.scale = static_cast<std::uint32_t>(attr.as_uint(out.scale));
  }
  if (pugi::xml_attribute attr = page_setup.attribute("fitToWidth"); attr) {
    out.fit_to_width = static_cast<std::uint32_t>(attr.as_uint(out.fit_to_width));
  }
  if (pugi::xml_attribute attr = page_setup.attribute("fitToHeight"); attr) {
    out.fit_to_height = static_cast<std::uint32_t>(attr.as_uint(out.fit_to_height));
  }
}

void apply_structured_page_margins(const pugi::xml_node& page_margins, PageMargins& out) {
  const auto margin = [&page_margins](const char* name, double& field) {
    double value = 0.0;
    if (parse_xsd_nonneg_double(attr_str(page_margins, name), &value)) {
      field = value;
    }
  };
  margin("left", out.left);
  margin("right", out.right);
  margin("top", out.top);
  margin("bottom", out.bottom);
  margin("header", out.header);
  margin("footer", out.footer);
}

void read_manual_breaks(const pugi::xml_node& breaks_node, std::vector<ManualBreak>& out) {
  if (!breaks_node) {
    return;
  }
  for (pugi::xml_node brk = breaks_node.child("brk"); brk; brk = brk.next_sibling("brk")) {
    ManualBreak entry;
    // `id` is already the 0-based index the break precedes: Excel writes
    // `<brk id="20"/>` for a break placed before row 21, i.e. the count of
    // rows on the page the break ends. Subtracting one here moved every
    // manual break in an Excel-authored file one track early, and the
    // writer's matching increment put it back on save -- so a Formulon
    // read/write cycle looked clean while Formulon and Excel disagreed
    // about where the page ended. Measured against Excel 365: authoring a
    // manual break before row 21 / column D produces `id="20"` / `id="3"`.
    entry.id = static_cast<std::uint32_t>(brk.attribute("id").as_uint(0));
    entry.min = static_cast<std::uint32_t>(brk.attribute("min").as_uint(0));
    entry.max = static_cast<std::uint32_t>(brk.attribute("max").as_uint(0));
    // `man` defaults to false (ECMA-376 §18.3.1.1): a break without it is
    // automatic and must not be re-emitted as a user break.
    entry.manual = read_xsd_bool(brk, "man", false);
    out.push_back(entry);
  }
  // Every other producer of these vectors - the C ABI's upsert / erase
  // pair, and the break shifting that follows a structural edit - keeps
  // them strictly increasing by `id`, and the consumers rely on it: the
  // mutators binary-search, the paginator scans forward, and the
  // enumeration API documents ascending order. A file is under no such
  // obligation, and a third-party writer emitting `<brk>` in authoring
  // order rather than sheet order would leave a loaded workbook where
  // `fm_sheet_remove_col_break` silently matches nothing and an upsert
  // appends a duplicate index. Normalising here makes the invariant hold
  // for the load path too, at the cost of the document order of a list
  // whose order carries no meaning.
  // Stable, so a repeated `id` keeps the span the document stated first
  // rather than an arbitrary one of the duplicates.
  std::stable_sort(out.begin(), out.end(), [](const ManualBreak& a, const ManualBreak& b) { return a.id < b.id; });
  out.erase(
      std::unique(out.begin(), out.end(), [](const ManualBreak& a, const ManualBreak& b) { return a.id == b.id; }),
      out.end());
  // The loader must not accept more breaks than the authoring API does,
  // or reading a file and adding one break to it would fail where the
  // same edit on a freshly built sheet succeeds.
  if (out.size() > kMaxManualBreaksPerAxis) {
    out.resize(kMaxManualBreaksPerAxis);
  }
}

bool read_fit_to_page(const pugi::xml_node& sheet_pr) {
  const pugi::xml_node page_setup_pr = sheet_pr.child("pageSetUpPr");
  if (!page_setup_pr) {
    return false;
  }
  return page_setup_pr.attribute("fitToPage").as_bool(false);
}

void refresh_structured_views(std::string_view element_name, std::string_view fragment, SheetPrintSettings& settings) {
  // `printOptions` and `headerFooter` have no structured projection; their
  // raw fragment is the whole model, so there is nothing to re-derive.
  const bool is_page_setup = element_name == "pageSetup";
  const bool is_page_margins = element_name == "pageMargins";
  const bool is_sheet_pr = element_name == "sheetPr";
  if (!is_page_setup && !is_page_margins && !is_sheet_pr) {
    return;
  }

  // Reset first so an attribute the new fragment omits reverts to the
  // ECMA-376 default rather than lingering from the previous fragment.
  // `fit_to_page` is owned by `<sheetPr>`, so each branch resets only the
  // fields its own element can state.
  const PageSetup defaults;
  if (is_page_setup) {
    settings.page_setup.orientation = defaults.orientation;
    settings.page_setup.paper_size = defaults.paper_size;
    settings.page_setup.scale = defaults.scale;
    settings.page_setup.fit_to_width = defaults.fit_to_width;
    settings.page_setup.fit_to_height = defaults.fit_to_height;
  } else if (is_page_margins) {
    settings.page_margins = PageMargins{};
  } else {
    settings.page_setup.fit_to_page = defaults.fit_to_page;
  }

  if (fragment.empty()) {
    return;
  }

  pugi::xml_document doc;
  if (!doc.load_buffer(fragment.data(), fragment.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return;
  }
  const pugi::xml_node root = doc.first_child();
  if (!root) {
    return;
  }
  if (is_page_setup) {
    apply_structured_page_setup(root, settings.page_setup);
  } else if (is_page_margins) {
    apply_structured_page_margins(root, settings.page_margins);
  } else {
    settings.page_setup.fit_to_page = read_fit_to_page(root);
  }
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
