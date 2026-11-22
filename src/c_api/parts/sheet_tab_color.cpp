//
// C ABI - sheet tab colour: get and set `<sheetPr><tabColor>`.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/print_settings_internal.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/strings.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_sheet;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::load_or_create_root;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_attr;
using formulon::c_api::parts::set_attr_double;
using formulon::c_api::parts::set_attr_u32;
using formulon::c_api::parts::SheetUse;

namespace {

using formulon::SheetPrintSettings;
using formulon::io::raw_xml;

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
          hex[7U - i] = formulon::strings::hex_digit_upper(color.rgb >> (4U * i));
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

}  // namespace

extern "C" fm_status_t fm_sheet_get_tab_color(const fm_workbook_t* wb, size_t sheet_index, fm_color_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  const formulon::ColorSpec& color = wb->workbook().sheet(sheet_index).print_settings().tab_color;
  *out = formulon::c_api::parts::color_spec_to_c(color);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_tab_color(fm_workbook_t* wb, size_t sheet_index, const fm_color_t* color) {
  clear_last_error();
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
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
  // A chartsheet's part takes the colour on save, as a worksheet's does.
  set_tab_color(wb->workbook().sheet(sheet_index).mutable_print_settings(), spec);
  return 0;
}