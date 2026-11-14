//
// Worksheet view and layout reader: `<sheetView>` / `<pane>`, `<sheetPr>`
// tab visibility, `<sheetFormatPr>` defaults, `<cols>` spans and `<row>`
// overrides. See sheet_layout_reader.h for the public contract.

#include "io/sheet_layout_reader.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/sheet_view_xml.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "io/xsd_double.h"
#include "io/xsd_int.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {

// ---------------------------------------------------------------------------
// View / layout helpers. Each lives in an anonymous namespace so the
// translation unit owns its parsing fences; the public driver
// `read_sheet_view_and_layout` composes them in document order.
// ---------------------------------------------------------------------------

namespace {

/// Parses the first window's `<sheetView>` into `view` and keeps the other
/// windows' verbatim. Of two views naming one window Excel keeps the later,
/// at the earlier one's place.
void ApplySheetView(const pugi::xml_node& worksheet, SheetView& view) {
  std::vector<pugi::xml_node> windows;
  for (pugi::xml_node node = worksheet.child("sheetViews").child("sheetView"); node;
       node = node.next_sibling("sheetView")) {
    const std::uint32_t id = attr_u32(node, "workbookViewId", 0U);
    const auto same = std::find_if(windows.begin(), windows.end(), [&](const pugi::xml_node& kept) {
      return attr_u32(kept, "workbookViewId", 0U) == id;
    });
    if (same == windows.end()) {
      windows.push_back(node);
    } else {
      *same = node;
    }
  }
  if (windows.empty()) {
    return;
  }
  read_sheet_view_element(windows.front(), view);
  for (std::size_t i = 1; i < windows.size(); ++i) {
    append_raw_xml(view.extra_views_xml, windows[i]);
  }
}

/// Parses `<sheetPr><tabHidden val="1"/></sheetPr>` into `view`.
/// OOXML also accepts `<sheetPr><tabColor .../>`, but only `tabHidden`
/// is a visibility-affecting flag for the worksheet part itself. The
/// workbook-level `<sheet state="hidden">` form is handled in
/// `read_ooxml`; this helper preserves any prior `tab_hidden` state so
/// the merge is OR-style.
void ApplySheetPrTabHidden(const pugi::xml_node& worksheet, SheetView& view) {
  pugi::xml_node sheet_pr = worksheet.child("sheetPr");
  if (!sheet_pr) {
    return;
  }
  // Some writers emit `tabHidden` as a direct attribute (`<sheetPr
  // tabHidden="1"/>`); others emit it as a child element with a `val`
  // attribute. Accept both shapes.
  if (attr_bool(sheet_pr, "tabHidden")) {
    view.tab_hidden = true;
  }
  if (pugi::xml_node child = sheet_pr.child("tabHidden"); child) {
    if (child.attribute("val")) {
      if (attr_bool(child, "val")) {
        view.tab_hidden = true;
      }
    } else {
      // Bare `<tabHidden/>` element with no `val` attribute is treated
      // as `val="1"` to match what Excel's older writers emit.
      view.tab_hidden = true;
    }
  }
}

/// Parses `<sheetFormatPr defaultColWidth defaultRowHeight baseColWidth
/// customHeight zeroHeight thickTop thickBottom outlineLevelRow
/// outlineLevelCol/>`
/// into `defaults`. The element appears before `<cols>` in the worksheet
/// part. Absent attributes leave the corresponding fields at their
/// struct defaults; `defaultColWidth` / `defaultRowHeight` also set the
/// `has_*` flags so a consumer can distinguish an explicit `0` from an
/// absent attribute.
///
/// A measurement outside the shared non-negative-double lexical space is
/// treated as absent, flag included: these three sizes are what the
/// paginator falls back to for every un-overridden track, so admitting an
/// infinity or a NaN here changes the page count of the whole sheet. A size
/// Excel refuses to open is kept as read.
void ApplySheetFormatDefaults(const pugi::xml_node& worksheet, SheetFormatDefaults& defaults) {
  pugi::xml_node fmt = worksheet.child("sheetFormatPr");
  if (!fmt) {
    return;
  }
  double measurement = 0.0;
  if (parse_xsd_nonneg_double(attr_str(fmt, "defaultColWidth"), &measurement)) {
    defaults.default_col_width = measurement;
    defaults.has_default_col_width = true;
  }
  if (parse_xsd_nonneg_double(attr_str(fmt, "defaultRowHeight"), &measurement)) {
    defaults.default_row_height = xlsx_default_row_height(measurement);
    defaults.has_default_row_height = true;
  }
  if (parse_xsd_nonneg_double(attr_str(fmt, "baseColWidth"), &measurement)) {
    // Excel reads a fractional base width as its whole part.
    defaults.base_col_width = std::floor(measurement);
  }
  defaults.custom_height = read_xsd_bool(fmt, "customHeight", false);
  defaults.zero_height = sheet_format_hides_rows(fmt);
  defaults.thick_top = read_xsd_bool(fmt, "thickTop", false);
  defaults.thick_bottom = read_xsd_bool(fmt, "thickBottom", false);
  defaults.outline_level_row = parse_outline_level(fmt.attribute("outlineLevelRow").value());
  defaults.outline_level_col = parse_outline_level(fmt.attribute("outlineLevelCol").value());
}

/// Parses `<cols><col min max width style hidden outlineLevel/></cols>` into
/// `layout.columns`. Width and style retain attribute presence, so explicit
/// zero values remain distinct from an omitted attribute. A valid span that
/// carries only hidden / outline metadata is retained even when it has no
/// width; pure `customWidth=1` / `bestFit=1` markers remain a no-op.
void ApplyColumnLayouts(const pugi::xml_node& worksheet, SheetLayout& layout) {
  pugi::xml_node cols = worksheet.child("cols");
  if (!cols) {
    return;
  }
  for (pugi::xml_node col = cols.child("col"); col; col = col.next_sibling("col")) {
    if (!col.attribute("min") || !col.attribute("max")) {
      continue;
    }
    const std::int32_t min_v = attr_i32(col, "min", 0);
    const std::int32_t max_v = attr_i32(col, "max", 0);
    if (min_v < 1 || max_v < min_v) {
      continue;
    }
    ColumnLayout entry;
    entry.first = static_cast<std::uint32_t>(min_v - 1);
    entry.last = static_cast<std::uint32_t>(max_v - 1);
    double width = 0.0;
    if (parse_xsd_nonneg_double(attr_str(col, "width"), &width)) {
      entry.width = width;
      entry.has_width = true;
    }
    if (pugi::xml_attribute style_attr = col.attribute("style"); style_attr) {
      entry.style_xf = attr_u32(col, "style", 0U);
      entry.has_style = true;
    }
    if (pugi::xml_attribute hidden_attr = col.attribute("hidden"); hidden_attr) {
      entry.hidden = attr_bool(col, "hidden");
    }
    if (pugi::xml_attribute outline_attr = col.attribute("outlineLevel"); outline_attr) {
      entry.outline_level = parse_outline_level(outline_attr.value());
    }
    entry.show_phonetic = read_xsd_bool(col, "phonetic", false);
    entry.collapsed = read_xsd_bool(col, "collapsed", false);
    if (!entry.has_width && !entry.has_style && !col.attribute("hidden") && !col.attribute("outlineLevel") &&
        !entry.show_phonetic && !entry.collapsed) {
      // A span with only a marker such as `bestFit` has no observable
      // layout state in this model.
      continue;
    }
    layout.columns.push_back(entry);
  }
  normalize_column_spans(layout);
}

/// Walks `<sheetData><row .../></sheetData>` collecting per-row
/// overrides (height / hidden / outline / custom row style) into
/// `layout.row_overrides`. A row `s=` attribute is effective only when
/// `customFormat="1"`; when customFormat is true but `s` is absent, the
/// effective style is the explicit default xf 0.
/// Rows that carry only `r` (the row number) are skipped — they are
/// just position markers and have no override payload — except under
/// `zeroHeight`, where a present row is visible and its entry says so.
void ApplyRowOverrides(const pugi::xml_node& worksheet, SheetLayout& layout, bool default_hidden) {
  pugi::xml_node sheet_data = worksheet.child("sheetData");
  if (!sheet_data) {
    return;
  }
  for (pugi::xml_node row = sheet_data.child("row"); row; row = row.next_sibling("row")) {
    pugi::xml_attribute r_attr = row.attribute("r");
    pugi::xml_attribute ht_attr = row.attribute("ht");
    pugi::xml_attribute hidden_attr = row.attribute("hidden");
    pugi::xml_attribute outline_attr = row.attribute("outlineLevel");
    pugi::xml_attribute custom_format_attr = row.attribute("customFormat");
    pugi::xml_attribute style_attr = row.attribute("s");
    const bool custom_format = custom_format_attr && read_xsd_bool(row, "customFormat", false);
    // An `ht` outside the shared non-negative-double lexical space is
    // treated as absent for both the contribute test below and the stored
    // override, so a row carrying nothing else does not become an
    // all-defaults entry.
    double height = 0.0;
    const bool has_height = ht_attr && parse_xsd_nonneg_double(attr_str(row, "ht"), &height);
    const bool show_phonetic = read_xsd_bool(row, "ph", false);
    const bool collapsed = read_xsd_bool(row, "collapsed", false);
    if (!has_height && !hidden_attr && !outline_attr && !custom_format && !show_phonetic && !collapsed &&
        !default_hidden) {
      continue;
    }
    if (!r_attr) {
      continue;
    }
    // A row number outside the shared non-negative-integer lexical space
    // is treated as absent and the override is dropped, on both read
    // paths: attaching the override to whatever prefix happened to parse
    // would silently restyle an unrelated row.
    std::uint32_t r_v = 0;
    if (!parse_xsd_nonneg_int(attr_str(row, "r"), &r_v) || r_v < 1U) {
      continue;
    }
    RowLayout entry;
    entry.row = r_v - 1U;
    entry.show_phonetic = show_phonetic;
    entry.collapsed = collapsed;
    if (has_height) {
      entry.height = height;
      entry.has_height = true;
      entry.custom_height = read_xsd_bool(row, "customHeight", false);
    }
    if (hidden_attr) {
      entry.hidden = read_xsd_bool(row, "hidden", false);
    }
    if (outline_attr) {
      entry.outline_level = parse_outline_level(outline_attr.value());
    }
    if (custom_format) {
      entry.has_style = true;
      entry.style_xf = style_attr ? attr_u32(row, "s", 0U) : 0U;
    }
    insert_row_override(layout, entry);
  }
}

}  // namespace

double xlsx_default_row_height(double pt) {
  constexpr double kTwipsPerPoint = 20.0;
  return std::floor(std::min(pt, kMaxDefaultRowHeightPt) * kTwipsPerPoint) / kTwipsPerPoint;
}

bool sheet_format_hides_rows(const pugi::xml_node& sheet_format_pr) {
  if (read_xsd_bool(sheet_format_pr, "zeroHeight", false)) {
    return true;
  }
  double height = 0.0;
  return read_xsd_bool(sheet_format_pr, "customHeight", false) &&
         parse_xsd_nonneg_double(attr_str(sheet_format_pr, "defaultRowHeight"), &height) &&
         xlsx_default_row_height(height) == 0.0;
}

Expected<void, Error> read_sheet_view_and_layout(const pugi::xml_document& sheet_doc, std::size_t sheet_index,
                                                 Workbook& workbook) {
  if (sheet_index >= workbook.sheet_count()) {
    std::string ctx("context=sheet_reader.view_layout sheet_index=");
    ctx.append(std::to_string(sheet_index));
    ctx.append(" sheet_count=");
    ctx.append(std::to_string(workbook.sheet_count()));
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet_index out of range", std::move(ctx));
  }
  pugi::xml_node worksheet = sheet_doc.child("worksheet");
  if (!worksheet) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "sheet doc: missing <worksheet> root",
                      "context=sheet_reader.view_layout");
  }
  Sheet& sheet = workbook.sheet(sheet_index);
  SheetView& view = sheet.mutable_view();
  SheetLayout& layout = sheet.mutable_layout();
  ApplySheetView(worksheet, view);
  ApplySheetPrTabHidden(worksheet, view);
  ApplySheetFormatDefaults(worksheet, sheet.mutable_format_defaults());
  ApplyColumnLayouts(worksheet, layout);
  ApplyRowOverrides(worksheet, layout, sheet.format_defaults().zero_height);
  return Expected<void, Error>::Ok();
}

}  // namespace io
}  // namespace formulon
