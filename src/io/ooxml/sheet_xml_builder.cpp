//
// Worksheet-part XML builders for the OOXML writer. See header for the
// caller contract; this TU owns all of the per-sheet body builders that
// were previously parked alongside the orchestrator in
// `src/io/ooxml_writer.cpp`.

#include "io/ooxml/sheet_xml_builder.h"

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/cf_overlay.h"
#include "io/cf_writer.h"
#include "io/ooxml/cell_ref_writer.h"
#include "io/ooxml/emission_plan.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/shared_strings_writer.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_writer_cell.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"

namespace formulon {
namespace io {
namespace {

// Forward declarations for sheet-view / column-layout builders. The
// definitions live at the bottom of the file alongside the other XML
// helpers; their signatures are needed up here so `BuildWorksheetXml`
// can call them.
std::string BuildSheetViewXml(const SheetView& view);
std::string BuildSheetFormatPrXml(const SheetFormatDefaults& defaults);
std::string BuildColsXml(const SheetLayout& layout);
std::string BuildSheetProtectionXml(const SheetProtection& p);
std::string BuildDimensionXml(const Sheet& sheet);

// Sheet visibility is owned by xl/workbook.xml's <sheet state="...">
// attribute.  Older producers also put tabHidden on the worksheet's raw
// <sheetPr>; preserving that stale flag would make a sheet hidden again after
// an API caller explicitly re-shows it.  Strip the legacy representation on
// every write and let WorkbookXmlBuilder emit the current model state.
std::string BuildNormalizedSheetPrXml(std::string_view raw_sheet_pr) {
  if (raw_sheet_pr.empty()) {
    return {};
  }

  pugi::xml_document doc;
  const pugi::xml_parse_result parsed = doc.load_buffer(raw_sheet_pr.data(), raw_sheet_pr.size());
  pugi::xml_node sheet_pr = doc.document_element();
  if (!parsed || !sheet_pr || std::string_view(sheet_pr.name()) != "sheetPr") {
    // Raw metadata comes from a successfully parsed source workbook, but do
    // not discard it should an unusual extension make the fragment unparsable
    // out of its original namespace context.
    return std::string(raw_sheet_pr);
  }

  sheet_pr.remove_attribute("tabHidden");
  for (pugi::xml_node child = sheet_pr.child("tabHidden"); child;) {
    const pugi::xml_node next = child.next_sibling("tabHidden");
    sheet_pr.remove_child(child);
    child = next;
  }

  return raw_xml(sheet_pr);
}

std::string BuildMergeCellsBlock(const Sheet& sheet) {
  if (sheet.merges().empty()) {
    return {};
  }
  std::string out;
  out.reserve(64 + sheet.merges().size() * 32);
  out.append("<mergeCells count=\"");
  out.append(std::to_string(sheet.merges().size()));
  out.append("\">");
  for (const MergeRange& m : sheet.merges()) {
    out.append("<mergeCell ref=\"");
    AppendRangeRef(out, m);
    out.append("\"/>");
  }
  out.append("</mergeCells>");
  return out;
}

std::string_view DataValidationTypeToString(std::uint8_t type) {
  switch (type) {
    case 1:
      return "whole";
    case 2:
      return "decimal";
    case 3:
      return "list";
    case 4:
      return "date";
    case 5:
      return "time";
    case 6:
      return "textLength";
    case 7:
      return "custom";
    default:
      return "";
  }
}

std::string_view DataValidationOperatorToString(std::uint8_t op) {
  switch (op) {
    case 1:
      return "notBetween";
    case 2:
      return "equal";
    case 3:
      return "notEqual";
    case 4:
      return "greaterThan";
    case 5:
      return "lessThan";
    case 6:
      return "greaterThanOrEqual";
    case 7:
      return "lessThanOrEqual";
    default:
      return "";  // 0 == between, omitted
  }
}

std::string_view DataValidationErrorStyleToString(std::uint8_t style) {
  switch (style) {
    case 1:
      return "warning";
    case 2:
      return "information";
    default:
      return "";  // 0 == stop (default)
  }
}

std::string BuildDataValidationsBlock(const Sheet& sheet) {
  if (sheet.validations().empty()) {
    return {};
  }
  std::string out;
  out.reserve(96 + sheet.validations().size() * 96);
  out.append("<dataValidations count=\"");
  out.append(std::to_string(sheet.validations().size()));
  out.append("\">");
  for (const DataValidation& v : sheet.validations()) {
    out.append("<dataValidation");
    if (const std::string_view t = DataValidationTypeToString(v.type); !t.empty()) {
      out.append(" type=\"");
      out.append(t);
      out.append("\"");
    }
    if (const std::string_view op = DataValidationOperatorToString(v.op); !op.empty()) {
      out.append(" operator=\"");
      out.append(op);
      out.append("\"");
    }
    if (const std::string_view es = DataValidationErrorStyleToString(v.error_style); !es.empty()) {
      out.append(" errorStyle=\"");
      out.append(es);
      out.append("\"");
    }
    if (!v.allow_blank) {
      out.append(" allowBlank=\"0\"");
    } else {
      out.append(" allowBlank=\"1\"");
    }
    if (v.show_input_message) {
      out.append(" showInputMessage=\"1\"");
    }
    if (v.show_error_message) {
      out.append(" showErrorMessage=\"1\"");
    }
    // `showDropDown` is inverted per ECMA-376: writing "1" SUPPRESSES the
    // in-cell arrow, so it is only emitted when the arrow should be
    // hidden. Omitting the attribute preserves Excel's default (shown).
    if (!v.show_dropdown) {
      out.append(" showDropDown=\"1\"");
    }
    if (!v.error_title.empty()) {
      out.append(" errorTitle=\"");
      AppendXmlAttrEscaped(out, v.error_title);
      out.append("\"");
    }
    if (!v.error_message.empty()) {
      out.append(" error=\"");
      AppendXmlAttrEscaped(out, v.error_message);
      out.append("\"");
    }
    if (!v.prompt_title.empty()) {
      out.append(" promptTitle=\"");
      AppendXmlAttrEscaped(out, v.prompt_title);
      out.append("\"");
    }
    if (!v.prompt_message.empty()) {
      out.append(" prompt=\"");
      AppendXmlAttrEscaped(out, v.prompt_message);
      out.append("\"");
    }
    out.append(" sqref=\"");
    for (std::size_t i = 0; i < v.ranges.size(); ++i) {
      if (i > 0) {
        out.push_back(' ');
      }
      AppendRangeRef(out, v.ranges[i]);
    }
    out.append("\">");
    if (!v.formula1.empty()) {
      out.append("<formula1>");
      AppendXmlEscaped(out, v.formula1);
      out.append("</formula1>");
    }
    if (!v.formula2.empty()) {
      out.append("<formula2>");
      AppendXmlEscaped(out, v.formula2);
      out.append("</formula2>");
    }
    out.append("</dataValidation>");
  }
  out.append("</dataValidations>");
  return out;
}

/// Builds the `<hyperlinks>` block. `rid_for_index` returns the rels-file
/// rId integer assigned to the hyperlink at index `i`, or 0 when no rId
/// was assigned (defensive — every external hyperlink gets one).
std::string BuildHyperlinksBlock(const Sheet& sheet, const std::vector<std::string>& rid_per_hyperlink) {
  if (sheet.hyperlinks().empty()) {
    return {};
  }
  std::string out;
  out.reserve(48 + sheet.hyperlinks().size() * 96);
  out.append("<hyperlinks>");
  for (std::size_t i = 0; i < sheet.hyperlinks().size(); ++i) {
    const Hyperlink& h = sheet.hyperlinks()[i];
    out.append("<hyperlink ref=\"");
    AppendCellRefForRef(out, h.row, h.col);
    if (h.row != h.last_row || h.col != h.last_col) {
      out.push_back(':');
      AppendCellRefForRef(out, h.last_row, h.last_col);
    }
    out.append("\"");
    if (i < rid_per_hyperlink.size() && !rid_per_hyperlink[i].empty()) {
      out.append(" r:id=\"");
      AppendXmlAttrEscaped(out, rid_per_hyperlink[i]);
      out.append("\"");
    }
    if (!h.location.empty()) {
      out.append(" location=\"");
      AppendXmlAttrEscaped(out, h.location);
      out.append("\"");
    }
    if (!h.tooltip.empty()) {
      out.append(" tooltip=\"");
      AppendXmlAttrEscaped(out, h.tooltip);
      out.append("\"");
    }
    if (!h.display.empty()) {
      out.append(" display=\"");
      AppendXmlAttrEscaped(out, h.display);
      out.append("\"");
    }
    out.append("/>");
  }
  out.append("</hyperlinks>");
  return out;
}

std::string PageSetupWithRelationshipId(std::string page_setup_xml, std::string_view rid) {
  if (page_setup_xml.empty() || rid.empty()) {
    return page_setup_xml;
  }
  auto replace_attr = [&](std::string_view attr_name) {
    const std::string needle = std::string(attr_name) + "=\"";
    const std::size_t pos = page_setup_xml.find(needle);
    if (pos == std::string::npos) {
      return false;
    }
    const std::size_t value_start = pos + needle.size();
    const std::size_t value_end = page_setup_xml.find('"', value_start);
    if (value_end == std::string::npos) {
      return false;
    }
    page_setup_xml.replace(value_start, value_end - value_start, rid);
    return true;
  };
  if (replace_attr("r:id") || replace_attr("id")) {
    return page_setup_xml;
  }
  const std::size_t insert_pos = page_setup_xml.rfind("/>");
  const std::size_t fallback_pos = page_setup_xml.rfind('>');
  const std::size_t pos = insert_pos != std::string::npos ? insert_pos : fallback_pos;
  if (pos == std::string::npos) {
    return page_setup_xml;
  }
  std::string attr(" r:id=\"");
  attr.append(rid);
  attr.push_back('"');
  page_setup_xml.insert(pos, attr);
  return page_setup_xml;
}

/// Emits a `<rowBreaks>` or `<colBreaks>` block for the given manual
/// breaks. Returns an empty string when `breaks` is empty so the caller
/// adds no bytes. The stored 0-based break index is converted back to
/// OOXML's 1-based form; `count` and `manualBreakCount` mirror the entry
/// count.
std::string BuildPageBreaksXml(std::string_view element, const std::vector<ManualBreak>& breaks) {
  if (breaks.empty()) {
    return {};
  }
  // Rough size estimate so the buffer rarely reallocates: a fixed
  // allowance for the wrapper element plus one allowance per `<brk>`.
  constexpr std::size_t kBreaksWrapperReserveBytes = 48U;
  constexpr std::size_t kPerBreakReserveBytes = 48U;
  std::string out;
  out.reserve(kBreaksWrapperReserveBytes + breaks.size() * kPerBreakReserveBytes);
  const std::string count = std::to_string(breaks.size());
  out.push_back('<');
  out.append(element);
  out.append(" count=\"");
  out.append(count);
  out.append("\" manualBreakCount=\"");
  out.append(count);
  out.append("\">");
  for (const ManualBreak& brk : breaks) {
    out.append("<brk id=\"");
    out.append(std::to_string(static_cast<std::uint64_t>(brk.id) + 1U));
    out.append("\" min=\"");
    out.append(std::to_string(brk.min));
    out.append("\" max=\"");
    out.append(std::to_string(brk.max));
    out.append("\"");
    if (brk.manual) {
      out.append(" man=\"1\"");
    }
    out.append("/>");
  }
  out.append("</");
  out.append(element);
  out.push_back('>');
  return out;
}

// ---------------------------------------------------------------------------
// View / layout part builders
// ---------------------------------------------------------------------------
//
// These helpers each return either an empty string (when the underlying
// field is at its default value, meaning the worksheet XML omits the
// element entirely) or a self-contained XML chunk that the caller
// inserts inside the `<worksheet>` element. Keeping them local to this
// translation unit avoids cross-bundle collisions; they consume only
// the public Sheet accessors documented in `src/sheet.h`.

/// Emits `<sheetViews><sheetView>...</sheetView></sheetViews>` for the
/// sheet's view state. Returns an empty string when every field is at
/// its default (zoom 100, no freeze panes, tab not hidden) — Excel
/// itself omits the element in that case.
std::string BuildSheetViewXml(const SheetView& view) {
  const bool zoom_default = view.zoom_scale == SheetView::kDefaultZoomScale;
  const bool no_freeze = view.freeze_rows == 0U && view.freeze_cols == 0U;
  const bool tab_default = !view.tab_hidden;
  // Display attributes at their schema defaults contribute nothing.
  const bool display_default = view.show_grid_lines && view.show_row_col_headers && view.show_zeros &&
                               !view.right_to_left && !view.tab_selected && view.view_mode.empty();
  if (zoom_default && no_freeze && tab_default && display_default) {
    return std::string();
  }
  std::string out;
  out.reserve(224);
  out.append("<sheetViews><sheetView");
  // Display attributes precede workbookViewId in Excel's own emission
  // order. Emit each only when it differs from its schema default so a
  // near-default sheet stays compact.
  if (!view.show_grid_lines) {
    out.append(" showGridLines=\"0\"");
  }
  if (!view.show_row_col_headers) {
    out.append(" showRowColHeaders=\"0\"");
  }
  if (!view.show_zeros) {
    out.append(" showZeros=\"0\"");
  }
  if (view.right_to_left) {
    out.append(" rightToLeft=\"1\"");
  }
  if (view.tab_selected) {
    out.append(" tabSelected=\"1\"");
  }
  if (!view.view_mode.empty()) {
    out.append(" view=\"");
    AppendXmlAttrEscaped(out, view.view_mode);
    out.push_back('"');
  }
  out.append(" workbookViewId=\"0\"");
  if (!zoom_default) {
    out.append(" zoomScale=\"");
    out.append(std::to_string(view.zoom_scale));
    out.push_back('"');
  }
  if (no_freeze) {
    out.append("/></sheetViews>");
    return out;
  }
  out.push_back('>');
  // OOXML attribute order for `<pane>`: xSplit, ySplit, topLeftCell,
  // activePane, state. We emit only the fields we own; the writer
  // does not yet model `topLeftCell` or `activePane`, so they are
  // absent. Excel gracefully accepts a freeze record without them.
  out.append("<pane");
  if (view.freeze_cols != 0U) {
    out.append(" xSplit=\"");
    out.append(std::to_string(view.freeze_cols));
    out.push_back('"');
  }
  if (view.freeze_rows != 0U) {
    out.append(" ySplit=\"");
    out.append(std::to_string(view.freeze_rows));
    out.push_back('"');
  }
  out.append(" state=\"frozen\"/></sheetView></sheetViews>");
  return out;
}

/// Emits the modelled subset of `<sheetFormatPr>`.  Unlike explicit `<col>`
/// and `<row>` records, these metrics apply to every un-overridden track, so
/// silently omitting them changes pagination and visible geometry across the
/// whole worksheet.
std::string BuildSheetFormatPrXml(const SheetFormatDefaults& defaults) {
  if (!defaults.has_default_col_width && !defaults.has_default_row_height &&
      defaults.base_col_width == ooxml_defaults::kBaseColWidthChars) {
    return std::string();
  }
  auto append_double = [](std::string& out, double value) {
    char buf[64];
    std::snprintf(buf, sizeof(buf), "%.17g", value);
    out.append(buf);
  };
  std::string out("<sheetFormatPr");
  if (defaults.base_col_width != ooxml_defaults::kBaseColWidthChars) {
    out.append(" baseColWidth=\"");
    append_double(out, defaults.base_col_width);
    out.push_back('"');
  }
  if (defaults.has_default_col_width) {
    out.append(" defaultColWidth=\"");
    append_double(out, defaults.default_col_width);
    out.push_back('"');
  }
  if (defaults.has_default_row_height) {
    out.append(" defaultRowHeight=\"");
    append_double(out, defaults.default_row_height);
    out.push_back('"');
  }
  out.append("/>");
  return out;
}

/// Emits `<sheetProtection .../>` matching the structure ECMA-376
/// §18.3.1.85 prescribes. Returns an empty string when
/// `p.enabled == false` so the caller can drop the surrounding
/// indentation cleanly. Boolean attributes are emitted only when they
/// differ from their per-attribute schema default (eleven action flags
/// default to true, the rest to false), which keeps output compact and
/// preserves explicit unlocks of otherwise-locked-by-default actions.
std::string BuildSheetProtectionXml(const SheetProtection& p) {
  if (!p.enabled) {
    return std::string();
  }
  std::string out;
  out.reserve(256);
  out.append("<sheetProtection");
  if (!p.algorithm_name.empty()) {
    out.append(" algorithmName=\"");
    AppendXmlAttrEscaped(out, p.algorithm_name);
    out.push_back('"');
  }
  if (!p.hash_value.empty()) {
    out.append(" hashValue=\"");
    AppendXmlAttrEscaped(out, p.hash_value);
    out.push_back('"');
  }
  if (!p.salt_value.empty()) {
    out.append(" saltValue=\"");
    AppendXmlAttrEscaped(out, p.salt_value);
    out.push_back('"');
  }
  if (p.spin_count != 0U) {
    out.append(" spinCount=\"");
    out.append(std::to_string(p.spin_count));
    out.push_back('"');
  }
  if (!p.legacy_password.empty()) {
    out.append(" password=\"");
    AppendXmlAttrEscaped(out, p.legacy_password);
    out.push_back('"');
  }
  // Boolean attributes — emit only when the value differs from the
  // attribute's ECMA-376 §18.3.1.85 schema default. Eleven action flags
  // default to true (locked); the rest default to false. Emitting only
  // deltas both matches Excel's own compact output and — crucially —
  // preserves an explicit unlock (`formatCells="0"`) that a blanket
  // "emit when true" writer would drop, silently re-locking the action on
  // the next load. Order mirrors Excel's emission order.
  const auto append_bool = [&out](const char* name, bool v, bool default_value) {
    if (v == default_value) {
      return;
    }
    out.push_back(' ');
    out.append(name);
    out.append(v ? "=\"1\"" : "=\"0\"");
  };
  append_bool("sheet", p.sheet, false);
  append_bool("objects", p.objects, false);
  append_bool("scenarios", p.scenarios, false);
  append_bool("formatCells", p.format_cells, true);
  append_bool("formatColumns", p.format_columns, true);
  append_bool("formatRows", p.format_rows, true);
  append_bool("insertColumns", p.insert_columns, true);
  append_bool("insertRows", p.insert_rows, true);
  append_bool("insertHyperlinks", p.insert_hyperlinks, true);
  append_bool("deleteColumns", p.delete_columns, true);
  append_bool("deleteRows", p.delete_rows, true);
  append_bool("selectLockedCells", p.select_locked_cells, false);
  append_bool("sort", p.sort, true);
  append_bool("autoFilter", p.auto_filter, true);
  append_bool("pivotTables", p.pivot_tables, true);
  append_bool("selectUnlockedCells", p.select_unlocked_cells, false);
  out.append("/>");
  return out;
}

/// Emits `<cols>...</cols>` containing one `<col>` per
/// `ColumnLayout` entry. Returns an empty string when there are no
/// column layout overrides.
std::string BuildColsXml(const SheetLayout& layout) {
  if (layout.columns.empty()) {
    return std::string();
  }
  std::string out;
  out.reserve(32U + layout.columns.size() * 96U);
  out.append("<cols>");
  for (const ColumnLayout& col : layout.columns) {
    out.append("<col min=\"");
    out.append(std::to_string(col.first + 1U));
    out.append("\" max=\"");
    out.append(std::to_string(col.last + 1U));
    out.push_back('"');
    // Keep the legacy convenience of treating a non-zero programmatic
    // width as logically explicit, while preserving an explicit width="0"
    // and distinguishing it from an absent width.
    const bool has_width = HasExplicitColumnWidth(col);
    if (has_width) {
      out.append(" width=\"");
      char buf[32];
      // %.17g is round-trip safe under IEEE 754, so a recalc-save does
      // not drift the column metric. Matches the row-height writer.
      std::snprintf(buf, sizeof(buf), "%.17g", col.width);
      out.append(buf);
      // Excel emits `customWidth="1"` whenever an explicit `width` is
      // present so a reload preserves the column metric.
      out.append("\" customWidth=\"1\"");
    }
    if (col.has_style) {
      out.append(" style=\"");
      out.append(std::to_string(col.style_xf));
      out.push_back('\"');
    }
    if (col.hidden) {
      out.append(" hidden=\"1\"");
    }
    if (col.outline_level != 0U) {
      out.append(" outlineLevel=\"");
      out.append(std::to_string(static_cast<unsigned int>(col.outline_level)));
      out.push_back('"');
    }
    out.append("/>");
  }
  out.append("</cols>");
  return out;
}

/// Emits `<dimension ref="A1:..."/>` for the sheet's populated bounding
/// box. ECMA-376 places `<dimension>` between `<sheetPr>` and
/// `<sheetViews>`; some readers (and Excel's Name Box) use it to seed the
/// used range, so an accurate box avoids a divergent used-range guess.
///
/// A cell counts as populated when it carries a formula or a non-blank
/// cached value, mirroring the pagination engine's used-range walk. An
/// empty sheet emits `<dimension ref="A1"/>` (Excel's convention for a
/// sheet with no content).
std::string BuildDimensionXml(const Sheet& sheet) {
  bool any = false;
  std::uint32_t min_row = 0;
  std::uint32_t min_col = 0;
  std::uint32_t max_row = 0;
  std::uint32_t max_col = 0;
  for (const auto& [row_index, cells] : sheet.rows()) {
    for (std::size_t c = 0; c < cells.size(); ++c) {
      const Cell& cell = cells[c];
      if (cell.formula_text.empty() && cell.cached_value.is_blank()) {
        continue;
      }
      const auto col_index = static_cast<std::uint32_t>(c);
      if (!any) {
        min_row = max_row = row_index;
        min_col = max_col = col_index;
        any = true;
        continue;
      }
      min_row = std::min(min_row, row_index);
      max_row = std::max(max_row, row_index);
      min_col = std::min(min_col, col_index);
      max_col = std::max(max_col, col_index);
    }
  }
  std::string out;
  out.append("<dimension ref=\"");
  if (!any) {
    out.append("A1");
  } else {
    AppendCellRefForRef(out, min_row, min_col);
    if (min_row != max_row || min_col != max_col) {
      out.push_back(':');
      AppendCellRefForRef(out, max_row, max_col);
    }
  }
  out.append("\"/>");
  return out;
}

}  // namespace

std::string BuildWorksheetXml(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                              const std::vector<std::string>& table_rids,
                              const std::vector<std::string>& hyperlink_rids, std::string_view printer_settings_rid,
                              std::string_view drawing_rid, std::string_view legacy_drawing_rid,
                              const SharedStrings* shared_strings, std::size_t dxf_count) {
  const std::string sheet_view_xml = BuildSheetViewXml(sheet.view());
  const std::string sheet_format_xml = BuildSheetFormatPrXml(sheet.format_defaults());
  const std::string cols_xml = BuildColsXml(sheet.layout());
  const std::string sheet_data = BuildSheetDataXml(sheet, shared_strings);
  // Conditional-format blocks live between <sheetData> and <tableParts>
  // in ECMA-376 document order. Empty list => empty string, no
  // wrapper.
  const std::string cf_xml = write_conditional_formattings(sheet.conditional_formats(), dxf_count);
  // Data-bar settings with no legacy attribute live in the worksheet
  // `<extLst>`, which is emitted much further down; build them here so
  // the CF model is read once.
  const std::string ext_lst_xml =
      merge_x14_cf_entries(sheet.ext_lst_xml(), build_x14_cf_overlay_entries(sheet.conditional_formats()));
  const std::string merges_xml = BuildMergeCellsBlock(sheet);
  const std::string dv_xml = BuildDataValidationsBlock(sheet);
  const std::string hl_xml = BuildHyperlinksBlock(sheet, hyperlink_rids);
  const SheetPrintSettings& print = sheet.print_settings();
  const WorksheetRawExtensions& raw_extensions = sheet.raw_extensions();
  const std::string page_setup_xml = PageSetupWithRelationshipId(print.page_setup_xml, printer_settings_rid);
  const std::string sheet_pr_xml = BuildNormalizedSheetPrXml(print.sheet_pr_xml);
  std::string out;
  out.reserve(192U + sheet_view_xml.size() + sheet_format_xml.size() + cols_xml.size() + sheet_data.size() +
              cf_xml.size() + merges_xml.size() + dv_xml.size() + hl_xml.size() + sheet_pr_xml.size() +
              print.page_margins_xml.size() + page_setup_xml.size() + sheet_tables.size() * 96);
  out.append(kXmlDecl);
  out.append(
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
  // Re-declare the source worksheet root's extra namespaces so prefixed
  // attributes carried in any raw capture below resolve (keeps the output
  // well-formed; mirrors the workbook-root handling).
  out.append(sheet.root_extra_ns_attrs());
  out.append(">\n");
  // ECMA-376 element order: sheetPr -> dimension -> sheetViews ->
  // sheetFormatPr -> cols -> sheetData -> conditionalFormatting ->
  // pageMargins -> pageSetup -> rowBreaks -> colBreaks -> tableParts.
  // We currently emit a subset; the helpers stay quiet when their
  // underlying field is at default values so absent metadata yields no
  // extra bytes.
  if (!sheet_pr_xml.empty()) {
    out.append("  ");
    out.append(sheet_pr_xml);
    out.push_back('\n');
  }
  out.append("  ");
  out.append(BuildDimensionXml(sheet));
  out.push_back('\n');
  if (!sheet_view_xml.empty()) {
    out.append("  ");
    out.append(sheet_view_xml);
    out.push_back('\n');
  }
  if (!sheet_format_xml.empty()) {
    out.append("  ");
    out.append(sheet_format_xml);
    out.push_back('\n');
  }
  if (!cols_xml.empty()) {
    out.append("  ");
    out.append(cols_xml);
    out.push_back('\n');
  }
  out.append("  ");
  out.append(sheet_data);
  out.push_back('\n');
  // <sheetProtection> sits between <sheetData> and <mergeCells> per
  // ECMA-376 document order. Helper returns "" when protection is
  // disabled, leaving no trailing whitespace in that case.
  {
    const std::string sp_xml = BuildSheetProtectionXml(sheet.protection());
    if (!sp_xml.empty()) {
      out.append("  ");
      out.append(sp_xml);
      out.push_back('\n');
    }
  }
  if (!raw_extensions.protected_ranges_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.protected_ranges_xml);
    out.push_back('\n');
  }
  if (!raw_extensions.scenarios_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.scenarios_xml);
    out.push_back('\n');
  }
  // <autoFilter> sits between <sheetProtection>/<scenarios> and
  // <mergeCells> in ECMA-376 document order. Round-tripped verbatim.
  if (!sheet.auto_filter_xml().empty()) {
    out.append("  ");
    out.append(sheet.auto_filter_xml());
    out.push_back('\n');
  }
  if (!raw_extensions.custom_sheet_views_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.custom_sheet_views_xml);
    out.push_back('\n');
  }
  // Merge cells precede CF in ECMA-376 document order.
  if (!merges_xml.empty()) {
    out.append("  ");
    out.append(merges_xml);
    out.push_back('\n');
  }
  if (!raw_extensions.phonetic_pr_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.phonetic_pr_xml);
    out.push_back('\n');
  }
  if (!cf_xml.empty()) {
    out.append("  ");
    out.append(cf_xml);
    out.push_back('\n');
  }
  if (!dv_xml.empty()) {
    out.append("  ");
    out.append(dv_xml);
    out.push_back('\n');
  }
  if (!hl_xml.empty()) {
    out.append("  ");
    out.append(hl_xml);
    out.push_back('\n');
  }
  // <printOptions> sits between <hyperlinks> and <pageMargins>.
  if (!print.print_options_xml.empty()) {
    out.append("  ");
    out.append(print.print_options_xml);
    out.push_back('\n');
  }
  if (!print.page_margins_xml.empty()) {
    out.append("  ");
    out.append(print.page_margins_xml);
    out.push_back('\n');
  }
  if (!page_setup_xml.empty()) {
    out.append("  ");
    out.append(page_setup_xml);
    out.push_back('\n');
  }
  // <headerFooter> follows <pageSetup> and precedes <rowBreaks>.
  if (!print.header_footer_xml.empty()) {
    out.append("  ");
    out.append(print.header_footer_xml);
    out.push_back('\n');
  }
  // Manual page breaks. ECMA-376 places <rowBreaks>/<colBreaks> after
  // <pageSetup> and before drawing parts / <tableParts>.
  {
    const std::string row_breaks_xml = BuildPageBreaksXml("rowBreaks", print.manual_row_breaks);
    if (!row_breaks_xml.empty()) {
      out.append("  ");
      out.append(row_breaks_xml);
      out.push_back('\n');
    }
    const std::string col_breaks_xml = BuildPageBreaksXml("colBreaks", print.manual_col_breaks);
    if (!col_breaks_xml.empty()) {
      out.append("  ");
      out.append(col_breaks_xml);
      out.push_back('\n');
    }
  }
  if (!raw_extensions.ignored_errors_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.ignored_errors_xml);
    out.push_back('\n');
  }
  // <drawing> precedes <tableParts> in ECMA-376 worksheet element order.
  // The referenced DrawingML part round-trips through passthrough; here
  // we only re-emit the reference so the part stays reachable.
  if (!drawing_rid.empty()) {
    out.append("  <drawing r:id=\"");
    out.append(drawing_rid);
    out.append("\"/>\n");
  }
  if (!legacy_drawing_rid.empty()) {
    out.append("  <legacyDrawing r:id=\"");
    out.append(legacy_drawing_rid);
    out.append("\"/>\n");
  }
  if (!raw_extensions.legacy_drawing_hf_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.legacy_drawing_hf_xml);
    out.push_back('\n');
  }
  if (!raw_extensions.picture_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.picture_xml);
    out.push_back('\n');
  }
  if (!raw_extensions.ole_objects_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.ole_objects_xml);
    out.push_back('\n');
  }
  if (!raw_extensions.controls_xml.empty()) {
    out.append("  ");
    out.append(raw_extensions.controls_xml);
    out.push_back('\n');
  }
  if (!sheet_tables.empty()) {
    out.append("  <tableParts count=\"");
    out.append(std::to_string(sheet_tables.size()));
    out.append("\">\n");
    for (std::size_t i = 0; i < sheet_tables.size(); ++i) {
      out.append("    <tablePart r:id=\"");
      // `table_rids` is index-aligned with `sheet_tables`: it names the
      // id `BuildSheetRels` actually assigned to this table's
      // relationship, which may not be `rId(i+1)` when the sheet also
      // carries unknown relationships occupying lower-numbered ids.
      if (i < table_rids.size()) {
        out.append(table_rids[i]);
      }
      out.append("\"/>\n");
    }
    out.append("  </tableParts>\n");
  }
  // Worksheet-level `<extLst>` is the last child of `<worksheet>` in
  // ECMA-376 order (after `<tableParts>`). Re-emit the captured block
  // verbatim so 2010+ extension data (x14 conditional formatting, etc.)
  // survives the round trip.
  if (!ext_lst_xml.empty()) {
    out.append("  ");
    out.append(ext_lst_xml);
    out.push_back('\n');
  }
  out.append("</worksheet>\n");
  return out;
}

SheetRelsResult BuildSheetRels(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                               const std::vector<EmissionPlan::PivotTablePlan>& sheet_pivot_tables,
                               const EmissionPlan::CommentsPlan& comments_plan, const EmissionPlan& plan) {
  SheetRelsResult res;
  std::string& out = res.xml;
  out.reserve(256 + (sheet_tables.size() + sheet_pivot_tables.size() + sheet.hyperlinks().size()) * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  std::uint32_t next_rid = 1;
  std::unordered_set<std::string> used_rids;
  // Preserve unmodelled relationship ids exactly: raw worksheet XML can
  // refer to them (for example `<oleObject r:id="rId7"/>`). Reserve them
  // before minting ids for modelled features to avoid a collision.
  for (const UnknownRelationship& relationship : sheet.unknown_relationships()) {
    if (!relationship.id.empty()) {
      used_rids.insert(relationship.id);
    }
  }
  auto next_unique_rid = [&]() {
    std::string rid;
    do {
      rid = "rId" + std::to_string(next_rid++);
    } while (used_rids.count(rid) != 0U);
    used_rids.insert(rid);
    return rid;
  };
  // Every relationship this function writes goes through `add_rel` so
  // `res.relationship_count` always reflects exactly the `<Relationship>`
  // elements present in `res.xml`, whichever `AppendRelationship`
  // overload the call site uses.
  auto add_rel = [&](auto&&... args) {
    AppendRelationship(out, std::forward<decltype(args)>(args)...);
    ++res.relationship_count;
  };
  res.table_rids.reserve(sheet_tables.size());
  for (std::size_t i = 0; i < sheet_tables.size(); ++i) {
    const std::string target = "../tables/table" + std::to_string(sheet_tables[i].numeric_id) + ".xml";
    const std::string rid = next_unique_rid();
    add_rel(rid, kRelTable, target);
    res.table_rids.push_back(rid);
  }
  // Pivot-table relationships follow the table relationships, with rId
  // numbering continuing in sequence so each rel id is unique within
  // the sheet rels file.
  for (std::size_t i = 0; i < sheet_pivot_tables.size(); ++i) {
    const std::string target = "../pivotTables/pivotTable" + std::to_string(sheet_pivot_tables[i].numeric_id) + ".xml";
    add_rel(next_unique_rid(), kRelPivotTable, target);
  }
  // Hyperlink relationships. Each external hyperlink target gets one
  // entry. Relative ordering is preserved so the round-trip writes the
  // rIds in the same order the reader observed.
  res.hyperlink_rids.reserve(sheet.hyperlinks().size());
  std::unordered_map<std::string, std::string> hyperlink_targets;
  hyperlink_targets.reserve(sheet.hyperlinks().size());
  for (const Hyperlink& h : sheet.hyperlinks()) {
    if (h.target.empty()) {
      // Pure internal links carry their target through the inline
      // `location=` attribute instead of a rels entry.
      res.hyperlink_rids.emplace_back();
      continue;
    }
    // Reuse the source `rid` when present so byte-identical round-trips
    // are possible; otherwise mint a fresh rIdN counter.
    std::string rid;
    if (!h.rid.empty()) {
      const auto assigned = hyperlink_targets.find(h.rid);
      if (assigned != hyperlink_targets.end() && assigned->second == h.target) {
        // Multiple hyperlink elements may share one relationship when they
        // point at the same external target. Preserve that source sharing.
        res.hyperlink_rids.push_back(h.rid);
        continue;
      }
      if (used_rids.count(h.rid) == 0U) {
        rid = h.rid;
        used_rids.insert(rid);
      }
    }
    if (rid.empty()) {
      rid = next_unique_rid();
    }
    hyperlink_targets.emplace(rid, h.target);
    res.hyperlink_rids.push_back(rid);
    add_rel(rid, kRelHyperlink, h.target, /*target_external=*/true, /*escape_target=*/true);
  }
  const SheetPrintSettings& print = sheet.print_settings();
  if (!print.printer_settings_path.empty()) {
    if (!print.printer_settings_rid.empty() && used_rids.count(print.printer_settings_rid) == 0U) {
      res.printer_settings_rid = print.printer_settings_rid;
      used_rids.insert(res.printer_settings_rid);
    } else {
      res.printer_settings_rid = next_unique_rid();
    }
    add_rel(res.printer_settings_rid, kRelPrinterSettings, TargetRelativeToWorksheet(print.printer_settings_path));
  }
  // Drawing (DrawingML) relationship. The part body, its own rels, and
  // any anchored media round-trip through passthrough; here we re-mint a
  // fresh rId so the worksheet's `<drawing r:id>` element resolves.
  if (!sheet.drawing_rel_target().empty()) {
    res.drawing_rid = next_unique_rid();
    add_rel(res.drawing_rid, kRelDrawing, TargetRelativeToWorksheet(sheet.drawing_rel_target()));
  }
  // Comments + VML relationships when the sheet has comments. The
  // comments rel comes first; the VML rel follows so the two ids are
  // adjacent and readers that scan top-down see them as a pair.
  if (comments_plan.numeric_id != 0) {
    const std::string comments_target = "../comments" + std::to_string(comments_plan.numeric_id) + ".xml";
    const std::string vml_target = "../drawings/vmlDrawing" + std::to_string(comments_plan.numeric_id) + ".vml";
    add_rel(next_unique_rid(), kRelComments, comments_target);
    res.legacy_drawing_rid = next_unique_rid();
    add_rel(res.legacy_drawing_rid, kRelVmlDrawing, vml_target);
  }
  for (const UnknownRelationship& relationship : sheet.unknown_relationships()) {
    // Internal unknown relationships are meaningful only when the
    // corresponding passthrough payload survived collision handling. A
    // generated part at the same path is not a substitute: it may have a
    // different type or schema from the source edge's target.
    if (relationship.id.empty() || (!relationship.target_external &&
                                    (relationship.target.empty() || !HasPassthroughPart(plan, relationship.target)))) {
      continue;
    }
    const std::string target =
        relationship.target_external ? relationship.target : TargetRelativeToWorksheet(relationship.target);
    add_rel(relationship.id, relationship.type, target, relationship.target_external, /*escape_target=*/true);
  }
  out.append("</Relationships>\n");
  return res;
}

}  // namespace io
}  // namespace formulon
