// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include <unordered_set>
#include <vector>

#include "io/cf_writer.h"
#include "io/ooxml/cell_ref_writer.h"
#include "io/ooxml/emission_plan.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_writer_cell.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "sheet.h"

namespace formulon {
namespace io {
namespace {

// Forward declarations for sheet-view / column-layout builders. The
// definitions live at the bottom of the file alongside the other XML
// helpers; their signatures are needed up here so `BuildWorksheetXml`
// can call them.
std::string BuildSheetViewXml(const SheetView& view);
std::string BuildColsXml(const SheetLayout& layout);
std::string BuildSheetProtectionXml(const SheetProtection& p);

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
    if (!v.error_title.empty()) {
      out.append(" errorTitle=\"");
      AppendXmlEscaped(out, v.error_title);
      out.append("\"");
    }
    if (!v.error_message.empty()) {
      out.append(" error=\"");
      AppendXmlEscaped(out, v.error_message);
      out.append("\"");
    }
    if (!v.prompt_title.empty()) {
      out.append(" promptTitle=\"");
      AppendXmlEscaped(out, v.prompt_title);
      out.append("\"");
    }
    if (!v.prompt_message.empty()) {
      out.append(" prompt=\"");
      AppendXmlEscaped(out, v.prompt_message);
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
    out.append("\"");
    if (i < rid_per_hyperlink.size() && !rid_per_hyperlink[i].empty()) {
      out.append(" r:id=\"");
      AppendXmlEscaped(out, rid_per_hyperlink[i]);
      out.append("\"");
    }
    if (!h.location.empty()) {
      out.append(" location=\"");
      AppendXmlEscaped(out, h.location);
      out.append("\"");
    }
    if (!h.tooltip.empty()) {
      out.append(" tooltip=\"");
      AppendXmlEscaped(out, h.tooltip);
      out.append("\"");
    }
    if (!h.display.empty()) {
      out.append(" display=\"");
      AppendXmlEscaped(out, h.display);
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
  if (zoom_default && no_freeze && tab_default) {
    return std::string();
  }
  std::string out;
  out.reserve(192);
  out.append("<sheetViews><sheetView workbookViewId=\"0\"");
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

/// Emits `<sheetProtection .../>` matching the structure ECMA-376
/// §18.3.1.85 prescribes. Returns an empty string when
/// `p.enabled == false` so the caller can drop the surrounding
/// indentation cleanly. Boolean attributes are emitted only when their
/// value is `true`; the spec defaults absent attributes to `false`.
std::string BuildSheetProtectionXml(const SheetProtection& p) {
  if (!p.enabled) {
    return std::string();
  }
  std::string out;
  out.reserve(256);
  out.append("<sheetProtection");
  if (!p.algorithm_name.empty()) {
    out.append(" algorithmName=\"");
    AppendXmlEscaped(out, p.algorithm_name);
    out.push_back('"');
  }
  if (!p.hash_value.empty()) {
    out.append(" hashValue=\"");
    AppendXmlEscaped(out, p.hash_value);
    out.push_back('"');
  }
  if (!p.salt_value.empty()) {
    out.append(" saltValue=\"");
    AppendXmlEscaped(out, p.salt_value);
    out.push_back('"');
  }
  if (p.spin_count != 0U) {
    out.append(" spinCount=\"");
    out.append(std::to_string(p.spin_count));
    out.push_back('"');
  }
  if (!p.legacy_password.empty()) {
    out.append(" password=\"");
    AppendXmlEscaped(out, p.legacy_password);
    out.push_back('"');
  }
  // Boolean attributes — only emit when `true`. Order mirrors Excel's
  // own emission order so byte-identical round-trips are achievable
  // for the common cases.
  const auto append_bool = [&out](const char* name, bool v) {
    if (!v) {
      return;
    }
    out.push_back(' ');
    out.append(name);
    out.append("=\"1\"");
  };
  append_bool("sheet", p.sheet);
  append_bool("objects", p.objects);
  append_bool("scenarios", p.scenarios);
  append_bool("formatCells", p.format_cells);
  append_bool("formatColumns", p.format_columns);
  append_bool("formatRows", p.format_rows);
  append_bool("insertColumns", p.insert_columns);
  append_bool("insertRows", p.insert_rows);
  append_bool("insertHyperlinks", p.insert_hyperlinks);
  append_bool("deleteColumns", p.delete_columns);
  append_bool("deleteRows", p.delete_rows);
  append_bool("selectLockedCells", p.select_locked_cells);
  append_bool("sort", p.sort);
  append_bool("autoFilter", p.auto_filter);
  append_bool("pivotTables", p.pivot_tables);
  append_bool("selectUnlockedCells", p.select_unlocked_cells);
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
    if (col.width > 0.0) {
      out.append(" width=\"");
      char buf[32];
      std::snprintf(buf, sizeof(buf), "%.6g", col.width);
      out.append(buf);
      // Excel emits `customWidth="1"` whenever an explicit `width` is
      // present so a reload preserves the column metric.
      out.append("\" customWidth=\"1\"");
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

}  // namespace

std::string BuildWorksheetXml(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                              const std::vector<std::string>& hyperlink_rids, std::string_view printer_settings_rid) {
  const std::string sheet_view_xml = BuildSheetViewXml(sheet.view());
  const std::string cols_xml = BuildColsXml(sheet.layout());
  const std::string sheet_data = BuildSheetDataXml(sheet);
  // Conditional-format blocks live between <sheetData> and <tableParts>
  // in ECMA-376 document order. Empty list => empty string, no
  // wrapper.
  const std::string cf_xml = write_conditional_formattings(sheet.conditional_formats());
  const std::string merges_xml = BuildMergeCellsBlock(sheet);
  const std::string dv_xml = BuildDataValidationsBlock(sheet);
  const std::string hl_xml = BuildHyperlinksBlock(sheet, hyperlink_rids);
  const SheetPrintSettings& print = sheet.print_settings();
  const std::string page_setup_xml = PageSetupWithRelationshipId(print.page_setup_xml, printer_settings_rid);
  std::string out;
  out.reserve(192U + sheet_view_xml.size() + cols_xml.size() + sheet_data.size() + cf_xml.size() + merges_xml.size() +
              dv_xml.size() + hl_xml.size() + print.sheet_pr_xml.size() + print.page_margins_xml.size() +
              page_setup_xml.size() + sheet_tables.size() * 96);
  out.append(kXmlDecl);
  out.append(
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");
  // ECMA-376 element order: sheetPr -> dimension -> sheetViews ->
  // sheetFormatPr -> cols -> sheetData -> conditionalFormatting ->
  // pageMargins -> pageSetup -> rowBreaks -> colBreaks -> tableParts.
  // We currently emit a subset; the helpers stay quiet when their
  // underlying field is at default values so absent metadata yields no
  // extra bytes.
  if (!print.sheet_pr_xml.empty()) {
    out.append("  ");
    out.append(print.sheet_pr_xml);
    out.push_back('\n');
  }
  if (!sheet_view_xml.empty()) {
    out.append("  ");
    out.append(sheet_view_xml);
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
  // Merge cells precede CF in ECMA-376 document order.
  if (!merges_xml.empty()) {
    out.append("  ");
    out.append(merges_xml);
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
  if (!sheet_tables.empty()) {
    out.append("  <tableParts count=\"");
    out.append(std::to_string(sheet_tables.size()));
    out.append("\">\n");
    for (std::size_t i = 0; i < sheet_tables.size(); ++i) {
      out.append("    <tablePart r:id=\"rId");
      out.append(std::to_string(i + 1));
      out.append("\"/>\n");
    }
    out.append("  </tableParts>\n");
  }
  out.append("</worksheet>\n");
  return out;
}

SheetRelsResult BuildSheetRels(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                               const std::vector<EmissionPlan::PivotTablePlan>& sheet_pivot_tables,
                               const EmissionPlan::CommentsPlan& comments_plan) {
  SheetRelsResult res;
  std::string& out = res.xml;
  out.reserve(256 + (sheet_tables.size() + sheet_pivot_tables.size() + sheet.hyperlinks().size()) * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  std::uint32_t next_rid = 1;
  std::unordered_set<std::string> used_rids;
  auto next_unique_rid = [&]() {
    std::string rid;
    do {
      rid = "rId" + std::to_string(next_rid++);
    } while (used_rids.count(rid) != 0U);
    used_rids.insert(rid);
    return rid;
  };
  for (std::size_t i = 0; i < sheet_tables.size(); ++i) {
    const std::string target = "../tables/table" + std::to_string(sheet_tables[i].numeric_id) + ".xml";
    AppendRelationship(out, next_unique_rid(), kRelTable, target);
  }
  // Pivot-table relationships follow the table relationships, with rId
  // numbering continuing in sequence so each rel id is unique within
  // the sheet rels file.
  for (std::size_t i = 0; i < sheet_pivot_tables.size(); ++i) {
    const std::string target = "../pivotTables/pivotTable" + std::to_string(sheet_pivot_tables[i].numeric_id) + ".xml";
    AppendRelationship(out, next_unique_rid(), kRelPivotTable, target);
  }
  // Hyperlink relationships. Each external hyperlink target gets one
  // entry. Relative ordering is preserved so the round-trip writes the
  // rIds in the same order the reader observed.
  res.hyperlink_rids.reserve(sheet.hyperlinks().size());
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
    if (!h.rid.empty() && used_rids.count(h.rid) == 0U) {
      rid = h.rid;
      used_rids.insert(rid);
    } else {
      rid = next_unique_rid();
    }
    res.hyperlink_rids.push_back(rid);
    AppendRelationship(out, rid, kRelHyperlink, h.target, /*target_external=*/true, /*escape_target=*/true);
  }
  const SheetPrintSettings& print = sheet.print_settings();
  if (!print.printer_settings_path.empty()) {
    if (!print.printer_settings_rid.empty() && used_rids.count(print.printer_settings_rid) == 0U) {
      res.printer_settings_rid = print.printer_settings_rid;
      used_rids.insert(res.printer_settings_rid);
    } else {
      res.printer_settings_rid = next_unique_rid();
    }
    AppendRelationship(out, res.printer_settings_rid, kRelPrinterSettings,
                       TargetRelativeToWorksheet(print.printer_settings_path));
  }
  // Comments + VML relationships when the sheet has comments. The
  // comments rel comes first; the VML rel follows so the two ids are
  // adjacent and readers that scan top-down see them as a pair.
  if (comments_plan.numeric_id != 0) {
    const std::string comments_target = "../comments" + std::to_string(comments_plan.numeric_id) + ".xml";
    const std::string vml_target = "../drawings/vmlDrawing" + std::to_string(comments_plan.numeric_id) + ".vml";
    AppendRelationship(out, next_unique_rid(), kRelComments, comments_target);
    AppendRelationship(out, next_unique_rid(), kRelVmlDrawing, vml_target);
  }
  out.append("</Relationships>\n");
  return res;
}

}  // namespace io
}  // namespace formulon
