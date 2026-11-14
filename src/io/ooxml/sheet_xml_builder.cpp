//
// Worksheet-part XML builders for the OOXML writer. See header for the
// caller contract; this TU owns all of the per-sheet body builders that
// were previously parked alongside the orchestrator in
// `src/io/ooxml_writer.cpp`.

#include "io/ooxml/sheet_xml_builder.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell_style_save_plan.h"
#include "hyperlink_address.h"
#include "io/auto_filter_xml.h"
#include "io/cf_overlay.h"
#include "io/cf_writer.h"
#include "io/dv_overlay.h"
#include "io/ooxml/cell_ref_writer.h"
#include "io/ooxml/emission_plan.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/shared_strings_writer.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_writer_cell.h"
#include "io/protection_xml.h"
#include "io/sheet_phonetic.h"
#include "io/sheet_view_xml.h"
#include "io/sqref_text.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/a1_ref.h"
#include "utils/double_format.h"

namespace formulon {
namespace io {
namespace {

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
  out.append("<mergeCells");
  append_xml_attr_uint(out, "count", sheet.merges().size());
  out.append(">");
  for (const MergeRange& m : sheet.merges()) {
    out.append("<mergeCell ref=\"");
    AppendRangeRef(out, m);
    out.append("\"/>");
  }
  out.append("</mergeCells>");
  return out;
}

std::string BuildDataValidationsBlock(const Sheet& sheet, const parser::ExternalBookIndexer* indexer,
                                      const std::vector<std::string>* sheet_names) {
  // Validations stored in the x14 extension are written by `merge_x14_dv_entries`.
  std::vector<const DataValidation*> main_dvs;
  for (const DataValidation& v : sheet.validations()) {
    if (!data_validation_in_x14(v, sheet.name(), sheet_names)) {
      main_dvs.push_back(&v);
    }
  }
  if (main_dvs.empty()) {
    return {};
  }
  std::string out;
  out.reserve(96 + main_dvs.size() * 96);
  out.append("<dataValidations");
  append_xml_attr_uint(out, "count", main_dvs.size());
  out.append(sheet.validations_extra_attrs());
  out.append(">");
  const auto append_formula = [&](const char* tag, const std::string& formula) {
    if (formula.empty()) {
      return;
    }
    out.append("<").append(tag).append(">");
    AppendXmlEscaped(out, storage_feature_formula(formula, indexer));
    out.append("</").append(tag).append(">");
  };
  for (const DataValidation* dv : main_dvs) {
    const DataValidation& v = *dv;
    out.append("<dataValidation");
    append_data_validation_attrs(out, v);
    out.append(" sqref=\"");
    out.append(format_sqref_text(v.ranges));
    out.append("\">");
    append_formula("formula1", v.formula1);
    append_formula("formula2", v.formula2);
    out.append("</dataValidation>");
  }
  out.append("</dataValidations>");
  return out;
}

/// Builds the `<hyperlinks>` block. `rid_for_index` returns the rels-file
/// rId integer assigned to the hyperlink at index `i`, or 0 when no rId
/// was assigned (defensive — every external hyperlink gets one). Each
/// `display` is the one `saved_hyperlink` gives for `book_directory`.
std::string BuildHyperlinksBlock(const Sheet& sheet, const std::vector<std::string>& rid_per_hyperlink,
                                 std::string_view book_directory) {
  if (sheet.hyperlinks().empty()) {
    return {};
  }
  std::string out;
  out.reserve(48 + sheet.hyperlinks().size() * 96);
  out.append("<hyperlinks>");
  for (std::size_t i = 0; i < sheet.hyperlinks().size(); ++i) {
    const Hyperlink& h = sheet.hyperlinks()[i];
    out.append("<hyperlink ref=\"");
    AppendRangeRef(out, h.row, h.col, h.last_row, h.last_col);
    out.append("\"");
    if (i < rid_per_hyperlink.size() && !rid_per_hyperlink[i].empty()) {
      append_xml_attr(out, "r:id", rid_per_hyperlink[i]);
    }
    if (!h.location.empty()) {
      append_xml_attr(out, "location", h.location);
    }
    if (!h.tooltip.empty()) {
      append_xml_attr(out, "tooltip", h.tooltip);
    }
    const SavedHyperlink saved = saved_hyperlink(h, book_directory);
    if (!saved.display.empty()) {
      append_xml_attr(out, "display", saved.display);
    }
    out.append("/>");
  }
  out.append("</hyperlinks>");
  return out;
}

/// Completes a `<pageMargins>` element with the attributes ECMA-376
/// §18.3.1.62 makes required.
///
/// All six margins are mandatory, and Excel opens a package whose element
/// omits any of them in repair mode. The model deliberately stores only
/// the margins something actually stated -- that is what lets a caller
/// tell "0.7 because the file says so" from "0.7 because nothing says
/// otherwise" -- so the completion happens here, on the way out, where it
/// costs that distinction nothing. It also covers a source file that
/// arrived incomplete, which no authoring-side fix would reach.
std::string PageMarginsWithRequiredAttributes(std::string page_margins_xml) {
  if (page_margins_xml.empty()) {
    return page_margins_xml;
  }
  struct RequiredMargin {
    std::string_view attr;
    double schema_default;
  };
  static constexpr RequiredMargin kRequired[] = {
      {"left", ooxml_defaults::kPageMarginSideInches},
      {"right", ooxml_defaults::kPageMarginSideInches},
      {"top", ooxml_defaults::kPageMarginTopBottomInches},
      {"bottom", ooxml_defaults::kPageMarginTopBottomInches},
      {"header", ooxml_defaults::kPageMarginHeaderFooterInches},
      {"footer", ooxml_defaults::kPageMarginHeaderFooterInches},
  };
  std::string missing;
  for (const RequiredMargin& required : kRequired) {
    const std::string needle = std::string(" ") + std::string(required.attr) + "=\"";
    if (page_margins_xml.find(needle) != std::string::npos) {
      continue;
    }
    missing.push_back(' ');
    missing.append(required.attr);
    missing.append("=\"");
    formulon::format_double(missing, required.schema_default);
    missing.push_back('"');
  }
  if (missing.empty()) {
    return page_margins_xml;
  }
  const std::size_t self_closing = page_margins_xml.rfind("/>");
  const std::size_t pos = self_closing != std::string::npos ? self_closing : page_margins_xml.rfind('>');
  if (pos == std::string::npos) {
    return page_margins_xml;
  }
  page_margins_xml.insert(pos, missing);
  return page_margins_xml;
}

/// Points `<pageSetup>`'s relationship id at `rid`, or removes the
/// attribute entirely when `rid` is empty.
///
/// An empty `rid` means this save writes no printer-settings
/// relationship for the sheet, so a `r:id` carried over from the source
/// element would name an id the sheet rels file does not declare. That
/// dangles exactly like a relationship with no part, so the attribute
/// goes rather than being left behind.
std::string PageSetupWithRelationshipId(std::string page_setup_xml, std::string_view rid) {
  if (page_setup_xml.empty()) {
    return page_setup_xml;
  }
  auto find_attr = [&page_setup_xml](std::string_view attr_name, std::size_t& value_start, std::size_t& value_end) {
    const std::string needle = std::string(attr_name) + "=\"";
    const std::size_t pos = page_setup_xml.find(needle);
    if (pos == std::string::npos) {
      return std::string::npos;
    }
    value_start = pos + needle.size();
    value_end = page_setup_xml.find('"', value_start);
    return value_end == std::string::npos ? std::string::npos : pos;
  };
  auto patch_attr = [&](std::string_view attr_name) {
    std::size_t value_start = 0;
    std::size_t value_end = 0;
    const std::size_t attr_start = find_attr(attr_name, value_start, value_end);
    if (attr_start == std::string::npos) {
      return false;
    }
    if (rid.empty()) {
      // Take the separating whitespace with the attribute so removing the
      // only one does not leave `<pageSetup />`.
      std::size_t erase_from = attr_start;
      while (erase_from > 0U && page_setup_xml[erase_from - 1U] == ' ') {
        --erase_from;
      }
      page_setup_xml.erase(erase_from, value_end + 1U - erase_from);
      return true;
    }
    page_setup_xml.replace(value_start, value_end - value_start, rid);
    return true;
  };
  if (patch_attr("r:id") || patch_attr("id")) {
    return page_setup_xml;
  }
  if (rid.empty()) {
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
/// adds no bytes. The stored index is emitted as-is: OOXML's `id` is the
/// same 0-based index the break precedes, not a 1-based one (Excel 365
/// writes `id="20"` for a break before row 21).
///
/// `count` is every `<brk>` child; `manualBreakCount` is only those
/// carrying `man="1"`. A sheet can hold both kinds at once - Excel records
/// the automatic breaks it computed alongside the ones the user placed -
/// and the two attributes coincide only when every break is manual.
/// Overstating `manualBreakCount` makes Excel's page-break preview draw
/// solid user-break lines where the file declares computed ones.
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
  std::size_t manual_count = 0;
  for (const ManualBreak& brk : breaks) {
    if (brk.manual) {
      ++manual_count;
    }
  }
  out.push_back('<');
  out.append(element);
  append_xml_attr_uint(out, "count", breaks.size());
  append_xml_attr_uint(out, "manualBreakCount", manual_count);
  out.append(">");
  for (const ManualBreak& brk : breaks) {
    out.append("<brk");
    append_xml_attr_uint(out, "id", brk.id);
    if (brk.min != 0U) {
      append_xml_attr_uint(out, "min", brk.min);
    }
    append_xml_attr_uint(out, "max", brk.max);
    if (brk.manual) {
      emit_xsd_bool_attr(out, "man", true);
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
/// sheet's view state. The element is always written, a default view as
/// `<sheetView workbookViewId="0"/>`, as the desktop application does: on a
/// worksheet without it, Mac Excel opens every custom row height at 4/3 of
/// its stored value.
std::string BuildSheetViewXml(const SheetView& view) {
  std::string out("<sheetViews>");
  out.append(build_sheet_view_element_xml(view, 0U));
  out.append(view.extra_views_xml);
  out.append("</sheetViews>");
  return out;
}

/// Emits the modelled subset of `<sheetFormatPr>`.  Unlike explicit `<col>`
/// and `<row>` records, these metrics apply to every un-overridden track, so
/// silently omitting them changes pagination and visible geometry across the
/// whole worksheet. The outline summaries stay at most 7: Excel opens a row or
/// column at level 8 but refuses a file whose summary says 8.
std::string BuildSheetFormatPrXml(const SheetFormatDefaults& defaults) {
  if (!defaults.has_default_col_width && !defaults.has_default_row_height && !defaults.custom_height &&
      !defaults.zero_height && !defaults.thick_top && !defaults.thick_bottom && defaults.outline_level_row == 0U &&
      defaults.outline_level_col == 0U && defaults.base_col_width == ooxml_defaults::kBaseColWidthChars) {
    return std::string();
  }
  std::string out("<sheetFormatPr");
  if (defaults.base_col_width != ooxml_defaults::kBaseColWidthChars) {
    append_xml_attr_number(out, "baseColWidth", defaults.base_col_width);
  }
  if (defaults.has_default_col_width) {
    append_xml_attr_number(out, "defaultColWidth", defaults.default_col_width);
  }
  // The schema requires defaultRowHeight whenever the element is written.
  append_xml_attr_number(
      out, "defaultRowHeight",
      defaults.has_default_row_height ? defaults.default_row_height : ooxml_defaults::kStandardRowHeightPt);
  if (defaults.custom_height) {
    emit_xsd_bool_attr(out, "customHeight", true);
  }
  if (defaults.zero_height) {
    emit_xsd_bool_attr(out, "zeroHeight", true);
  }
  if (defaults.thick_top) {
    emit_xsd_bool_attr(out, "thickTop", true);
  }
  if (defaults.thick_bottom) {
    emit_xsd_bool_attr(out, "thickBottom", true);
  }
  if (defaults.outline_level_row != 0U) {
    append_xml_attr_uint(out, "outlineLevelRow", std::min(defaults.outline_level_row, kMaxOutlineLevel));
  }
  if (defaults.outline_level_col != 0U) {
    append_xml_attr_uint(out, "outlineLevelCol", std::min(defaults.outline_level_col, kMaxOutlineLevel));
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
  append_password_hash_attrs(out, p.algorithm_name, p.hash_value, p.salt_value, p.spin_count);
  if (!p.legacy_password.empty()) {
    append_xml_attr(out, "password", p.legacy_password);
  }
  // Boolean attributes — emit only when the value differs from the
  // attribute's ECMA-376 §18.3.1.85 schema default. Eleven action flags
  // default to true (locked); the rest default to false. Emitting only
  // deltas both matches Excel's own compact output and — crucially —
  // preserves an explicit unlock (`formatCells="0"`) that a blanket
  // "emit when true" writer would drop, silently re-locking the action on
  // the next load. Order mirrors Excel's emission order.
  for (const SheetProtectionFlag& flag : kSheetProtectionFlags) {
    if (p.*flag.member != flag.default_value) {
      emit_xsd_bool_attr(out, flag.attr, p.*flag.member);
    }
  }
  out.append("/>");
  return out;
}

/// Emits `<cols>...</cols>` containing one `<col>` per
/// `ColumnLayout` entry. Returns an empty string when there are no
/// column layout overrides. `style_plan` renumbers each `style`; null writes
/// model indices.
std::string BuildColsXml(const SheetLayout& layout, const CellStyleSavePlan* style_plan) {
  if (layout.columns.empty()) {
    return std::string();
  }
  std::string out;
  out.reserve(32U + layout.columns.size() * 96U);
  out.append("<cols>");
  for (const ColumnLayout& col : layout.columns) {
    out.append("<col");
    append_xml_attr_uint(out, "min", col.first + 1U);
    append_xml_attr_uint(out, "max", col.last + 1U);
    // Keep the legacy convenience of treating a non-zero programmatic
    // width as logically explicit, while preserving an explicit width="0"
    // and distinguishing it from an absent width.
    const bool has_width = HasExplicitColumnWidth(col);
    if (has_width) {
      // Shortest round-trip spelling, the same one cell values use: a
      // recalc-save neither drifts the column metric nor respells 8.7 as
      // 8.6999999999999993. Matches the row-height writer.
      append_xml_attr_number(out, "width", col.width);
      // Excel emits `customWidth="1"` whenever an explicit `width` is
      // present so a reload preserves the column metric.
      emit_xsd_bool_attr(out, "customWidth", true);
    }
    // A style written as xf 0 is left out; the width stays.
    if (col.has_style && !(style_plan != nullptr && style_plan->collapsed_to_default(col.style_xf))) {
      append_xml_attr_uint(out, "style",
                           style_plan != nullptr ? style_plan->cell_xf_index(col.style_xf) : col.style_xf);
    }
    if (col.hidden) {
      emit_xsd_bool_attr(out, "hidden", true);
    }
    if (col.outline_level != 0U) {
      append_xml_attr_uint(out, "outlineLevel", col.outline_level);
    }
    if (col.collapsed) {
      emit_xsd_bool_attr(out, "collapsed", true);
    }
    if (col.show_phonetic) {
      emit_xsd_bool_attr(out, "phonetic", true);
    }
    out.append("/>");
  }
  out.append("</cols>");
  return out;
}

/// Emits `<dimension ref="A1:..."/>`. ECMA-376 places `<dimension>` between
/// `<sheetPr>` and `<sheetViews>`; some readers (and Excel's Name Box) use it
/// to seed the used range. The box is `sheet_dimension`, the rule the .xlsb
/// writer's BrtWsDim shares; with nothing to cover it is `A1`, Excel's form
/// for an empty sheet.
std::string BuildDimensionXml(const Sheet& sheet, const CellStyleSavePlan* style_plan) {
  const std::optional<SheetDimension> box = sheet_dimension(sheet, style_plan);
  const bool any = box.has_value();
  const std::uint32_t min_row = any ? box->first_row : 0U;
  const std::uint32_t min_col = any ? box->first_col : 0U;
  const std::uint32_t max_row = any ? box->last_row : 0U;
  const std::uint32_t max_col = any ? box->last_col : 0U;
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
                              const SharedStrings* shared_strings, const std::vector<DifferentialFormat>& dxfs,
                              std::size_t sheet_index, std::uint32_t dynamic_array_cm_index,
                              const xlsb::NameShapes& name_shapes, const parser::ExternalBookIndexer* indexer,
                              const std::vector<std::string>* sheet_names, const RichErrorVmMap* rich_error_vm,
                              std::string_view book_directory, const std::optional<PhoneticProperties>& absent_phonetic,
                              const CellStyleSavePlan* style_plan,
                              const std::unordered_set<CellAddress, CellAddressHash>* recalc_marks) {
  const std::string sheet_view_xml = BuildSheetViewXml(sheet.view());
  const std::string sheet_format_xml =
      BuildSheetFormatPrXml(WithDerivedOutlineLevels(sheet.format_defaults(), sheet.layout()));
  const std::string cols_xml = BuildColsXml(sheet.layout(), style_plan);
  const SheetPrintSettings& print = sheet.print_settings();
  const WorksheetRawExtensions& raw_extensions = sheet.raw_extensions();
  const std::string page_setup_xml = PageSetupWithRelationshipId(print.page_setup_xml, printer_settings_rid);
  const std::string sheet_pr_xml = BuildNormalizedSheetPrXml(print.sheet_pr_xml);
  std::string out;
  out.reserve(192U + sheet_view_xml.size() + sheet_format_xml.size() + cols_xml.size() + sheet_pr_xml.size() +
              print.page_margins_xml.size() + page_setup_xml.size() + sheet_tables.size() * 96);
  out.append(kXmlDecl);
  out.append("<worksheet");
  append_xml_ns_decl(out, "", kNsSpreadsheetMl);
  append_xml_ns_decl(out, "r", kNsOfficeRelationships);
  // Re-declare the source worksheet root's extra namespaces so prefixed
  // attributes carried in any raw capture below resolve (keeps the output
  // well-formed; mirrors the workbook-root handling).
  out.append(sheet.root_extra_ns_attrs());
  out.append(">\n");
  // Children go out in the ECMA-376 §18.3.1.99 order. The generated
  // elements below appear in that order already; `flush_raw_before`
  // interleaves the retained source children between them, driven by the
  // slot each one recorded at read time. A single forward cursor is
  // enough because the retained list is sorted by slot.
  std::size_t raw_cursor = 0;
  const auto emit_line = [&](std::string_view xml) {
    out.append("  ");
    out.append(xml);
    out.push_back('\n');
  };
  const auto flush_raw_upto = [&](std::size_t slot) {
    while (raw_cursor < raw_extensions.size() && raw_extensions[raw_cursor].slot < slot) {
      emit_line(raw_extensions[raw_cursor].xml);
      ++raw_cursor;
    }
  };
  const auto flush_raw_before = [&](std::string_view element_name) {
    flush_raw_upto(worksheet_child::slot_of(element_name));
  };
  // The helpers stay quiet when their underlying field is at default
  // values so absent metadata yields no extra bytes.
  if (!sheet_pr_xml.empty()) {
    emit_line(sheet_pr_xml);
  }
  emit_line(BuildDimensionXml(sheet, style_plan));
  if (!sheet_view_xml.empty()) {
    emit_line(sheet_view_xml);
  }
  if (!sheet_format_xml.empty()) {
    emit_line(sheet_format_xml);
  }
  if (!cols_xml.empty()) {
    emit_line(cols_xml);
  }
  // Built after the preceding parts and before the blocks below; external
  // book numbering depends on that order.
  out.append("  ");
  AppendSheetDataXml(out, sheet, shared_strings, dynamic_array_cm_index, name_shapes, indexer, rich_error_vm,
                     style_plan, recalc_marks);
  out.push_back('\n');
  // Conditional-format blocks live between <sheetData> and <tableParts>
  // in ECMA-376 document order. Empty list => empty string, no
  // wrapper.
  // Rules naming another sheet go to the x14 extension instead.
  const CfSheetContext cf_sheet{sheet.name(), sheet_index, &dxfs, sheet_names};
  const std::string cf_xml = write_conditional_formattings(sheet.conditional_formats(), dxfs.size(), indexer, cf_sheet);
  // Data-bar settings with no legacy attribute live in the worksheet
  // `<extLst>`, which is emitted much further down; build them here so
  // the CF model is read once.
  const std::string dv_xml = BuildDataValidationsBlock(sheet, indexer, sheet_names);
  // With no `<dataValidations>` written, its attributes move to the x14 container.
  const std::string ext_lst_xml = merge_x14_dv_entries(
      merge_x14_cf_entries(sheet.ext_lst_xml(), sheet.conditional_formats(), indexer, cf_sheet), sheet.validations(),
      sheet.name(), indexer, dv_xml.empty() ? std::string_view(sheet.validations_extra_attrs()) : std::string_view(),
      sheet_names);
  const std::string merges_xml = BuildMergeCellsBlock(sheet);
  const std::string hl_xml = BuildHyperlinksBlock(sheet, hyperlink_rids, book_directory);
  // <sheetProtection> sits between <sheetData> and <mergeCells> per
  // ECMA-376 document order. Helper returns "" when protection is
  // disabled, leaving no trailing whitespace in that case.
  flush_raw_before("sheetProtection");
  {
    const std::string sp_xml = BuildSheetProtectionXml(sheet.protection());
    if (!sp_xml.empty()) {
      emit_line(sp_xml);
    }
  }
  // <autoFilter> sits between <sheetProtection>/<scenarios> and
  // <mergeCells> in ECMA-376 document order, serialized from the model.
  flush_raw_before("autoFilter");
  if (sheet.has_auto_filter()) {
    emit_line(serialize_auto_filter(*sheet.auto_filter()));
  }
  // Merge cells precede CF in ECMA-376 document order.
  flush_raw_before("mergeCells");
  if (!merges_xml.empty()) {
    emit_line(merges_xml);
  }
  if (absent_phonetic && !sheet_phonetic_properties(sheet)) {
    flush_raw_before("phoneticPr");
    emit_line(sheet_phonetic_xml(*absent_phonetic));
  }
  flush_raw_before("conditionalFormatting");
  if (!cf_xml.empty()) {
    emit_line(cf_xml);
  }
  flush_raw_before("dataValidations");
  if (!dv_xml.empty()) {
    emit_line(dv_xml);
  }
  flush_raw_before("hyperlinks");
  if (!hl_xml.empty()) {
    emit_line(hl_xml);
  }
  // <printOptions> sits between <hyperlinks> and <pageMargins>.
  flush_raw_before("printOptions");
  if (!print.print_options_xml.empty()) {
    emit_line(print.print_options_xml);
  }
  if (!print.page_margins_xml.empty()) {
    emit_line(PageMarginsWithRequiredAttributes(print.page_margins_xml));
  }
  flush_raw_before("pageSetup");
  if (!page_setup_xml.empty()) {
    emit_line(page_setup_xml);
  }
  // <headerFooter> follows <pageSetup> and precedes <rowBreaks>.
  flush_raw_before("headerFooter");
  if (!print.header_footer_xml.empty()) {
    emit_line(print.header_footer_xml);
  }
  // Manual page breaks. ECMA-376 places <rowBreaks>/<colBreaks> after
  // <pageSetup> and before drawing parts / <tableParts>.
  {
    flush_raw_before("rowBreaks");
    const std::string row_breaks_xml = BuildPageBreaksXml("rowBreaks", print.manual_row_breaks);
    if (!row_breaks_xml.empty()) {
      emit_line(row_breaks_xml);
    }
    flush_raw_before("colBreaks");
    const std::string col_breaks_xml = BuildPageBreaksXml("colBreaks", print.manual_col_breaks);
    if (!col_breaks_xml.empty()) {
      emit_line(col_breaks_xml);
    }
  }
  // <drawing> precedes <tableParts> in ECMA-376 worksheet element order.
  // The referenced DrawingML part round-trips through passthrough; here
  // we only re-emit the reference so the part stays reachable.
  flush_raw_before("drawing");
  if (!drawing_rid.empty()) {
    out.append("  <drawing r:id=\"");
    out.append(drawing_rid);
    out.append("\"/>\n");
  }
  flush_raw_before("legacyDrawing");
  if (!legacy_drawing_rid.empty()) {
    out.append("  <legacyDrawing r:id=\"");
    out.append(legacy_drawing_rid);
    out.append("\"/>\n");
  }
  flush_raw_before("tableParts");
  if (!sheet_tables.empty()) {
    out.append("  <tableParts");
    append_xml_attr_uint(out, "count", sheet_tables.size());
    out.append(">\n");
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
  flush_raw_before("extLst");
  if (!ext_lst_xml.empty()) {
    emit_line(ext_lst_xml);
  }
  // Anything the cursor has not reached -- an unplaceable element that
  // trailed the last schema child in the source -- still goes out rather
  // than being dropped at the end of the sweep.
  flush_raw_upto(worksheet_child::kCount + 1U);
  out.append("</worksheet>\n");
  return out;
}

SheetRelsResult BuildSheetRels(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                               const std::vector<EmissionPlan::PivotTablePlan>& sheet_pivot_tables,
                               const EmissionPlan::CommentsPlan& comments_plan, const EmissionPlan& plan,
                               WriteDiagnostics* diagnostics) {
  SheetRelsResult res;
  // One place records an internal relationship the package cannot honour,
  // so the log line and the counter can never disagree about how many
  // were dropped.
  const auto drop_rel = [&](std::string_view type, std::string_view target) {
    RecordDroppedRelationship("ooxml_writer.sheet_rel_skipped", type, target, diagnostics);
  };
  std::string& out = res.xml;
  out.reserve(256 + (sheet_tables.size() + sheet_pivot_tables.size() + sheet.hyperlinks().size()) * 192);
  out.append(kXmlDecl);
  BeginRelationships(out);
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
  // Pivot-table relationships follow the table relationships, listed with
  // descending rIds as Excel writes them.
  std::vector<std::string> pivot_rids;
  pivot_rids.reserve(sheet_pivot_tables.size());
  for (std::size_t i = 0; i < sheet_pivot_tables.size(); ++i) {
    pivot_rids.push_back(next_unique_rid());
  }
  for (std::size_t i = 0; i < sheet_pivot_tables.size(); ++i) {
    const std::string target = "../pivotTables/pivotTable" + std::to_string(sheet_pivot_tables[i].numeric_id) + ".xml";
    add_rel(pivot_rids[pivot_rids.size() - 1U - i], kRelPivotTable, target);
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
    add_rel(rid, kRelHyperlink, EncodeTargetSpaces(saved_hyperlink(h, plan.book_directory).target),
            /*target_external=*/true, /*escape_target=*/true);
  }
  const SheetPrintSettings& print = sheet.print_settings();
  if (!print.printer_settings_path.empty()) {
    // The reader records the path from the source sheet rels whether or
    // not the archive actually carried the part, so a truncated package
    // reaches here naming a `printerSettings*.bin` this save will not
    // write. Leaving `res.printer_settings_rid` empty also keeps the
    // `<pageSetup>` element from referencing the id.
    if (!HasPassthroughPart(plan, print.printer_settings_path)) {
      drop_rel(kRelPrinterSettings, print.printer_settings_path);
    } else {
      if (!print.printer_settings_rid.empty() && used_rids.count(print.printer_settings_rid) == 0U) {
        res.printer_settings_rid = print.printer_settings_rid;
        used_rids.insert(res.printer_settings_rid);
      } else {
        res.printer_settings_rid = next_unique_rid();
      }
      add_rel(res.printer_settings_rid, kRelPrinterSettings, TargetRelativeToWorksheet(print.printer_settings_path));
    }
  }
  // Drawing (DrawingML) relationship. The part body, its own rels, and
  // any anchored media round-trip through passthrough; here we re-mint a
  // fresh rId so the worksheet's `<drawing r:id>` element resolves.
  //
  // Same rule as the printer settings above, for the same reason: the
  // reader records the target from the source sheet rels without checking
  // the archive carried it, so a truncated package arrives naming a
  // `drawingN.xml` this save does not write. An empty `res.drawing_rid`
  // also suppresses the `<drawing>` element, so neither half of the
  // reference outlives the part.
  if (!sheet.drawing_rel_target().empty()) {
    if (!HasPassthroughPart(plan, sheet.drawing_rel_target())) {
      drop_rel(kRelDrawing, sheet.drawing_rel_target());
    } else {
      res.drawing_rid = next_unique_rid();
      add_rel(res.drawing_rid, kRelDrawing, TargetRelativeToWorksheet(sheet.drawing_rel_target()));
    }
  }
  // Comments + VML relationships when the sheet has comments. The
  // comments rel comes first; the VML rel follows so the two ids are
  // adjacent and readers that scan top-down see them as a pair.
  if (comments_plan.numeric_id != 0) {
    const std::string vml_target = "../drawings/vmlDrawing" + std::to_string(comments_plan.numeric_id) + ".vml";
    if (!comments_plan.comments_path.empty()) {
      add_rel(next_unique_rid(), kRelComments, TargetRelativeToWorksheet(comments_plan.comments_path));
    }
    res.legacy_drawing_rid = next_unique_rid();
    add_rel(res.legacy_drawing_rid, kRelVmlDrawing, vml_target);
    if (!comments_plan.threaded_path.empty()) {
      add_rel(next_unique_rid(), kRelThreadedComment, TargetRelativeToWorksheet(comments_plan.threaded_path));
    }
  }
  for (const UnknownRelationship& relationship : sheet.unknown_relationships()) {
    // Internal unknown relationships are meaningful only when the
    // corresponding passthrough payload survived collision handling. A
    // generated part at the same path is not a substitute: it may have a
    // different type or schema from the source edge's target.
    if (relationship.id.empty()) {
      continue;
    }
    if (!relationship.target_external &&
        (relationship.target.empty() || !HasPassthroughPart(plan, relationship.target))) {
      drop_rel(relationship.type, relationship.target);
      continue;
    }
    const std::string target =
        relationship.target_external ? relationship.target : TargetRelativeToWorksheet(relationship.target);
    add_rel(relationship.id, relationship.type, target, relationship.target_external, /*escape_target=*/true);
  }
  EndRelationships(out);
  return res;
}

}  // namespace io
}  // namespace formulon
