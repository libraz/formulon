#include "io/ooxml/worksheet_metadata_reader.h"

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <utility>
#include <vector>

#include "io/auto_filter_xml.h"
#include "io/cf_reader.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/print_settings_parse.h"
#include "io/package_diagnostics.h"
#include "io/sheet_layout_reader.h"
#include "io/sheet_overlay_reader.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace ooxml {
namespace {

/// True when some part of the model already carries `name`'s content, so
/// a save rebuilds the element rather than replaying the source bytes.
///
/// This is the complement of `WorksheetRawChild`: every `<worksheet>`
/// child not listed here is preserved verbatim. Adding model support for
/// an element means adding its name here, which is also what stops the
/// same content being written twice.
bool WorksheetChildIsModelled(std::string_view name) {
  // `dimension`, `drawing`, `legacyDrawing` and `tableParts` are absent
  // from the model proper but are still regenerated on save -- from the
  // cell bounding box and from the relationship ids the writer mints --
  // so replaying the source copy would duplicate them.
  static constexpr std::string_view kModelled[] = {
      "sheetPr",         "dimension",       "sheetViews",   "sheetFormatPr", "cols",
      "sheetData",       "sheetProtection", "autoFilter",   "mergeCells",    "conditionalFormatting",
      "dataValidations", "hyperlinks",      "printOptions", "pageMargins",   "pageSetup",
      "headerFooter",    "rowBreaks",       "colBreaks",    "drawing",       "legacyDrawing",
      "tableParts",      "extLst",
  };
  for (const std::string_view modelled : kModelled) {
    if (modelled == name) {
      return true;
    }
  }
  return false;
}

/// Captures every `<worksheet>` child the model does not fold in, tagged
/// with the schema slot the writer must put it back at.
///
/// The sweep is by exclusion, so an element no release has modelled yet
/// still survives a round trip. A name outside the schema sequence takes
/// the slot of the child before it, keeping it between the same two
/// siblings. The result is stable-sorted by slot so the writer can emit
/// it with a single forward cursor, which also repairs a source whose
/// children were not in schema order.
void CaptureUnconsumedWorksheetChildren(const pugi::xml_node& worksheet, WorksheetRawExtensions& out) {
  out.clear();
  std::size_t previous_slot = 0;
  for (pugi::xml_node child = worksheet.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string_view name = child.name();
    const std::size_t slot = worksheet_child::slot_of(name);
    if (slot != worksheet_child::kCount) {
      previous_slot = slot;
    }
    if (WorksheetChildIsModelled(name)) {
      continue;
    }
    out.push_back(WorksheetRawChild{static_cast<std::uint32_t>(slot == worksheet_child::kCount ? previous_slot : slot),
                                    raw_xml(child)});
  }
  sort_by_index(out, [](const WorksheetRawChild& lhs, const WorksheetRawChild& rhs) { return lhs.slot < rhs.slot; });
}

}  // namespace

/// Returns a copy of `sheet_xml` with the `<sheetData>` element's children
/// removed (the open / close tags are kept as an empty element). This lets
/// the SAX path parse the small non-cell worksheet metadata as a DOM
/// without materialising the full cell tree the SAX scanner exists to
/// avoid. Returns the input unchanged when there is no `<sheetData>` (or
/// it is already empty / self-closing).
std::vector<std::uint8_t> build_worksheet_shell_bytes(const std::vector<std::uint8_t>& sheet_xml) {
  const std::string_view sv(reinterpret_cast<const char*>(sheet_xml.data()), sheet_xml.size());
  constexpr std::string_view kOpen = "<sheetData";
  constexpr std::string_view kClose = "</sheetData>";
  // Locate the `<sheetData` open tag, requiring a real name boundary after
  // it so `<sheetDataX>` (hypothetical) does not match.
  std::size_t open = std::string_view::npos;
  for (std::size_t from = 0;;) {
    const std::size_t hit = sv.find(kOpen, from);
    if (hit == std::string_view::npos) {
      break;
    }
    const std::size_t after = hit + kOpen.size();
    const char d = after < sv.size() ? sv[after] : '\0';
    if (d == ' ' || d == '\t' || d == '\r' || d == '\n' || d == '>' || d == '/') {
      open = hit;
      break;
    }
    from = hit + 1;
  }
  if (open == std::string_view::npos) {
    return sheet_xml;
  }
  const std::size_t gt = sv.find('>', open);
  if (gt == std::string_view::npos || sv[gt - 1] == '/') {
    // Malformed, or a self-closing `<sheetData/>` with no children.
    return sheet_xml;
  }
  const std::size_t close = sv.find(kClose, gt + 1);
  if (close == std::string_view::npos) {
    return sheet_xml;
  }
  const std::size_t close_end = close + kClose.size();
  std::vector<std::uint8_t> out;
  out.reserve((gt + 1) + kClose.size() + (sheet_xml.size() - close_end));
  out.insert(out.end(), sheet_xml.begin(), sheet_xml.begin() + static_cast<std::ptrdiff_t>(gt + 1));
  out.insert(out.end(), kClose.begin(), kClose.end());
  out.insert(out.end(), sheet_xml.begin() + static_cast<std::ptrdiff_t>(close_end), sheet_xml.end());
  return out;
}

/// Reads every non-cell worksheet element (siblings of `<sheetData>`) from
/// `doc` into sheet `i`: conditional formats, view / layout, merges,
/// hyperlinks, data validations, sheet protection, and the raw print
/// settings (sheetPr / pageMargins / pageSetup / printOptions /
/// headerFooter / autoFilter / row + col breaks). Shared between the DOM
/// path (full document) and the SAX path (metadata shell). Per-row
/// overrides (`<row ht=/hidden=>`) live inside `<sheetData>` and are only
/// populated on the DOM path; on the SAX shell that content is stripped.
/// Every overlay entry these readers drop lands in `diagnostics`.
Expected<void, Error> apply_worksheet_metadata(const pugi::xml_document& doc, std::size_t i, Workbook& wb,
                                               ReadDiagnostics* diagnostics) {
  const pugi::xml_node worksheet = doc.child("worksheet");
  ASSIGN_OR_RETURN(auto cfs, read_conditional_formats(worksheet, diagnostics));
  // The `<dxfs>` table is loaded before any sheet, so this is the first
  // point where a rule's `dxfId` can be checked against it. Both the DOM
  // and the SAX path reach the model through here, which is what makes
  // "a loaded workbook holds no unresolvable dxf_id" hold for the reader
  // as a whole rather than for one of its two paths.
  normalize_cf_dxf_ids(cfs, wb.styles().dxfs.size());
  wb.sheet(i).mutable_conditional_formats() = std::move(cfs);
  RETURN_IF_ERROR(read_sheet_view_and_layout(doc, i, wb));
  ASSIGN_OR_RETURN(auto merges, read_merges(worksheet, diagnostics));
  wb.sheet(i).mutable_merges() = std::move(merges);
  ASSIGN_OR_RETURN(auto hls, read_hyperlinks(worksheet, diagnostics));
  wb.sheet(i).mutable_hyperlinks() = std::move(hls);
  ASSIGN_OR_RETURN(auto dvs, read_data_validations(worksheet, diagnostics));
  wb.sheet(i).mutable_validations() = std::move(dvs);
  wb.sheet(i).mutable_protection() = read_sheet_protection(worksheet);
  CaptureUnconsumedWorksheetChildren(worksheet, wb.sheet(i).mutable_raw_extensions());

  SheetPrintSettings& print = wb.sheet(i).mutable_print_settings();
  if (pugi::xml_node sheet_pr = worksheet.child("sheetPr"); sheet_pr) {
    // Capture the whole `<sheetPr>` verbatim whenever it exists — it may
    // carry only `tabColor` / `codeName` (VBA binding) with no
    // `<pageSetUpPr>` child, and gating on that child dropped such sheets'
    // `<sheetPr>` entirely on save. The structured `fit_to_page` view is
    // populated additionally when `<pageSetUpPr>` is present.
    print.sheet_pr_xml = raw_xml(sheet_pr);
    print.page_setup.fit_to_page = read_fit_to_page(sheet_pr);
  }
  if (pugi::xml_node page_margins = worksheet.child("pageMargins")) {
    print.page_margins_xml = raw_xml(page_margins);
    apply_structured_page_margins(page_margins, print.page_margins);
  }
  if (pugi::xml_node page_setup = worksheet.child("pageSetup")) {
    print.page_setup_xml = raw_xml(page_setup);
    print.printer_settings_rid = ooxml::relationship_ref_id(page_setup);
    apply_structured_page_setup(page_setup, print.page_setup);
  }
  if (pugi::xml_node print_options = worksheet.child("printOptions")) {
    print.print_options_xml = raw_xml(print_options);
  }
  if (pugi::xml_node header_footer = worksheet.child("headerFooter")) {
    print.header_footer_xml = raw_xml(header_footer);
  }
  if (pugi::xml_node auto_filter = worksheet.child("autoFilter")) {
    wb.sheet(i).set_auto_filter(auto_filter_from_xml(raw_xml(auto_filter)));
  }
  // Worksheet-level `<extLst>` holds the *data* for 2010+ extensions —
  // notably `x14:conditionalFormattings` (DataBar negative-fill / axis /
  // gradient), linked to legacy `cfRule`s by the base `id` attribute
  // (see `cf_reader.h`, which decodes the DataBar fields out of this
  // block into `cf::DataBarSpec`). Capture the block raw regardless, so
  // any other 2010+ extension content it carries (unrelated to CF)
  // survives a save cycle unchanged. Living in this shared helper means
  // the SAX path recovers it too, via the metadata shell.
  if (pugi::xml_node ext_lst = worksheet.child("extLst")) {
    wb.sheet(i).set_ext_lst_xml(raw_xml(ext_lst));
  }
  // Capture the worksheet root's extra namespace declarations so any
  // prefixed attribute carried inside a raw capture above resolves when
  // re-emitted (mirrors the workbook-root handling; keeps the output
  // well-formed).
  wb.sheet(i).set_root_extra_ns_attrs(capture_root_extra_ns_attrs(worksheet));
  read_manual_breaks(worksheet.child("rowBreaks"), print.manual_row_breaks);
  read_manual_breaks(worksheet.child("colBreaks"), print.manual_col_breaks);
  return Expected<void, Error>::Ok();
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
