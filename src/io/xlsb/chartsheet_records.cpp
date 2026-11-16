#include "io/xlsb/chartsheet_records.h"

#include <cstdint>
#include <string>
#include <vector>

#include "io/xlsb/print_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet_layout.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtBeginSheet = 129;
constexpr std::uint16_t kBrtEndSheet = 130;
constexpr std::uint16_t kBrtBeginCsViews = 139;
constexpr std::uint16_t kBrtEndCsViews = 140;
constexpr std::uint16_t kBrtBeginCsView = 141;
constexpr std::uint16_t kBrtEndCsView = 142;
constexpr std::uint16_t kBrtCsProp = 651;
constexpr std::uint16_t kBrtCsProtection = 669;
constexpr std::uint16_t kBrtFrtBegin = 37;
constexpr std::uint16_t kBrtFrtEnd = 38;
constexpr std::uint16_t kBrtRevisionUid = 3072;

// BrtCsProp as Excel writes it for a chartsheet with no tab colour or code name.
constexpr std::uint8_t kDefaultCsProp[] = {0x01, 0x00, 0x00, 0x40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0};
// BrtBeginCsView flags: the tab is selected, the chart is zoomed to the window.
constexpr std::uint16_t kCsViewTabSelected = 0x0001U;
constexpr std::uint16_t kCsViewZoomToFit = 0x0002U;
// BrtBeginCsView: flags (u16), zoom percent (u32), workbook view (u32).
constexpr std::size_t kCsViewBytes = 10;
constexpr std::uint32_t kDefaultZoom = 100;
// BrtCsProtection as Excel writes it for a chartsheet that is not protected.
constexpr std::size_t kCsProtectionBytes = 10;

bool is_default_cs_prop(ByteSpan payload) {
  if (payload.size != sizeof(kDefaultCsProp)) {
    return false;
  }
  for (std::size_t i = 0; i < payload.size; ++i) {
    if (payload.data[i] != kDefaultCsProp[i]) {
      return false;
    }
  }
  return true;
}

bool all_zero(ByteSpan payload) {
  for (std::size_t i = 0; i < payload.size; ++i) {
    if (payload.data[i] != 0U) {
      return false;
    }
  }
  return true;
}

}  // namespace

Expected<ChartsheetXml, Error> chartsheet_xml_from_records(ByteSpan part) {
  ChartsheetXml out;
  std::string views;
  std::string drawing;
  std::vector<std::uint8_t> print_records;
  ByteSpan cursor = part;
  while (cursor.size > 0) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    const ByteSpan payload = rec.payload;
    switch (rec.type) {
      case kBrtBeginSheet:
      case kBrtEndSheet:
      case kBrtBeginCsViews:
      case kBrtEndCsViews:
      case kBrtEndCsView:
      case kBrtFrtBegin:
      case kBrtFrtEnd:
      case kBrtRevisionUid:
        break;
      case kBrtCsProp:
        out.unconverted += is_default_cs_prop(payload) ? 0U : 1U;
        break;
      case kBrtBeginCsView: {
        if (payload.size < kCsViewBytes) {
          ++out.unconverted;
          break;
        }
        const std::uint16_t flags = load_u16(payload.data);
        const std::uint32_t zoom = load_u32(payload.data + 2);
        const std::uint32_t view = load_u32(payload.data + 6);
        out.unconverted += (flags & ~(kCsViewTabSelected | kCsViewZoomToFit)) != 0U ? 1U : 0U;
        views.append("<sheetView");
        if ((flags & kCsViewTabSelected) != 0U) {
          views.append(" tabSelected=\"1\"");
        }
        if (zoom != kDefaultZoom) {
          append_xml_attr_uint(views, "zoomScale", zoom);
        }
        append_xml_attr_uint(views, "workbookViewId", view);
        if ((flags & kCsViewZoomToFit) != 0U) {
          views.append(" zoomToFit=\"1\"");
        }
        views.append("/>");
        break;
      }
      case kBrtCsProtection:
        out.unconverted += all_zero(payload) ? 0U : 1U;
        break;
      case kBrtMargins:
      case kBrtPrintOptions:
      case kBrtPageSetup:
      case kBrtBeginHeaderFooter:
      case kBrtEndHeaderFooter:
        emit_record(print_records, rec.type, payload);
        break;
      case kBrtDrawing: {
        ByteSpan p = payload;
        ASSIGN_OR_RETURN(const std::string rid, read_xlwidestring(p));
        drawing = "<drawing r:id=\"";
        drawing.append(rid);
        drawing.append("\"/>");
        break;
      }
      default:
        ++out.unconverted;
        break;
    }
  }
  SheetPrintSettings print;
  decode_print_records(print_records, print);
  out.unconverted += unwritten_print_families(print_records, print);
  // A worksheet page setup or print options record has no place in a chartsheet.
  out.unconverted += print.page_setup_xml.empty() ? 0U : 1U;
  out.unconverted += print.print_options_xml.empty() ? 0U : 1U;

  // CT_Chartsheet order: sheetPr, sheetViews, (sheetProtection), pageMargins, (pageSetup), headerFooter, drawing.
  std::string& xml = out.xml;
  xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
  xml.append(
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetPr/>");
  if (!views.empty()) {
    xml.append("<sheetViews>").append(views).append("</sheetViews>");
  }
  xml.append(print.page_margins_xml);
  xml.append(print.header_footer_xml);
  xml.append(drawing);
  xml.append("</chartsheet>");
  return out;
}

Expected<ChartsheetBin, Error> chartsheet_records_from_xml(std::string_view xml) {
  pugi::xml_document doc;
  if (!doc.load_buffer(xml.data(), xml.size()) || local_name(doc.document_element()) != "chartsheet") {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "chartsheet part does not parse",
                      "context=xlsb_chartsheet_writer");
  }
  const pugi::xml_node root = doc.document_element();
  ChartsheetBin out;
  std::vector<std::uint8_t>& dst = out.bytes;
  const std::vector<std::uint8_t> none;
  emit_record(dst, kBrtBeginSheet, none);

  const pugi::xml_node sheet_pr = child_local(root, "sheetPr");
  out.unconverted += sheet_pr && (sheet_pr.first_attribute() || sheet_pr.first_child()) ? 1U : 0U;
  emit_record(dst, kBrtCsProp, ByteSpan{kDefaultCsProp, sizeof(kDefaultCsProp)});

  emit_record(dst, kBrtBeginCsViews, none);
  for (const pugi::xml_node view : children_local(child_local(root, "sheetViews"), "sheetView")) {
    std::uint16_t flags = 0;
    std::uint32_t zoom = kDefaultZoom;
    std::uint32_t book_view = 0;
    for (const pugi::xml_attribute attr : view.attributes()) {
      const std::string_view name = attr.name();
      if (name == "tabSelected") {
        if (attr.as_bool()) {
          flags |= kCsViewTabSelected;
        }
      } else if (name == "zoomToFit") {
        if (attr.as_bool()) {
          flags |= kCsViewZoomToFit;
        }
      } else if (name == "zoomScale") {
        zoom = attr.as_uint(kDefaultZoom);
      } else if (name == "workbookViewId") {
        book_view = attr.as_uint();
      } else {
        ++out.unconverted;
      }
    }
    out.unconverted += view.first_child() ? 1U : 0U;
    std::vector<std::uint8_t> payload;
    emit_u16(payload, flags);
    emit_u32(payload, zoom);
    emit_u32(payload, book_view);
    emit_record(dst, kBrtBeginCsView, payload);
    emit_record(dst, kBrtEndCsView, none);
  }
  emit_record(dst, kBrtEndCsViews, none);

  // Excel writes the protection record unset when the sheet is not protected.
  out.unconverted += child_local(root, "sheetProtection") ? 1U : 0U;
  emit_record(dst, kBrtCsProtection, std::vector<std::uint8_t>(kCsProtectionBytes, 0U));

  SheetPrintSettings print;
  if (const pugi::xml_node margins = child_local(root, "pageMargins")) {
    print.page_margins_xml = raw_xml(margins);
  }
  if (const pugi::xml_node header_footer = child_local(root, "headerFooter")) {
    print.header_footer_xml = raw_xml(header_footer);
  }
  emit_chartsheet_print_records(dst, print);
  out.unconverted += child_local(root, "pageSetup") ? 1U : 0U;

  if (const pugi::xml_node drawing = child_local(root, "drawing")) {
    std::vector<std::uint8_t> payload;
    std::string_view rid;
    for (const pugi::xml_attribute attr : drawing.attributes()) {
      if (local_name(attr.name()) == "id") {
        rid = attr.value();
      }
    }
    emit_xlwidestring(payload, rid);
    emit_record(dst, kBrtDrawing, payload);
  }
  for (const pugi::xml_node child : root.children()) {
    const std::string_view name = local_name(child);
    if (child.type() == pugi::node_element && name != "sheetPr" && name != "sheetViews" && name != "sheetProtection" &&
        name != "pageMargins" && name != "pageSetup" && name != "headerFooter" && name != "drawing") {
      ++out.unconverted;
    }
  }
  emit_record(dst, kBrtEndSheet, none);
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
