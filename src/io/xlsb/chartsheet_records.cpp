#include "io/xlsb/chartsheet_records.h"

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/print_records.h"
#include "io/xlsb/protection_records.h"
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
constexpr std::uint16_t kBrtCsProtectionIso = 679;
constexpr std::uint16_t kBrtCsPageSetup = 652;
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
// BrtCsProtection: legacy password hash (u16), contents locked (u32), objects locked (u32).
constexpr std::size_t kCsProtectionBytes = 10;
// BrtCsProp: flags (u16), the tab colour (8 bytes), the code name (XLWideString).
constexpr std::size_t kCsPropColorAt = 2;
constexpr std::uint8_t kColorAuto = 0x00;   // fValidRGB clear, automatic
constexpr std::uint8_t kColorRgb = 0x05;    // fValidRGB, xColorType 2
constexpr std::uint8_t kColorTheme = 0x07;  // fValidRGB, xColorType 3
constexpr std::uint8_t kColorIndexAuto = 0x40;
// BrtCsPageSetup: paper (u32), horizontal and vertical dpi (u32), copies (u32), first page (u16), flags (u16),
// printer settings relationship (XLNullableWideString). Flags as Excel writes them.
constexpr std::size_t kCsPageSetupFixedBytes = 20;
constexpr std::uint16_t kPageLandscape = 0x0001U;
constexpr std::uint16_t kPageBlackAndWhite = 0x0004U;
constexpr std::uint16_t kPageUseFirstPage = 0x0010U;
constexpr std::uint16_t kPageDraft = 0x0020U;
constexpr std::uint16_t kPageKnownFlags = kPageLandscape | kPageBlackAndWhite | kPageUseFirstPage | kPageDraft;

constexpr char kHexDigits[] = "0123456789ABCDEF";

// `rgba` (red, green, blue, alpha) as the AARRGGBB text OOXML writes.
std::string hex_argb(const std::uint8_t* rgba) {
  std::string text;
  for (const std::uint8_t byte : {rgba[3], rgba[0], rgba[1], rgba[2]}) {
    text.push_back(kHexDigits[byte >> 4U]);
    text.push_back(kHexDigits[byte & 0x0FU]);
  }
  return text;
}

// AARRGGBB text as a 32-bit ARGB value; nullopt unless it is exactly eight hex digits.
std::optional<std::uint32_t> parse_argb(std::string_view text) {
  if (text.size() != 8U) {
    return std::nullopt;
  }
  std::uint32_t value = 0;
  for (const char c : text) {
    const std::uint32_t digit = c >= '0' && c <= '9'   ? static_cast<std::uint32_t>(c - '0')
                                : c >= 'A' && c <= 'F' ? static_cast<std::uint32_t>(c - 'A' + 10)
                                : c >= 'a' && c <= 'f' ? static_cast<std::uint32_t>(c - 'a' + 10)
                                                       : 16U;
    if (digit == 16U) {
      return std::nullopt;
    }
    value = (value << 4U) | digit;
  }
  return value;
}

// `<sheetPr>` of a BrtCsProp payload; nullopt for a state whose XML is not measured (a code name, a tinted or
// indexed colour, other flags).
std::optional<std::string> sheet_pr_xml(ByteSpan payload) {
  if (payload.size != sizeof(kDefaultCsProp) || load_u16(payload.data) != 0x0001U ||
      load_u32(payload.data + 10) != 0U) {
    return std::nullopt;
  }
  const std::uint8_t* color = payload.data + kCsPropColorAt;
  const std::int16_t tint = static_cast<std::int16_t>(load_u16(color + 2));
  if (color[0] == kColorAuto && color[1] == kColorIndexAuto && tint == 0) {
    return std::string("<sheetPr/>");
  }
  if (color[0] == kColorRgb && tint == 0) {
    return "<sheetPr><tabColor rgb=\"" + hex_argb(color + 4) + "\"/></sheetPr>";
  }
  if (color[0] == kColorTheme && tint == 0) {
    std::string xml = "<sheetPr><tabColor";
    append_xml_attr_uint(xml, "theme", color[1]);
    return xml + "/></sheetPr>";
  }
  return std::nullopt;
}

// BrtCsProp for `<sheetPr>`; nullopt when it states something without a measured record form.
std::optional<std::vector<std::uint8_t>> cs_prop_record(const pugi::xml_node& sheet_pr) {
  std::vector<std::uint8_t> payload(kDefaultCsProp, kDefaultCsProp + sizeof(kDefaultCsProp));
  if (!sheet_pr) {
    return payload;
  }
  if (sheet_pr.first_attribute()) {
    return std::nullopt;
  }
  for (const pugi::xml_node child : sheet_pr.children()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    if (local_name(child) != "tabColor") {
      return std::nullopt;
    }
    std::uint8_t* color = payload.data() + kCsPropColorAt;
    const pugi::xml_attribute rgb = child.attribute("rgb");
    const pugi::xml_attribute theme = child.attribute("theme");
    if (child.attribute("tint") || child.attribute("indexed") || child.attribute("auto") || (rgb && theme)) {
      return std::nullopt;
    }
    if (rgb) {
      const std::optional<std::uint32_t> parsed = parse_argb(rgb.value());
      if (!parsed) {
        return std::nullopt;
      }
      const std::uint32_t argb = *parsed;
      color[0] = kColorRgb;
      color[1] = 0xFF;
      color[4] = static_cast<std::uint8_t>(argb >> 16U);
      color[5] = static_cast<std::uint8_t>(argb >> 8U);
      color[6] = static_cast<std::uint8_t>(argb);
      color[7] = static_cast<std::uint8_t>(argb >> 24U);
    } else if (theme) {
      // Excel also stores the theme colour's rendered value, which this writer does not resolve.
      return std::nullopt;
    }
  }
  return payload;
}

// `<pageSetup>` of a BrtCsPageSetup payload; nullopt for a state whose XML is not measured.
std::optional<std::string> page_setup_xml(ByteSpan payload) {
  if (payload.size < kCsPageSetupFixedBytes + 4U || load_u32(payload.data + kCsPageSetupFixedBytes) != 0xFFFFFFFFU) {
    return std::nullopt;
  }
  const std::uint32_t paper = load_u32(payload.data);
  const std::uint32_t h_dpi = load_u32(payload.data + 4);
  const std::uint32_t v_dpi = load_u32(payload.data + 8);
  const std::uint32_t copies = load_u32(payload.data + 12);
  const std::uint16_t first = load_u16(payload.data + 16);
  const std::uint16_t flags = load_u16(payload.data + 18);
  if ((flags & ~kPageKnownFlags) != 0U || copies == 0U || ((flags & kPageUseFirstPage) == 0U && first != 1U)) {
    return std::nullopt;
  }
  std::string xml = "<pageSetup";
  append_xml_attr_uint(xml, "paperSize", paper);
  if ((flags & kPageUseFirstPage) != 0U) {
    append_xml_attr_uint(xml, "firstPageNumber", first);
  }
  xml.append((flags & kPageLandscape) != 0U ? " orientation=\"landscape\"" : " orientation=\"portrait\"");
  if ((flags & kPageBlackAndWhite) != 0U) {
    xml.append(" blackAndWhite=\"1\"");
  }
  if ((flags & kPageDraft) != 0U) {
    xml.append(" draft=\"1\"");
  }
  if ((flags & kPageUseFirstPage) != 0U) {
    xml.append(" useFirstPageNumber=\"1\"");
  }
  append_xml_attr_uint(xml, "horizontalDpi", h_dpi);
  append_xml_attr_uint(xml, "verticalDpi", v_dpi);
  if (copies != 1U) {
    append_xml_attr_uint(xml, "copies", copies);
  }
  xml.append("/>");
  return xml;
}

// BrtCsPageSetup for `<pageSetup>`; nullopt when it states something without a measured record form.
std::optional<std::vector<std::uint8_t>> cs_page_setup_record(const pugi::xml_node& page_setup) {
  std::uint32_t paper = 1;
  std::uint32_t h_dpi = 0;
  std::uint32_t v_dpi = 0;
  std::uint32_t copies = 1;
  std::uint16_t first = 1;
  std::uint16_t flags = 0;
  bool oriented = false;
  for (const pugi::xml_attribute attr : page_setup.attributes()) {
    const std::string_view name = attr.name();
    if (name == "paperSize") {
      paper = attr.as_uint();
    } else if (name == "horizontalDpi") {
      h_dpi = attr.as_uint();
    } else if (name == "verticalDpi") {
      v_dpi = attr.as_uint();
    } else if (name == "firstPageNumber") {
      first = static_cast<std::uint16_t>(attr.as_uint());
    } else if (name == "copies" && attr.as_uint() != 0U) {
      copies = attr.as_uint();
    } else if (name == "orientation" && std::string_view(attr.value()) == "landscape") {
      flags |= kPageLandscape;
      oriented = true;
    } else if (name == "orientation" && std::string_view(attr.value()) == "portrait") {
      oriented = true;
    } else if (name == "blackAndWhite" && attr.as_bool()) {
      flags |= kPageBlackAndWhite;
    } else if (name == "draft" && attr.as_bool()) {
      flags |= kPageDraft;
    } else if (name == "useFirstPageNumber" && attr.as_bool()) {
      flags |= kPageUseFirstPage;
    } else {
      return std::nullopt;
    }
  }
  if (!oriented) {
    return std::nullopt;
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, paper);
  emit_u32(payload, h_dpi);
  emit_u32(payload, v_dpi);
  emit_u32(payload, copies);
  emit_u16(payload, first);
  emit_u16(payload, flags);
  emit_u32(payload, 0xFFFFFFFFU);
  return payload;
}

// `<sheetProtection>` of BrtCsProtection and an optional BrtCsProtectionIso; empty when unprotected, nullopt for
// a state whose XML is not measured (a legacy password).
std::optional<std::string> protection_xml(ByteSpan legacy, ByteSpan iso) {
  if (legacy.size != kCsProtectionBytes || load_u16(legacy.data) != 0U) {
    return std::nullopt;
  }
  const bool content = load_u32(legacy.data + 2) != 0U;
  const bool objects = load_u32(legacy.data + 6) != 0U;
  std::string xml;
  if (iso.size > 0U) {
    ByteSpan p = iso;
    auto spin = read_u32(p);
    auto iso_content = read_u32(p);
    auto iso_objects = read_u32(p);
    auto hash_size = read_u32(p);
    if (!spin || !iso_content || !iso_objects || !hash_size || hash_size.value() > p.size) {
      return std::nullopt;
    }
    const std::string hash = base64_encode(p.data, hash_size.value());
    p.data += hash_size.value();
    p.size -= hash_size.value();
    auto salt_size = read_u32(p);
    if (!salt_size || salt_size.value() > p.size) {
      return std::nullopt;
    }
    const std::string salt = base64_encode(p.data, salt_size.value());
    p.data += salt_size.value();
    p.size -= salt_size.value();
    auto algorithm = read_xlwidestring(p);
    if (!algorithm || p.size != 0U) {
      return std::nullopt;
    }
    append_xml_attr(xml, "algorithmName", algorithm.value());
    append_xml_attr(xml, "hashValue", hash);
    append_xml_attr(xml, "saltValue", salt);
    append_xml_attr_uint(xml, "spinCount", spin.value());
  }
  if (!content && !objects && xml.empty()) {
    return std::string();
  }
  if (content) {
    xml.append(" content=\"1\"");
  }
  if (objects) {
    xml.append(" objects=\"1\"");
  }
  return "<sheetProtection" + xml + "/>";
}

// BrtCsProtectionIso (when a hash is set) and BrtCsProtection for `<sheetProtection>`; false when it states
// something without a measured record form.
bool emit_cs_protection(std::vector<std::uint8_t>& dst, const pugi::xml_node& protection) {
  std::uint32_t content = 0;
  std::uint32_t objects = 0;
  std::string algorithm;
  std::vector<std::uint8_t> hash;
  std::vector<std::uint8_t> salt;
  std::uint32_t spin = 0;
  for (const pugi::xml_attribute attr : protection.attributes()) {
    const std::string_view name = attr.name();
    if (name == "content") {
      content = attr.as_bool() ? 1U : 0U;
    } else if (name == "objects") {
      objects = attr.as_bool() ? 1U : 0U;
    } else if (name == "algorithmName") {
      algorithm = attr.value();
    } else if (name == "hashValue") {
      if (!base64_decode(attr.value(), hash)) {
        return false;
      }
    } else if (name == "saltValue") {
      if (!base64_decode(attr.value(), salt)) {
        return false;
      }
    } else if (name == "spinCount") {
      spin = attr.as_uint();
    } else {
      return false;
    }
  }
  if (!algorithm.empty() || !hash.empty() || !salt.empty()) {
    if (algorithm.empty() || hash.empty()) {
      return false;
    }
    std::vector<std::uint8_t> payload;
    emit_u32(payload, spin);
    emit_u32(payload, content);
    emit_u32(payload, objects);
    emit_u32(payload, static_cast<std::uint32_t>(hash.size()));
    payload.insert(payload.end(), hash.begin(), hash.end());
    emit_u32(payload, static_cast<std::uint32_t>(salt.size()));
    payload.insert(payload.end(), salt.begin(), salt.end());
    emit_xlwidestring(payload, algorithm);
    emit_record(dst, kBrtCsProtectionIso, payload);
  }
  std::vector<std::uint8_t> payload;
  emit_u16(payload, 0U);
  emit_u32(payload, content);
  emit_u32(payload, objects);
  emit_record(dst, kBrtCsProtection, payload);
  return true;
}

}  // namespace

Expected<ChartsheetXml, Error> chartsheet_xml_from_records(ByteSpan part) {
  ChartsheetXml out;
  std::string sheet_pr = "<sheetPr/>";
  std::string views;
  std::string protection;
  std::string page_setup;
  std::vector<std::uint8_t> iso;
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
        if (std::optional<std::string> xml = sheet_pr_xml(payload)) {
          sheet_pr = std::move(*xml);
        } else {
          ++out.unconverted;
        }
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
      case kBrtCsProtectionIso:
        iso.assign(payload.data, payload.data + payload.size);
        break;
      case kBrtCsProtection:
        if (std::optional<std::string> xml = protection_xml(payload, ByteSpan{iso.data(), iso.size()})) {
          protection = std::move(*xml);
        } else {
          ++out.unconverted;
        }
        break;
      case kBrtCsPageSetup:
        if (std::optional<std::string> xml = page_setup_xml(payload)) {
          page_setup = std::move(*xml);
        } else {
          ++out.unconverted;
        }
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
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
  xml.append(sheet_pr);
  if (!views.empty()) {
    xml.append("<sheetViews>").append(views).append("</sheetViews>");
  }
  xml.append(protection);
  xml.append(print.page_margins_xml);
  xml.append(page_setup);
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

  std::optional<std::vector<std::uint8_t>> prop = cs_prop_record(child_local(root, "sheetPr"));
  if (!prop) {
    ++out.unconverted;
    prop.emplace(kDefaultCsProp, kDefaultCsProp + sizeof(kDefaultCsProp));
  }
  emit_record(dst, kBrtCsProp, *prop);

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
  const pugi::xml_node protection = child_local(root, "sheetProtection");
  if (!protection || !emit_cs_protection(dst, protection)) {
    out.unconverted += protection ? 1U : 0U;
    emit_record(dst, kBrtCsProtection, std::vector<std::uint8_t>(kCsProtectionBytes, 0U));
  }

  SheetPrintSettings print;
  if (const pugi::xml_node margins = child_local(root, "pageMargins")) {
    print.page_margins_xml = raw_xml(margins);
  }
  if (const pugi::xml_node header_footer = child_local(root, "headerFooter")) {
    print.header_footer_xml = raw_xml(header_footer);
  }
  SheetPrintSettings margins_only = print;
  margins_only.header_footer_xml.clear();
  emit_chartsheet_print_records(dst, margins_only);
  if (const pugi::xml_node page_setup = child_local(root, "pageSetup")) {
    if (std::optional<std::vector<std::uint8_t>> record = cs_page_setup_record(page_setup)) {
      emit_record(dst, kBrtCsPageSetup, *record);
    } else {
      ++out.unconverted;
    }
  }
  SheetPrintSettings header_only;
  header_only.header_footer_xml = print.header_footer_xml;
  emit_chartsheet_print_records(dst, header_only);

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
