#include "io/xlsb/chartsheet_records.h"

#include <algorithm>
#include <array>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/print_settings_parse.h"
#include "io/passthrough_part_index.h"
#include "io/xlsb/brt_color.h"
#include "io/xlsb/print_records.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_properties_records.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "sheet_layout.h"
#include "utils/a1_ref.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"

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
constexpr std::uint8_t kColorAuto = 0x00;     // fValidRGB clear, automatic
constexpr std::uint8_t kColorIndexed = 0x03;  // fValidRGB, xColorType 1
constexpr std::uint8_t kColorRgb = 0x05;      // fValidRGB, xColorType 2
constexpr std::uint8_t kColorTheme = 0x07;    // fValidRGB, xColorType 3
constexpr std::uint8_t kColorIndexAuto = 0x40;
// BrtCsPageSetup: paper (u32), horizontal and vertical dpi (u32), copies (u32), first page (u16), flags (u16),
// printer settings relationship (XLNullableWideString). Flags as Excel writes them.
constexpr std::size_t kCsPageSetupFixedBytes = 20;
constexpr std::uint32_t kNullWideString = 0xFFFFFFFFU;
// The dpi an absent horizontalDpi or verticalDpi states.
constexpr std::uint32_t kDefaultDpi = 600;
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

// `text` as a hex number; nullopt unless it is exactly `digits` hex digits (eight for AARRGGBB).
std::optional<std::uint32_t> parse_hex(std::string_view text, std::size_t digits) {
  if (text.size() != digits) {
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

// `<sheetPr>` of a BrtCsProp payload; nullopt for a state whose XML is not measured (a code name, a tinted RGB or
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
  if (color[0] == kColorIndexed && tint == 0) {
    std::string xml = "<sheetPr><tabColor";
    append_xml_attr_uint(xml, "indexed", color[1]);
    return xml + "/></sheetPr>";
  }
  if (color[0] == kColorTheme) {
    std::string xml = "<sheetPr><tabColor";
    append_xml_attr_uint(xml, "theme", color[1]);
    if (tint != 0) {
      append_xml_attr_number(xml, "tint", static_cast<double>(tint) / kBrtColorTintScale);
    }
    return xml + "/></sheetPr>";
  }
  return std::nullopt;
}

// BrtCsProp for `<sheetPr>`; nullopt when it states something without a measured record form.
std::optional<std::vector<std::uint8_t>> cs_prop_record(const pugi::xml_node& sheet_pr, const Theme& theme) {
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
    const pugi::xml_attribute rgb = child.attribute("rgb");
    const pugi::xml_attribute theme_index = child.attribute("theme");
    const pugi::xml_attribute indexed = child.attribute("indexed");
    const pugi::xml_attribute tint = child.attribute("tint");
    const pugi::xml_attribute automatic = child.attribute("auto");
    const int selectors = (rgb ? 1 : 0) + (theme_index ? 1 : 0) + (indexed ? 1 : 0) + (automatic ? 1 : 0);
    if (selectors > 1 || (tint && !theme_index) || (automatic && !automatic.as_bool())) {
      return std::nullopt;
    }
    if (automatic) {
      continue;  // Excel saves an automatic tab colour as none.
    }
    ColorSpec spec;
    if (indexed) {
      spec.kind = ColorSpec::Kind::kIndexed;
      spec.indexed = indexed.as_uint();
    } else if (rgb) {
      const std::optional<std::uint32_t> parsed = parse_hex(rgb.value(), 8U);
      if (!parsed) {
        return std::nullopt;
      }
      spec.kind = ColorSpec::Kind::kRgb;
      spec.rgb = *parsed;
    } else if (theme_index) {
      spec.kind = ColorSpec::Kind::kTheme;
      spec.theme = theme_index.as_uint();
      spec.tint = tint.as_double();
    }
    const std::array<std::uint8_t, 8> color = encode_tab_color(spec, theme);
    std::copy(color.begin(), color.end(), payload.begin() + kCsPropColorAt);
  }
  return payload;
}

// `<pageSetup>` of a BrtCsPageSetup payload; nullopt for a state whose XML is not measured.
std::optional<std::string> page_setup_xml(ByteSpan payload) {
  if (payload.size < kCsPageSetupFixedBytes + 4U) {
    return std::nullopt;
  }
  std::string rid;
  if (load_u32(payload.data + kCsPageSetupFixedBytes) != kNullWideString) {
    ByteSpan tail{payload.data + kCsPageSetupFixedBytes, payload.size - kCsPageSetupFixedBytes};
    auto read = read_xlwidestring(tail);
    if (!read || tail.size != 0U || read.value().empty()) {
      return std::nullopt;
    }
    rid = std::move(read.value());
  } else if (payload.size != kCsPageSetupFixedBytes + 4U) {
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
  if (h_dpi != kDefaultDpi) {
    append_xml_attr_uint(xml, "horizontalDpi", h_dpi);
  }
  if (v_dpi != kDefaultDpi) {
    append_xml_attr_uint(xml, "verticalDpi", v_dpi);
  }
  if (copies != 1U) {
    append_xml_attr_uint(xml, "copies", copies);
  }
  if (!rid.empty()) {
    append_xml_attr(xml, "r:id", rid);
  }
  xml.append("/>");
  return xml;
}

// BrtCsPageSetup for `<pageSetup>`; nullopt when it states something without a measured record form.
std::optional<std::vector<std::uint8_t>> cs_page_setup_record(const pugi::xml_node& page_setup) {
  std::uint32_t paper = 1;
  std::uint32_t h_dpi = kDefaultDpi;
  std::uint32_t v_dpi = kDefaultDpi;
  std::uint32_t copies = 1;
  std::string_view rid;
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
    } else if (local_name(attr.name()) == "id" && *attr.value() != '\0') {
      rid = attr.value();
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
  if (rid.empty()) {
    emit_u32(payload, kNullWideString);
  } else {
    emit_xlwidestring(payload, rid);
  }
  return payload;
}

// `<sheetProtection>` of BrtCsProtection and an optional BrtCsProtectionIso; empty when unprotected, nullopt for
// a state whose XML is not measured (a legacy password beside a hash).
std::optional<std::string> protection_xml(ByteSpan legacy, ByteSpan iso) {
  if (legacy.size != kCsProtectionBytes) {
    return std::nullopt;
  }
  const std::uint16_t password = load_u16(legacy.data);
  const bool content = load_u32(legacy.data + 2) != 0U;
  const bool objects = load_u32(legacy.data + 6) != 0U;
  std::string xml;
  if (password != 0U) {
    if (iso.size > 0U) {
      return std::nullopt;
    }
    xml.append(" password=\"");
    for (int shift = 12; shift >= 0; shift -= 4) {
      xml.push_back(kHexDigits[(password >> shift) & 0x0FU]);
    }
    xml.push_back('"');
  }
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
  std::uint16_t password = 0;
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
    } else if (name == "password") {
      const std::optional<std::uint32_t> legacy = parse_hex(attr.value(), 4U);
      if (!legacy || *legacy == 0U) {
        return false;
      }
      password = static_cast<std::uint16_t>(*legacy);
    } else {
      return false;
    }
  }
  if (password != 0U && (!algorithm.empty() || !hash.empty() || !salt.empty())) {
    return false;
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
  emit_u16(payload, password);
  emit_u32(payload, content);
  emit_u32(payload, objects);
  emit_record(dst, kBrtCsProtection, payload);
  return true;
}

/// One attribute of a start tag: its raw text with the whitespace before it, its name and its value.
struct TagAttribute {
  std::string_view raw;
  std::string_view name;
  std::string_view value;
};

/// The first `<sheetView` start tag of a chartsheet part (any namespace prefix), split into its attributes.
struct ViewTag {
  std::size_t begin = 0;
  std::size_t end = 0;
  std::string_view head;
  std::vector<TagAttribute> attributes;
  std::string_view tail;
};

std::optional<ViewTag> first_view_tag(std::string_view xml) {
  constexpr std::string_view kNameEnd = " \t\r\n/>";
  for (std::size_t at = xml.find('<'); at != std::string_view::npos; at = xml.find('<', at + 1U)) {
    const std::size_t name_end = xml.find_first_of(kNameEnd, at + 1U);
    if (name_end == std::string_view::npos) {
      return std::nullopt;
    }
    std::string_view name = xml.substr(at + 1U, name_end - at - 1U);
    if (const std::size_t colon = name.find(':'); colon != std::string_view::npos) {
      name.remove_prefix(colon + 1U);
    }
    if (name != "sheetView") {
      continue;
    }
    ViewTag tag;
    tag.begin = at;
    tag.head = xml.substr(at, name_end - at);
    std::size_t pos = name_end;
    while (true) {
      const std::size_t start = pos;
      pos = xml.find_first_not_of(" \t\r\n", pos);
      if (pos == std::string_view::npos) {
        return std::nullopt;
      }
      if (xml[pos] == '/' || xml[pos] == '>') {
        const std::size_t close = xml.find('>', pos);
        if (close == std::string_view::npos) {
          return std::nullopt;
        }
        tag.tail = xml.substr(start, close + 1U - start);
        tag.end = close + 1U;
        return tag;
      }
      const std::size_t eq = xml.find('=', pos);
      const std::size_t quote = eq == std::string_view::npos ? eq : xml.find_first_of("\"'", eq);
      const std::size_t value_end = quote == std::string_view::npos ? quote : xml.find(xml[quote], quote + 1U);
      if (value_end == std::string_view::npos) {
        return std::nullopt;
      }
      const std::size_t name_stop = xml.find_first_of(" \t\r\n=", pos);
      tag.attributes.push_back(TagAttribute{xml.substr(start, value_end + 1U - start), xml.substr(pos, name_stop - pos),
                                            xml.substr(quote + 1U, value_end - quote - 1U)});
      pos = value_end + 1U;
    }
  }
  return std::nullopt;
}

// A stored zoom, or the default when it is outside the 10..400 a view can hold, as a worksheet view reads it.
std::uint32_t valid_zoom(std::uint32_t zoom) {
  return zoom >= 10U && zoom <= 400U ? zoom : kDefaultZoom;
}

bool xsd_true(std::string_view value) {
  return value == "1" || value == "true";
}

ChartsheetView view_of(const ViewTag& tag) {
  ChartsheetView view;
  for (const TagAttribute& attr : tag.attributes) {
    if (attr.name == "zoomScale") {
      std::size_t pos = 0;
      std::uint32_t value = 0;
      const bool whole = a1::parse_uint(attr.value, &pos, &value) && pos == attr.value.size();
      view.zoom_scale = whole ? valid_zoom(value) : kDefaultZoom;
    } else if (attr.name == "zoomToFit") {
      view.zoom_to_fit = xsd_true(attr.value);
    } else if (attr.name == "tabSelected") {
      view.tab_selected = xsd_true(attr.value);
    }
  }
  return view;
}

/// The first BrtBeginCsView of `part` that carries the whole view, or nullopt.
std::optional<ByteSpan> first_cs_view(ByteSpan part) {
  ByteSpan cursor = part;
  while (cursor.size > 0) {
    Expected<XlsbRecord, Error> rec = read_record(cursor);
    if (!rec) {
      return std::nullopt;
    }
    if (rec.value().type == kBrtBeginCsView && rec.value().payload.size >= kCsViewBytes) {
      return rec.value().payload;
    }
  }
  return std::nullopt;
}

}  // namespace

std::optional<ChartsheetView> chartsheet_view_from_xml(std::string_view xml) {
  const std::optional<ViewTag> tag = first_view_tag(xml);
  if (!tag) {
    return std::nullopt;
  }
  return view_of(*tag);
}

std::optional<ChartsheetView> chartsheet_view_from_records(ByteSpan part) {
  const std::optional<ByteSpan> record = first_cs_view(part);
  if (!record) {
    return std::nullopt;
  }
  const std::uint16_t flags = load_u16(record->data);
  ChartsheetView view;
  view.tab_selected = (flags & kCsViewTabSelected) != 0U;
  view.zoom_to_fit = (flags & kCsViewZoomToFit) != 0U;
  view.zoom_scale = valid_zoom(load_u32(record->data + 2));
  return view;
}

std::string chartsheet_xml_with_view(std::string_view xml, ChartsheetView view) {
  const std::optional<ViewTag> tag = first_view_tag(xml);
  if (!tag || view_of(*tag) == view) {
    return std::string(xml);
  }
  // tabSelected and zoomScale go where the first of them or workbookViewId stood, zoomToFit last: Excel's order.
  std::string lead;
  if (view.tab_selected) {
    lead.append(" tabSelected=\"1\"");
  }
  if (view.zoom_scale != kDefaultZoom) {
    append_xml_attr_uint(lead, "zoomScale", view.zoom_scale);
  }
  std::string out(xml.substr(0, tag->begin));
  out.append(tag->head);
  bool lead_placed = false;
  for (const TagAttribute& attr : tag->attributes) {
    const bool restated = attr.name == "tabSelected" || attr.name == "zoomScale";
    if (attr.name == "zoomToFit") {
      continue;
    }
    if (!lead_placed && (restated || attr.name == "workbookViewId")) {
      out.append(lead);
      lead_placed = true;
    }
    if (!restated) {
      out.append(attr.raw);
    }
  }
  if (!lead_placed) {
    out.append(lead);
  }
  if (view.zoom_to_fit) {
    out.append(" zoomToFit=\"1\"");
  }
  out.append(tag->tail);
  out.append(xml.substr(tag->end));
  return out;
}

std::vector<std::uint8_t> chartsheet_records_with_view(ByteSpan part, ChartsheetView view) {
  std::vector<std::uint8_t> out(part.data, part.data + part.size);
  if (const std::optional<ByteSpan> record = first_cs_view(part)) {
    std::uint8_t* at = out.data() + (record->data - part.data);
    std::uint16_t flags = load_u16(record->data) & static_cast<std::uint16_t>(~(kCsViewTabSelected | kCsViewZoomToFit));
    flags |= view.tab_selected ? kCsViewTabSelected : 0U;
    flags |= view.zoom_to_fit ? kCsViewZoomToFit : 0U;
    store_u16(at, flags);
    store_u32(at + 2, view.zoom_scale);
  }
  return out;
}

void apply_chartsheet_part(std::string_view xml, Sheet& sheet) {
  if (const std::optional<ChartsheetView> view = chartsheet_view_from_xml(xml)) {
    SheetView& model = sheet.mutable_view();
    model.zoom_scale = view->zoom_scale;
    model.zoom_to_fit = view->zoom_to_fit;
    model.tab_selected = view->tab_selected;
  }
  pugi::xml_document doc;
  if (doc.load_buffer(xml.data(), xml.size()) && local_name(doc.document_element()) == "chartsheet") {
    ooxml::read_print_settings(doc.document_element(), sheet.mutable_print_settings());
  }
}

ChartsheetView chartsheet_view_of(const SheetView& view) {
  return ChartsheetView{view.zoom_scale, view.zoom_to_fit, view.tab_selected};
}

std::vector<PassthroughPart> restated_chartsheet_parts(const Workbook& wb) {
  std::vector<PassthroughPart> out;
  const PassthroughPartIndex index(wb.passthrough_parts());
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const Sheet& sheet = wb.sheet(i);
    const PassthroughPart* part =
        sheet.kind() == SheetKind::kChartsheet ? index.find(sheet.opaque_ooxml_part_path()) : nullptr;
    if (part == nullptr) {
      continue;
    }
    const ChartsheetView view = chartsheet_view_of(sheet.view());
    const ByteSpan bytes{part->bytes.data(), part->bytes.size()};
    if (strings::ends_with(part->path, ".bin")) {
      const std::optional<ChartsheetView> stored = chartsheet_view_from_records(bytes);
      if (stored && *stored != view) {
        out.push_back(*part);
        out.back().bytes = chartsheet_records_with_view(bytes, view);
      }
      continue;
    }
    const std::string_view xml(reinterpret_cast<const char*>(bytes.data), bytes.size);
    const std::optional<ChartsheetView> stored = chartsheet_view_from_xml(xml);
    if (stored && *stored != view) {
      const std::string rewritten = chartsheet_xml_with_view(xml, view);
      out.push_back(*part);
      out.back().bytes.assign(rewritten.begin(), rewritten.end());
    }
  }
  return out;
}

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

Expected<ChartsheetBin, Error> chartsheet_records_from_xml(std::string_view xml, const Theme& theme) {
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

  std::optional<std::vector<std::uint8_t>> prop = cs_prop_record(child_local(root, "sheetPr"), theme);
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
