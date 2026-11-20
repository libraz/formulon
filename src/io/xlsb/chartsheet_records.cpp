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
    } else if (name == "orientation" && std::string_view(attr.value()) == "portrait") {
      // Portrait is the clear bit.
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

// The hash fields of a BrtCsProtectionIso payload (spin count, content and objects flags, hash, salt, algorithm
// name); nullopt when it is truncated.
struct IsoHash {
  std::uint32_t spin_count = 0;
  std::string hash;
  std::string salt;
  std::string algorithm;
};

std::optional<IsoHash> parse_protection_iso(ByteSpan iso) {
  ByteSpan p = iso;
  auto spin = read_u32(p);
  auto iso_content = read_u32(p);
  auto iso_objects = read_u32(p);
  auto hash_size = read_u32(p);
  if (!spin || !iso_content || !iso_objects || !hash_size || hash_size.value() > p.size) {
    return std::nullopt;
  }
  IsoHash out;
  out.spin_count = spin.value();
  out.hash = base64_encode(p.data, hash_size.value());
  p.data += hash_size.value();
  p.size -= hash_size.value();
  auto salt_size = read_u32(p);
  if (!salt_size || salt_size.value() > p.size) {
    return std::nullopt;
  }
  out.salt = base64_encode(p.data, salt_size.value());
  p.data += salt_size.value();
  p.size -= salt_size.value();
  auto algorithm = read_xlwidestring(p);
  if (!algorithm || p.size != 0U) {
    return std::nullopt;
  }
  out.algorithm = std::move(algorithm.value());
  return out;
}

// A legacy password hash as the four hex digits `<sheetProtection password>` holds.
std::string legacy_password_text(std::uint16_t password) {
  std::string text;
  for (int shift = 12; shift >= 0; shift -= 4) {
    text.push_back(kHexDigits[(password >> shift) & 0x0FU]);
  }
  return text;
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
    xml.append(" password=\"").append(legacy_password_text(password)).push_back('"');
  }
  if (iso.size > 0U) {
    const std::optional<IsoHash> hash = parse_protection_iso(iso);
    if (!hash) {
      return std::nullopt;
    }
    append_xml_attr(xml, "algorithmName", hash->algorithm);
    append_xml_attr(xml, "hashValue", hash->hash);
    append_xml_attr(xml, "saltValue", hash->salt);
    append_xml_attr_uint(xml, "spinCount", hash->spin_count);
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

// The protection `<sheetProtection>` states; nullopt when it holds an attribute without a measured record form.
std::optional<ChartsheetProtection> stated_protection(const pugi::xml_node& protection) {
  ChartsheetProtection out;
  for (const pugi::xml_attribute attr : protection.attributes()) {
    const std::string_view name = attr.name();
    if (name == "content") {
      out.content_locked = attr.as_bool();
    } else if (name == "objects") {
      out.objects_locked = attr.as_bool();
    } else if (name == "algorithmName") {
      out.algorithm_name = attr.value();
    } else if (name == "hashValue") {
      out.hash_value = attr.value();
    } else if (name == "saltValue") {
      out.salt_value = attr.value();
    } else if (name == "spinCount") {
      out.spin_count = attr.as_uint();
    } else if (name == "password") {
      out.legacy_password_hash = attr.value();
    } else {
      return std::nullopt;
    }
  }
  return out;
}

// BrtCsProtectionIso (when a hash is set) and BrtCsProtection for `protection`; false, leaving `dst` as it was,
// when it states something without a measured record form.
bool emit_cs_protection(std::vector<std::uint8_t>& dst, const ChartsheetProtection& protection) {
  const std::uint32_t content = protection.content_locked ? 1U : 0U;
  const std::uint32_t objects = protection.objects_locked ? 1U : 0U;
  std::uint16_t password = 0;
  std::vector<std::uint8_t> hash;
  std::vector<std::uint8_t> salt;
  if (!protection.hash_value.empty() && !base64_decode(protection.hash_value, hash)) {
    return false;
  }
  if (!protection.salt_value.empty() && !base64_decode(protection.salt_value, salt)) {
    return false;
  }
  if (!protection.legacy_password_hash.empty()) {
    const std::optional<std::uint32_t> legacy = parse_hex(protection.legacy_password_hash, 4U);
    if (!legacy || *legacy == 0U) {
      return false;
    }
    password = static_cast<std::uint16_t>(*legacy);
  }
  const std::string& algorithm = protection.algorithm_name;
  if (password != 0U && (!algorithm.empty() || !hash.empty() || !salt.empty())) {
    return false;
  }
  const bool hashed = !algorithm.empty() || !hash.empty() || !salt.empty();
  if (hashed && (algorithm.empty() || hash.empty())) {
    return false;
  }
  if (hashed) {
    std::vector<std::uint8_t> payload;
    emit_u32(payload, protection.spin_count);
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

// `<sheetProtection>` stating `protection`.
std::string protection_element(const ChartsheetProtection& protection) {
  std::string xml = "<sheetProtection";
  if (!protection.legacy_password_hash.empty()) {
    append_xml_attr(xml, "password", protection.legacy_password_hash);
  }
  if (!protection.algorithm_name.empty()) {
    append_xml_attr(xml, "algorithmName", protection.algorithm_name);
  }
  if (!protection.hash_value.empty()) {
    append_xml_attr(xml, "hashValue", protection.hash_value);
  }
  if (!protection.salt_value.empty()) {
    append_xml_attr(xml, "saltValue", protection.salt_value);
  }
  if (protection.spin_count != 0U) {
    append_xml_attr_uint(xml, "spinCount", protection.spin_count);
  }
  if (protection.content_locked) {
    xml.append(" content=\"1\"");
  }
  if (protection.objects_locked) {
    xml.append(" objects=\"1\"");
  }
  xml.append("/>");
  return xml;
}

// `<pageSetup>` as a chartsheet states it: Excel always writes `orientation`, right after `paperSize`.
std::string page_setup_with_orientation(const std::string& fragment) {
  pugi::xml_document doc;
  if (fragment.empty() || !doc.load_buffer(fragment.data(), fragment.size()) || !doc.document_element() ||
      doc.document_element().attribute("orientation")) {
    return fragment;
  }
  pugi::xml_node node = doc.document_element();
  if (const pugi::xml_attribute paper = node.attribute("paperSize")) {
    node.insert_attribute_after("orientation", paper).set_value("portrait");
  } else {
    node.prepend_attribute("orientation").set_value("portrait");
  }
  return raw_xml(node);
}

// BrtCsPageSetup for a `<pageSetup>` fragment; nullopt when it does not parse or states something without a
// measured record form.
std::optional<std::vector<std::uint8_t>> cs_page_setup_record_of(const std::string& fragment) {
  pugi::xml_document doc;
  if (!doc.load_buffer(fragment.data(), fragment.size()) || !doc.document_element()) {
    return std::nullopt;
  }
  return cs_page_setup_record(doc.document_element());
}

/// One attribute of a start tag: its raw text with the whitespace before it, its name and its value.
struct TagAttribute {
  std::string_view raw;
  std::string_view name;
  std::string_view value;
};

/// The first start tag of an element in a chartsheet part (any namespace prefix), split into its attributes.
struct StartTag {
  std::size_t begin = 0;
  std::size_t end = 0;
  std::string_view head;
  std::vector<TagAttribute> attributes;
  std::string_view tail;
};

std::optional<StartTag> first_start_tag(std::string_view xml, std::string_view element) {
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
    if (name != element) {
      continue;
    }
    StartTag tag;
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

ChartsheetView view_of(const StartTag& tag) {
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

/// The payload of the first `type` record of `part` at least `min_bytes` long, or nullopt.
std::optional<ByteSpan> first_record(ByteSpan part, std::uint16_t type, std::size_t min_bytes) {
  ByteSpan cursor = part;
  while (cursor.size > 0) {
    Expected<XlsbRecord, Error> rec = read_record(cursor);
    if (!rec) {
      return std::nullopt;
    }
    if (rec.value().type == type && rec.value().payload.size >= min_bytes) {
      return rec.value().payload;
    }
  }
  return std::nullopt;
}

// The tab colour of a BrtCsProp payload at least `kCsPropColorAt + 8` bytes long, read as a worksheet's BrtWsProp
// colour is, whatever else the record states; none for the automatic colour.
ColorSpec cs_prop_tab_color(ByteSpan payload) {
  const std::uint8_t* color = payload.data + kCsPropColorAt;
  const std::uint32_t argb = pack_argb(color[7], color[4], color[5], color[6]);
  ColorSpec tab;
  if (!decode_brt_color_selector(static_cast<std::uint8_t>(color[0] >> 1U), color[1], load_u16(color + 2), argb, tab) ||
      (tab.kind == ColorSpec::Kind::kIndexed && color[1] == kBrtColorAutoIndex)) {
    return ColorSpec{};
  }
  return tab;
}

bool same_tab_color(const ColorSpec& a, const ColorSpec& b) {
  if (a.kind != b.kind) {
    return false;
  }
  switch (a.kind) {
    case ColorSpec::Kind::kRgb:
      return a.rgb == b.rgb;
    case ColorSpec::Kind::kTheme:
      return a.theme == b.theme && a.tint == b.tint;
    case ColorSpec::Kind::kIndexed:
      return a.indexed == b.indexed;
    case ColorSpec::Kind::kAuto:
    case ColorSpec::Kind::kNone:
      break;
  }
  return true;
}

// `<tabColor>` stating `color`, with the namespace prefix (`x:` or empty) of the element holding it.
std::string tab_color_xml(const ColorSpec& color, std::string_view prefix) {
  std::string xml = "<";
  xml.append(prefix);
  xml.append("tabColor");
  switch (color.kind) {
    case ColorSpec::Kind::kRgb:
      xml.append(" rgb=\"");
      for (int shift = 28; shift >= 0; shift -= 4) {
        xml.push_back(kHexDigits[(color.rgb >> static_cast<unsigned>(shift)) & 0xFU]);
      }
      xml.push_back('"');
      break;
    case ColorSpec::Kind::kTheme:
      append_xml_attr_uint(xml, "theme", color.theme);
      if (color.tint != 0.0) {
        append_xml_attr_number(xml, "tint", color.tint);
      }
      break;
    case ColorSpec::Kind::kIndexed:
      append_xml_attr_uint(xml, "indexed", color.indexed);
      break;
    case ColorSpec::Kind::kAuto:
    case ColorSpec::Kind::kNone:
      xml.append(" auto=\"1\"");
      break;
  }
  return xml + "/>";
}

// The namespace prefix of a start tag's element, with its colon; empty when it has none.
std::string_view tag_prefix(const StartTag& tag) {
  const std::string_view name = tag.head.substr(1U);
  const std::size_t colon = name.find(':');
  return colon == std::string_view::npos ? std::string_view() : name.substr(0, colon + 1U);
}

}  // namespace

std::optional<ChartsheetView> chartsheet_view_from_xml(std::string_view xml) {
  const std::optional<StartTag> tag = first_start_tag(xml, "sheetView");
  if (!tag) {
    return std::nullopt;
  }
  return view_of(*tag);
}

std::optional<ChartsheetView> chartsheet_view_from_records(ByteSpan part) {
  const std::optional<ByteSpan> record = first_record(part, kBrtBeginCsView, kCsViewBytes);
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

std::optional<ChartsheetProtection> chartsheet_protection_from_xml(std::string_view xml) {
  pugi::xml_document doc;
  if (!doc.load_buffer(xml.data(), xml.size())) {
    return std::nullopt;
  }
  const pugi::xml_node root = doc.document_element();
  if (local_name(root) != "chartsheet") {
    return std::nullopt;
  }
  for (const pugi::xml_node child : root.children()) {
    if (child.type() != pugi::node_element || local_name(child) != "sheetProtection") {
      continue;
    }
    ChartsheetProtection out;
    out.algorithm_name = child.attribute("algorithmName").value();
    out.hash_value = child.attribute("hashValue").value();
    out.salt_value = child.attribute("saltValue").value();
    out.spin_count = child.attribute("spinCount").as_uint();
    out.legacy_password_hash = child.attribute("password").value();
    out.content_locked = child.attribute("content").as_bool();
    out.objects_locked = child.attribute("objects").as_bool();
    return out;
  }
  return std::nullopt;
}

std::optional<ChartsheetProtection> chartsheet_protection_from_records(ByteSpan part) {
  std::optional<ByteSpan> legacy;
  std::optional<ByteSpan> iso;
  ByteSpan cursor = part;
  while (cursor.size > 0) {
    Expected<XlsbRecord, Error> rec = read_record(cursor);
    if (!rec) {
      return std::nullopt;
    }
    if (rec.value().type == kBrtCsProtectionIso) {
      iso = rec.value().payload;
    } else if (rec.value().type == kBrtCsProtection && rec.value().payload.size == kCsProtectionBytes) {
      legacy = rec.value().payload;
      break;
    }
  }
  if (!legacy) {
    return std::nullopt;
  }
  const std::uint16_t password = load_u16(legacy->data);
  ChartsheetProtection out;
  out.content_locked = load_u32(legacy->data + 2) != 0U;
  out.objects_locked = load_u32(legacy->data + 6) != 0U;
  if (password != 0U) {
    out.legacy_password_hash = legacy_password_text(password);
  }
  if (iso && iso->size > 0U) {
    if (const std::optional<IsoHash> hash = parse_protection_iso(*iso)) {
      out.algorithm_name = hash->algorithm;
      out.hash_value = hash->hash;
      out.salt_value = hash->salt;
      out.spin_count = hash->spin_count;
    }
  } else if (password == 0U && !out.content_locked && !out.objects_locked) {
    return std::nullopt;
  }
  return out;
}

std::string chartsheet_xml_with_view(std::string_view xml, ChartsheetView view) {
  const std::optional<StartTag> tag = first_start_tag(xml, "sheetView");
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
  if (const std::optional<ByteSpan> record = first_record(part, kBrtBeginCsView, kCsViewBytes)) {
    std::uint8_t* at = out.data() + (record->data - part.data);
    std::uint16_t flags = load_u16(record->data) & static_cast<std::uint16_t>(~(kCsViewTabSelected | kCsViewZoomToFit));
    flags |= view.tab_selected ? kCsViewTabSelected : 0U;
    flags |= view.zoom_to_fit ? kCsViewZoomToFit : 0U;
    store_u16(at, flags);
    store_u32(at + 2, view.zoom_scale);
  }
  return out;
}

std::string chartsheet_xml_with_tab_color(std::string_view xml, const ColorSpec& color) {
  pugi::xml_document doc;
  if (!doc.load_buffer(xml.data(), xml.size()) || local_name(doc.document_element()) != "chartsheet" ||
      same_tab_color(ooxml::read_tab_color(doc.document_element().child("sheetPr")), color)) {
    return std::string(xml);
  }
  const std::optional<StartTag> sheet_pr = first_start_tag(xml, "sheetPr");
  if (!sheet_pr) {
    // `<sheetPr>` leads the chartsheet's content.
    const std::optional<StartTag> root = first_start_tag(xml, "chartsheet");
    if (!root || root->tail.find('/') != std::string_view::npos) {
      return std::string(xml);
    }
    const std::string_view prefix = tag_prefix(*root);
    std::string out(xml.substr(0, root->end));
    out.append("<").append(prefix).append("sheetPr>").append(tab_color_xml(color, prefix));
    out.append("</").append(prefix).append("sheetPr>");
    out.append(xml.substr(root->end));
    return out;
  }
  const std::string_view name = sheet_pr->head.substr(1U);
  std::size_t rest = sheet_pr->end;
  if (sheet_pr->tail.find('/') == std::string_view::npos) {
    const std::size_t close = xml.find("</" + std::string(name), sheet_pr->end);
    const std::size_t close_end = close == std::string_view::npos ? close : xml.find('>', close);
    if (close_end == std::string_view::npos) {
      return std::string(xml);
    }
    rest = close_end + 1U;
  }
  // `<tabColor>` is the only child `<sheetPr>` holds here; without it the element closes empty, as Excel writes it.
  std::string out(xml.substr(0, sheet_pr->begin));
  out.append(sheet_pr->head);
  for (const TagAttribute& attr : sheet_pr->attributes) {
    out.append(attr.raw);
  }
  if (color.kind == ColorSpec::Kind::kNone) {
    out.append("/>");
  } else {
    out.append(">").append(tab_color_xml(color, tag_prefix(*sheet_pr))).append("</").append(name).append(">");
  }
  out.append(xml.substr(rest));
  return out;
}

std::vector<std::uint8_t> chartsheet_records_with_tab_color(ByteSpan part, const ColorSpec& color, const Theme& theme) {
  std::vector<std::uint8_t> out(part.data, part.data + part.size);
  const std::optional<ByteSpan> record = first_record(part, kBrtCsProp, kCsPropColorAt + 8U);
  if (!record) {
    return out;
  }
  if (!same_tab_color(cs_prop_tab_color(*record), color)) {
    const std::array<std::uint8_t, 8> bytes = encode_tab_color(color, theme);
    std::copy(bytes.begin(), bytes.end(), out.begin() + (record->data - part.data) + kCsPropColorAt);
  }
  return out;
}

std::optional<ColorSpec> chartsheet_tab_color_from_records(ByteSpan part) {
  const std::optional<ByteSpan> record = first_record(part, kBrtCsProp, kCsPropColorAt + 8U);
  if (!record) {
    return std::nullopt;
  }
  return cs_prop_tab_color(*record);
}

namespace {

// CT_Chartsheet's children in schema order.
constexpr std::array<std::string_view, 14> kChartsheetChildren = {
    "sheetPr", "sheetViews",    "sheetProtection", "customSheetViews", "pageMargins", "pageSetup",       "headerFooter",
    "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF",        "picture",     "webPublishItems", "extLst"};

// `xml` with the chartsheet child `name` replaced by `text`, dropped when `text` is empty, or inserted at its schema
// position when absent. Every byte outside the element is kept.
std::string replace_chartsheet_element(std::string_view xml, std::string_view name, std::string_view text) {
  if (const std::optional<StartTag> tag = first_start_tag(xml, name)) {
    std::size_t end = tag->end;
    if (tag->tail.find('/') == std::string_view::npos) {
      const std::size_t close = xml.find("</" + std::string(tag->head.substr(1)), tag->end);
      const std::size_t close_end = close == std::string_view::npos ? close : xml.find('>', close);
      if (close_end == std::string_view::npos) {
        return std::string(xml);
      }
      end = close_end + 1U;
    }
    std::string out(xml.substr(0, tag->begin));
    out.append(text).append(xml.substr(end));
    return out;
  }
  const std::optional<StartTag> root = first_start_tag(xml, "chartsheet");
  if (text.empty() || !root || !tag_prefix(*root).empty() || root->tail.find('/') != std::string_view::npos) {
    return std::string(xml);
  }
  std::size_t at = std::string_view::npos;
  const auto self = std::find(kChartsheetChildren.begin(), kChartsheetChildren.end(), name);
  for (auto later = self + 1; later < kChartsheetChildren.end() && at == std::string_view::npos; ++later) {
    if (const std::optional<StartTag> next = first_start_tag(xml, *later)) {
      at = next->begin;
    }
  }
  if (at == std::string_view::npos) {
    at = xml.rfind("</");
    if (at == std::string_view::npos || xml.substr(at).rfind("</chartsheet", 0) != 0U) {
      return std::string(xml);
    }
  }
  std::string out(xml.substr(0, at));
  out.append(text).append(xml.substr(at));
  return out;
}

// Positions of the record groups that follow BrtEndCsViews; a group's records are replaced together.
constexpr int kRankViewsEnd = 1;
constexpr int kRankProtection = 2;
constexpr int kRankMargins = 3;
constexpr int kRankPageSetup = 4;
constexpr int kRankHeaderFooter = 5;

// The position of a chartsheet record in the settings run that follows BrtEndCsViews; 0 for any other record.
int settings_rank(std::uint16_t type) {
  switch (type) {
    case kBrtEndCsViews:
      return kRankViewsEnd;
    case kBrtCsProtectionIso:
    case kBrtCsProtection:
      return kRankProtection;
    case kBrtMargins:
      return kRankMargins;
    case kBrtCsPageSetup:
      return kRankPageSetup;
    case kBrtBeginHeaderFooter:
    case kBrtEndHeaderFooter:
      return kRankHeaderFooter;
    default:
      return 0;
  }
}

struct RecordFrame {
  std::uint16_t type = 0;
  std::size_t begin = 0;
  std::size_t end = 0;
};

}  // namespace

std::string chartsheet_xml_with_settings(std::string_view xml, const Sheet& sheet,
                                         const std::optional<ChartsheetProtection>& stored_protection) {
  pugi::xml_document doc;
  if (!doc.load_buffer(xml.data(), xml.size()) || local_name(doc.document_element()) != "chartsheet") {
    return std::string(xml);
  }
  SheetPrintSettings stored;
  ooxml::read_print_settings(doc.document_element(), stored);
  const SheetPrintSettings& model = sheet.print_settings();
  std::string out(xml);
  if (model.page_margins_xml != stored.page_margins_xml) {
    out = replace_chartsheet_element(out, "pageMargins", model.page_margins_xml);
  }
  if (model.page_setup_xml != stored.page_setup_xml) {
    out = replace_chartsheet_element(out, "pageSetup", page_setup_with_orientation(model.page_setup_xml));
  }
  if (model.header_footer_xml != stored.header_footer_xml) {
    out = replace_chartsheet_element(out, "headerFooter", model.header_footer_xml);
  }
  if (sheet.chartsheet_protection() != stored_protection) {
    const std::optional<ChartsheetProtection>& protection = sheet.chartsheet_protection();
    out = replace_chartsheet_element(out, "sheetProtection", protection ? protection_element(*protection) : "");
  }
  return out;
}

std::vector<std::uint8_t> chartsheet_records_with_settings(ByteSpan part, const Sheet& sheet) {
  const std::vector<std::uint8_t> unchanged(part.data, part.data + part.size);
  std::vector<RecordFrame> frames;
  ByteSpan cursor = part;
  while (cursor.size > 0) {
    const std::size_t begin = part.size - cursor.size;
    Expected<XlsbRecord, Error> rec = read_record(cursor);
    if (!rec) {
      return unchanged;
    }
    frames.push_back(RecordFrame{rec.value().type, begin, part.size - cursor.size});
  }
  const Expected<ChartsheetXml, Error> converted = chartsheet_xml_from_records(part);
  pugi::xml_document doc;
  if (!converted || !doc.load_buffer(converted.value().xml.data(), converted.value().xml.size())) {
    return unchanged;
  }
  SheetPrintSettings stored;
  ooxml::read_print_settings(doc.document_element(), stored);
  const SheetPrintSettings& model = sheet.print_settings();

  // The replacement records of each settings group the model restates; an empty run removes the group.
  struct Group {
    int rank;
    std::vector<std::uint8_t> bytes;
  };
  std::vector<Group> groups;
  if (sheet.chartsheet_protection() != chartsheet_protection_from_records(part)) {
    // An unprotected sheet keeps its all-zero BrtCsProtection.
    std::vector<std::uint8_t> bytes;
    if (emit_cs_protection(bytes, sheet.chartsheet_protection().value_or(ChartsheetProtection{}))) {
      groups.push_back(Group{kRankProtection, std::move(bytes)});
    }
  }
  if (model.page_margins_xml != stored.page_margins_xml) {
    SheetPrintSettings margins;
    margins.page_margins_xml = model.page_margins_xml;
    Group group{kRankMargins, {}};
    emit_chartsheet_print_records(group.bytes, margins);
    groups.push_back(std::move(group));
  }
  if (model.page_setup_xml != stored.page_setup_xml) {
    Group group{kRankPageSetup, {}};
    bool encoded = true;
    if (!model.page_setup_xml.empty()) {
      const std::optional<std::vector<std::uint8_t>> record = cs_page_setup_record_of(model.page_setup_xml);
      encoded = record.has_value();
      if (encoded) {
        emit_record(group.bytes, kBrtCsPageSetup, *record);
      }
    }
    if (encoded) {
      groups.push_back(std::move(group));
    }
  }
  if (model.header_footer_xml != stored.header_footer_xml) {
    SheetPrintSettings header_footer;
    header_footer.header_footer_xml = model.header_footer_xml;
    Group group{kRankHeaderFooter, {}};
    emit_chartsheet_print_records(group.bytes, header_footer);
    groups.push_back(std::move(group));
  }
  if (groups.empty()) {
    return unchanged;
  }

  // A group the part lacks goes after the last record that precedes it in the run, else before the first that
  // follows it, else before BrtEndSheet.
  std::vector<std::vector<std::uint8_t>> inserted(frames.size() + 1U);
  std::vector<int> replaced_ranks;
  for (const Group& group : groups) {
    const auto present = std::any_of(frames.begin(), frames.end(), [&group](const RecordFrame& frame) {
      return settings_rank(frame.type) == group.rank;
    });
    if (present) {
      replaced_ranks.push_back(group.rank);
      continue;
    }
    std::size_t at = frames.size();
    bool placed = false;
    for (std::size_t i = frames.size(); i-- > 0 && !placed;) {
      const int rank = settings_rank(frames[i].type);
      if (rank != 0 && rank < group.rank) {
        at = i + 1U;
        placed = true;
      }
    }
    for (std::size_t i = 0; i < frames.size() && !placed; ++i) {
      const int rank = settings_rank(frames[i].type);
      if ((rank != 0 && rank > group.rank) || frames[i].type == kBrtEndSheet) {
        at = i;
        placed = true;
      }
    }
    inserted[at].insert(inserted[at].end(), group.bytes.begin(), group.bytes.end());
  }

  std::vector<std::uint8_t> out;
  std::vector<int> written;
  const auto group_of = [&groups](int rank) -> const Group* {
    for (const Group& group : groups) {
      if (group.rank == rank) {
        return &group;
      }
    }
    return nullptr;
  };
  for (std::size_t i = 0; i <= frames.size(); ++i) {
    out.insert(out.end(), inserted[i].begin(), inserted[i].end());
    if (i == frames.size()) {
      break;
    }
    const int rank = settings_rank(frames[i].type);
    const bool replaced = std::find(replaced_ranks.begin(), replaced_ranks.end(), rank) != replaced_ranks.end();
    if (!replaced) {
      out.insert(out.end(), part.data + frames[i].begin, part.data + frames[i].end);
    } else if (std::find(written.begin(), written.end(), rank) == written.end()) {
      const Group* group = group_of(rank);
      out.insert(out.end(), group->bytes.begin(), group->bytes.end());
      written.push_back(rank);
    }
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
  sheet.set_chartsheet_protection(chartsheet_protection_from_xml(xml));
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
  std::optional<LoadedTheme> theme;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const Sheet& sheet = wb.sheet(i);
    const PassthroughPart* part =
        sheet.kind() == SheetKind::kChartsheet ? index.find(sheet.opaque_ooxml_part_path()) : nullptr;
    if (part == nullptr) {
      continue;
    }
    const ChartsheetView view = chartsheet_view_of(sheet.view());
    const ColorSpec& tab_color = sheet.print_settings().tab_color;
    const ByteSpan bytes{part->bytes.data(), part->bytes.size()};
    std::vector<std::uint8_t> restated;
    if (strings::ends_with(part->path, ".bin")) {
      const std::optional<ChartsheetView> stored = chartsheet_view_from_records(bytes);
      restated = stored && *stored != view ? chartsheet_records_with_view(bytes, view) : part->bytes;
      if (!theme) {
        theme.emplace(wb.load_theme());
      }
      restated = chartsheet_records_with_tab_color(ByteSpan{restated.data(), restated.size()}, tab_color, theme->theme);
      restated = chartsheet_records_with_settings(ByteSpan{restated.data(), restated.size()}, sheet);
    } else {
      const std::string_view xml(reinterpret_cast<const char*>(bytes.data), bytes.size);
      const std::string rewritten =
          chartsheet_xml_with_settings(chartsheet_xml_with_tab_color(chartsheet_xml_with_view(xml, view), tab_color),
                                       sheet, chartsheet_protection_from_xml(xml));
      restated.assign(rewritten.begin(), rewritten.end());
    }
    if (restated != part->bytes) {
      out.push_back(*part);
      out.back().bytes = std::move(restated);
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
  const std::optional<ChartsheetProtection> stated = protection ? stated_protection(protection) : std::nullopt;
  if (!stated || !emit_cs_protection(dst, *stated)) {
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
