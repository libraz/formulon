#include "io/theme_part.h"

#include <array>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "utils/number_text.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon::io {
namespace {

constexpr const char* kThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
constexpr const char* kDefaultThemePath = "xl/theme/theme1.xml";

constexpr std::array<const char*, kThemeColorCount> kSchemeNames = {
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"};

/// Local part of a possibly prefixed element name.
std::string_view local_name(const char* qualified) {
  std::string_view name(qualified);
  const std::size_t colon = name.find(':');
  return colon == std::string_view::npos ? name : name.substr(colon + 1U);
}

/// Namespace prefix including the trailing colon, or empty.
std::string prefix_of(const char* qualified) {
  std::string_view name(qualified);
  const std::size_t colon = name.find(':');
  return colon == std::string_view::npos ? std::string() : std::string(name.substr(0U, colon + 1U));
}

pugi::xml_node find_child(const pugi::xml_node& parent, std::string_view local) {
  for (pugi::xml_node child : parent.children()) {
    if (child.type() == pugi::node_element && local_name(child.name()) == local) {
      return child;
    }
  }
  return pugi::xml_node();
}

pugi::xml_node clr_scheme_of(const pugi::xml_document& doc) {
  return find_child(find_child(doc.document_element(), "themeElements"), "clrScheme");
}

pugi::xml_node font_scheme_of(const pugi::xml_document& doc) {
  return find_child(find_child(doc.document_element(), "themeElements"), "fontScheme");
}

std::string hex6(std::uint32_t argb) {
  char buf[8];
  format_hex(buf, sizeof(buf), argb & 0xFFFFFFU, 6, true);
  return buf;
}

/// Reads one scheme slot (`a:sysClr lastClr` or `a:srgbClr val`).
std::optional<std::uint32_t> read_scheme_color(const pugi::xml_node& slot) {
  if (!slot) {
    return std::nullopt;
  }
  pugi::xml_node clr = slot.first_child();
  while (clr && clr.type() != pugi::node_element) {
    clr = clr.next_sibling();
  }
  if (!clr) {
    return std::nullopt;
  }
  const std::string_view kind = local_name(clr.name());
  const char* hex = nullptr;
  if (kind == "srgbClr") {
    hex = clr.attribute("val").value();
  } else if (kind == "sysClr") {
    hex = clr.attribute("lastClr").value();
  } else {
    return std::nullopt;
  }
  const std::uint32_t argb = parse_rgb_hex(hex, 0U);
  if (argb == 0U) {
    return std::nullopt;
  }
  return argb;
}

std::string font_face(const pugi::xml_node& font_node, bool east_asian) {
  if (!east_asian) {
    return find_child(font_node, "latin").attribute("typeface").value();
  }
  const std::string ea = find_child(font_node, "ea").attribute("typeface").value();
  if (!ea.empty()) {
    return ea;
  }
  for (pugi::xml_node child : font_node.children()) {
    if (child.type() == pugi::node_element && local_name(child.name()) == "font" &&
        std::string_view(child.attribute("script").value()) == "Jpan") {
      return child.attribute("typeface").value();
    }
  }
  return std::string();
}

std::optional<Theme> parse_theme(const pugi::xml_document& doc) {
  const pugi::xml_node clr = clr_scheme_of(doc);
  if (!clr) {
    return std::nullopt;
  }
  Theme out = default_theme();
  for (std::size_t i = 0; i < kThemeColorCount; ++i) {
    const std::optional<std::uint32_t> argb = read_scheme_color(find_child(clr, kSchemeNames[i]));
    if (!argb) {
      return std::nullopt;
    }
    out.colors[i] = *argb;
  }
  const pugi::xml_node fonts = font_scheme_of(doc);
  if (const pugi::xml_node major = find_child(fonts, "majorFont")) {
    out.fonts.major_latin = font_face(major, false);
    out.fonts.major_ea = font_face(major, true);
  }
  if (const pugi::xml_node minor = find_child(fonts, "minorFont")) {
    out.fonts.minor_latin = font_face(minor, false);
    out.fonts.minor_ea = font_face(minor, true);
  }
  return out;
}

/// Package path of the theme part: the workbook relationship's target when it
/// names one, otherwise the conventional path.
std::string theme_part_path(const Workbook& wb) {
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    if (!rel.target_external && rel.type == kRelTheme) {
      return rel.target;
    }
  }
  return kDefaultThemePath;
}

/// A minimal valid theme (colour scheme, font scheme, format scheme with three
/// entries per list) carrying the default colours and fonts.
std::string build_default_theme_xml() {
  const Theme& def = default_theme();
  std::string xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">"
      "<a:themeElements><a:clrScheme name=\"Office\">";
  for (std::size_t i = 0; i < kThemeColorCount; ++i) {
    xml += "<a:" + std::string(kSchemeNames[i]) + "><a:srgbClr val=\"" + hex6(def.colors[i]) +
           "\"/></a:" + kSchemeNames[i] + ">";
  }
  xml += "</a:clrScheme><a:fontScheme name=\"Office\">";
  const auto font = [&xml](const char* tag, const std::string& latin, const std::string& ea) {
    xml += std::string("<a:") + tag + "><a:latin typeface=\"" + latin +
           "\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/>" + "<a:font script=\"Jpan\" typeface=\"" + ea +
           "\"/></a:" + tag + ">";
  };
  font("majorFont", def.fonts.major_latin, def.fonts.major_ea);
  font("minorFont", def.fonts.minor_latin, def.fonts.minor_ea);
  xml += "</a:fontScheme><a:fmtScheme name=\"Office\">";
  const std::string fill = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
  const auto list = [&xml](const char* tag, const std::string& item) {
    xml += std::string("<a:") + tag + ">" + item + item + item + "</a:" + tag + ">";
  };
  list("fillStyleLst", fill);
  list("lnStyleLst", "<a:ln w=\"6350\">" + fill + "</a:ln>");
  list("effectStyleLst", "<a:effectStyle><a:effectLst/></a:effectStyle>");
  list("bgFillStyleLst", fill);
  xml += "</a:fmtScheme></a:themeElements></a:theme>";
  return xml;
}

/// Parses the workbook's theme part for editing, generating and registering a
/// default one when absent. Yields the part path and the parsed document.
Expected<std::string, Error> open_theme_for_edit(Workbook& wb, pugi::xml_document& doc) {
  std::string path = theme_part_path(wb);
  if (ooxml::find_passthrough_part(wb, path) == nullptr) {
    const std::string xml = build_default_theme_xml();
    PassthroughPart part(path, kThemeContentType, std::vector<std::uint8_t>(xml.begin(), xml.end()));
    RETURN_IF_ERROR(wb.add_passthrough_part(std::move(part)));
    wb.add_workbook_relationship(std::string(kRelTheme), path);
  }
  RETURN_IF_ERROR(ooxml::load_part_dom(wb, path, "theme", doc));
  if (!parse_theme(doc)) {
    return make_error(FormulonErrorCode::kIoXmlParse, "theme: part is not a parseable theme", "part=" + path);
  }
  return path;
}

void set_attr(pugi::xml_node node, const char* name, const std::string& value) {
  pugi::xml_attribute attr = node.attribute(name);
  if (!attr) {
    attr = node.append_attribute(name);
  }
  attr.set_value(value.c_str());
}

/// Writes the Latin and East Asian faces of one `a:majorFont` / `a:minorFont`.
void patch_font(pugi::xml_node font_node, const std::string& latin, const std::string& east_asian) {
  const std::string prefix = prefix_of(font_node.name());
  pugi::xml_node latin_node = find_child(font_node, "latin");
  if (!latin_node) {
    latin_node = font_node.prepend_child((prefix + "latin").c_str());
  }
  set_attr(latin_node, "typeface", latin);
  pugi::xml_node ea_node = find_child(font_node, "ea");
  if (ea_node && *ea_node.attribute("typeface").value() != '\0') {
    set_attr(ea_node, "typeface", east_asian);
    return;
  }
  for (pugi::xml_node child : font_node.children()) {
    if (child.type() == pugi::node_element && local_name(child.name()) == "font" &&
        std::string_view(child.attribute("script").value()) == "Jpan") {
      set_attr(child, "typeface", east_asian);
      return;
    }
  }
  pugi::xml_node jpan = font_node.append_child((prefix + "font").c_str());
  set_attr(jpan, "script", "Jpan");
  set_attr(jpan, "typeface", east_asian);
}

}  // namespace

LoadedTheme load_theme(const Workbook& wb) {
  LoadedTheme out;
  out.theme = default_theme();
  const PassthroughPart* part = ooxml::find_passthrough_part(wb, theme_part_path(wb));
  if (part == nullptr) {
    return out;
  }
  out.source = ThemeSource::kUnparseable;
  pugi::xml_document doc;
  if (!load_xml_buffer(doc, part->bytes, "theme", part->path)) {
    return out;
  }
  if (std::optional<Theme> parsed = parse_theme(doc)) {
    out.theme = std::move(*parsed);
    out.source = ThemeSource::kPart;
  }
  return out;
}

Expected<void, Error> set_theme_colors(Workbook& wb, const ThemeColors& colors) {
  pugi::xml_document doc;
  auto path_or = open_theme_for_edit(wb, doc);
  if (!path_or) {
    return Expected<void, Error>(path_or.error());
  }
  const pugi::xml_node clr = clr_scheme_of(doc);
  const std::string prefix = prefix_of(clr.name());
  for (std::size_t i = 0; i < kThemeColorCount; ++i) {
    pugi::xml_node slot = find_child(clr, kSchemeNames[i]);
    while (slot.first_child()) {
      slot.remove_child(slot.first_child());
    }
    set_attr(slot.append_child((prefix + "srgbClr").c_str()), "val", hex6(colors[i]));
  }
  return ooxml::store_part_dom(wb, path_or.value(), doc);
}

Expected<void, Error> set_theme_fonts(Workbook& wb, const ThemeFonts& fonts) {
  pugi::xml_document doc;
  auto path_or = open_theme_for_edit(wb, doc);
  if (!path_or) {
    return Expected<void, Error>(path_or.error());
  }
  const pugi::xml_node scheme = font_scheme_of(doc);
  const pugi::xml_node major = find_child(scheme, "majorFont");
  const pugi::xml_node minor = find_child(scheme, "minorFont");
  if (!major || !minor) {
    return make_error(FormulonErrorCode::kIoXmlParse, "theme: font scheme has no major/minor font",
                      "part=" + path_or.value());
  }
  patch_font(major, fonts.major_latin, fonts.major_ea);
  patch_font(minor, fonts.minor_latin, fonts.minor_ea);
  return ooxml::store_part_dom(wb, path_or.value(), doc);
}

}  // namespace formulon::io
