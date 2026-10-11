#include "io/theme_part.h"

#include <algorithm>
#include <array>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/part_dom.h"
#include "io/ooxml/part_graph.h"
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

/// The `a:font script typeface` entries of one `a:majorFont` / `a:minorFont`, in file order.
std::vector<ThemeScriptFont> script_fonts(const pugi::xml_node& font_node) {
  std::vector<ThemeScriptFont> out;
  for (pugi::xml_node child : font_node.children()) {
    if (child.type() == pugi::node_element && local_name(child.name()) == "font") {
      out.push_back({child.attribute("script").value(), child.attribute("typeface").value()});
    }
  }
  return out;
}

/// The default Office script faces, for a part that lacks the font block.
std::vector<ThemeScriptFont> office_script_fonts(bool major) {
  std::vector<ThemeScriptFont> out;
  for (std::size_t i = 0; i < kOfficeScriptCount; ++i) {
    out.push_back({office_script(i), office_script_face(i, major)});
  }
  return out;
}

std::optional<Theme> parse_theme(const pugi::xml_document& doc) {
  const pugi::xml_node clr = clr_scheme_of(doc);
  if (!clr) {
    return std::nullopt;
  }
  Theme out = default_theme();
  out.major_scripts = office_script_fonts(true);
  out.minor_scripts = office_script_fonts(false);
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
    out.major_scripts = script_fonts(major);
  }
  if (const pugi::xml_node minor = find_child(fonts, "minorFont")) {
    out.fonts.minor_latin = font_face(minor, false);
    out.fonts.minor_ea = font_face(minor, true);
    out.minor_scripts = script_fonts(minor);
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
/// entries per list) carrying the default colours, fonts and script fonts.
std::string build_default_theme_xml() {
  const Theme& def = default_theme();
  std::string xml = std::string(kXmlDecl) +
                    "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">"
                    "<a:themeElements><a:clrScheme name=\"Office\">";
  for (std::size_t i = 0; i < kThemeColorCount; ++i) {
    xml.append("<a:").append(kSchemeNames[i]).append("><a:srgbClr val=\"").append(hex6(def.colors[i]));
    xml.append("\"/></a:").append(kSchemeNames[i]).append(">");
  }
  xml += "</a:clrScheme><a:fontScheme name=\"Office\">";
  for (const bool major : {true, false}) {
    const char* tag = major ? "majorFont" : "minorFont";
    xml.append("<a:").append(tag).append("><a:latin typeface=\"");
    xml.append(major ? def.fonts.major_latin : def.fonts.minor_latin);
    xml.append("\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/>");
    for (std::size_t i = 0; i < kOfficeScriptCount; ++i) {
      xml.append("<a:font script=\"").append(office_script(i)).append("\" typeface=\"");
      xml.append(office_script_face(i, major)).append("\"/>");
    }
    xml.append("</a:").append(tag).append(">");
  }
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

void set_attr(pugi::xml_node node, const char* name, std::string_view value) {
  pugi::xml_attribute attr = node.attribute(name);
  if (!attr) {
    attr = node.append_attribute(name);
  }
  attr.set_value(value.data(), value.size());
}

/// Sets the face of the `a:font` entry for `script`, appending the entry when absent.
void set_script_face(pugi::xml_node font_node, std::string_view script, std::string_view typeface) {
  for (pugi::xml_node child : font_node.children()) {
    if (child.type() == pugi::node_element && local_name(child.name()) == "font" &&
        std::string_view(child.attribute("script").value()) == script) {
      set_attr(child, "typeface", typeface);
      return;
    }
  }
  pugi::xml_node added = font_node.append_child((prefix_of(font_node.name()) + "font").c_str());
  set_attr(added, "script", script);
  set_attr(added, "typeface", typeface);
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
  set_script_face(font_node, "Jpan", east_asian);
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
    return Expected<void, Error>(std::move(path_or.error()));
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
    return Expected<void, Error>(std::move(path_or.error()));
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

Expected<void, Error> set_theme_script_font(Workbook& wb, bool major, std::string_view script,
                                            std::string_view typeface) {
  bool valid = script.size() == 4U;
  for (const char c : script) {
    valid = valid && ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'));
  }
  if (!valid || typeface.empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument,
                      valid ? "theme: typeface must not be empty" : "theme: script must be four ASCII letters",
                      "script=" + std::string(script));
  }
  pugi::xml_document doc;
  auto path_or = open_theme_for_edit(wb, doc);
  if (!path_or) {
    return Expected<void, Error>(std::move(path_or.error()));
  }
  const pugi::xml_node font_node = find_child(font_scheme_of(doc), major ? "majorFont" : "minorFont");
  if (!font_node) {
    return make_error(FormulonErrorCode::kIoXmlParse, "theme: font scheme has no major/minor font",
                      "part=" + path_or.value());
  }
  set_script_face(font_node, script, typeface);
  return ooxml::store_part_dom(wb, path_or.value(), doc);
}

Expected<void, Error> reset_theme(Workbook& wb) {
  // The rel is dropped after the path is read: without it the path falls back to the default.
  const std::string path = theme_part_path(wb);
  if (ooxml::find_passthrough_part(wb, path) == nullptr) {
    return Expected<void, Error>();
  }
  std::vector<UnknownRelationship> rels = wb.unknown_workbook_rels();
  rels.erase(std::remove_if(rels.begin(), rels.end(),
                            [&path](const UnknownRelationship& rel) {
                              return !rel.target_external && rel.type == kRelTheme &&
                                     (rel.target == path || "xl/" + rel.target == path);
                            }),
             rels.end());
  wb.set_unknown_workbook_rels(std::move(rels));
  ooxml::remove_orphans(wb, {path});
  return Expected<void, Error>();
}

}  // namespace formulon::io
