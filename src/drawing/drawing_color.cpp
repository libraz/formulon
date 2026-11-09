#include "drawing/drawing_color.h"

#include <algorithm>
#include <cmath>
#include <cstdlib>

#include "color_resolve.h"
#include "io/xml_utils.h"
#include "utils/number_text.h"
#include "utils/strings.h"

namespace formulon {
namespace drawing_color {
namespace {

using pugi::xml_node;

constexpr long kUnit = 100000;
constexpr std::string_view kSchemeNames[] = {"bg1",     "tx1",     "bg2",     "tx2",     "accent1", "accent2",
                                             "accent3", "accent4", "accent5", "accent6", "hlink",   "folHlink"};

/// `a:lumMod` / `a:lumOff` of a tint, in 1/100000 units; {kUnit, 0} is the identity.
struct Luminance {
  long mod = kUnit;
  long off = 0;
};

Luminance luminance_of(double tint) {
  Luminance lum;
  if (tint > 0.0) {
    lum.off = std::lround(tint * static_cast<double>(kUnit));
    lum.mod = kUnit - lum.off;
  } else if (tint < 0.0) {
    lum.mod = std::lround((1.0 + tint) * static_cast<double>(kUnit));
  }
  return lum;
}

bool val_of(const xml_node& node, long& out) {
  const pugi::xml_attribute val = node.attribute("val");
  if (!val) {
    return false;
  }
  char* end = nullptr;
  const std::string text = val.value();
  out = std::strtol(text.c_str(), &end, 10);
  return end != text.c_str() && *end == '\0';
}

/// Tint of the `lumMod` / `lumOff` children, when they are a form a tint writes.
std::optional<double> tint_of(const xml_node& clr) {
  bool has_mod = false;
  bool has_off = false;
  long mod = 0;
  long off = 0;
  for (xml_node child = clr.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string_view name = io::local_name(child);
    if (name == "lumMod" && !has_mod) {
      has_mod = val_of(child, mod);
      if (!has_mod) {
        return std::nullopt;
      }
    } else if (name == "lumOff" && !has_off) {
      has_off = val_of(child, off);
      if (!has_off) {
        return std::nullopt;
      }
    } else if (name != "alpha") {
      return std::nullopt;
    }
  }
  if (has_off) {
    if (!has_mod || off <= 0 || mod + off != kUnit) {
      return std::nullopt;
    }
    return static_cast<double>(off) / static_cast<double>(kUnit);
  }
  if (has_mod) {
    if (mod < 0 || mod > kUnit) {
      return std::nullopt;
    }
    return static_cast<double>(mod) / static_cast<double>(kUnit) - 1.0;
  }
  return 0.0;
}

/// Applies `lumMod` / `lumOff` as the equivalent spreadsheet tint (the approximation for any form).
std::uint32_t apply_luminance(std::uint32_t argb, const xml_node& clr) {
  const xml_node mod = io::child_local(clr, "lumMod");
  const xml_node off = io::child_local(clr, "lumOff");
  if (!mod && !off) {
    return argb;
  }
  const double m = mod ? mod.attribute("val").as_double() / 100000.0 : 1.0;
  const double o = off ? off.attribute("val").as_double() / 100000.0 : 0.0;
  const double tint = o > 0.0 ? o : m - 1.0;
  return apply_tint(argb, std::clamp(tint, -1.0, 1.0));
}

ColorSpec rgb_spec(std::uint32_t argb) {
  ColorSpec spec;
  spec.kind = ColorSpec::Kind::kRgb;
  spec.rgb = argb;
  return spec;
}

}  // namespace

std::optional<std::uint32_t> scheme_index(std::string_view name) {
  if (name == "bg1" || name == "lt1")
    return 0U;
  if (name == "tx1" || name == "dk1")
    return 1U;
  if (name == "bg2" || name == "lt2")
    return 2U;
  if (name == "tx2" || name == "dk2")
    return 3U;
  if (name == "hlink")
    return 10U;
  if (name == "folHlink")
    return 11U;
  if (name.size() == 7 && strings::starts_with(name, "accent") && name[6] >= '1' && name[6] <= '6') {
    return 4U + static_cast<std::uint32_t>(name[6] - '1');
  }
  return std::nullopt;
}

bool remaps_color_map(const xml_node& clr_map_ovr) {
  const struct {
    const char* attr;
    const char* identity;
  } kIdentity[] = {{"bg1", "lt1"}, {"tx1", "dk1"}, {"bg2", "lt2"}, {"tx2", "dk2"}};
  for (const auto& entry : kIdentity) {
    const pugi::xml_attribute attr = clr_map_ovr.attribute(entry.attr);
    if (attr && std::string_view(attr.value()) != entry.identity) {
      return true;
    }
  }
  return false;
}

std::optional<ColorSpec> read_theme(const xml_node& clr, bool remapped) {
  if (io::local_name(clr) != "schemeClr") {
    return std::nullopt;
  }
  const std::string_view name = clr.attribute("val").value();
  const auto index = scheme_index(name);
  if (!index) {
    return std::nullopt;
  }
  if (remapped && (name == "bg1" || name == "tx1" || name == "bg2" || name == "tx2")) {
    return std::nullopt;
  }
  const auto tint = tint_of(clr);
  if (!tint) {
    return std::nullopt;
  }
  ColorSpec spec;
  spec.kind = ColorSpec::Kind::kTheme;
  spec.theme = *index;
  spec.tint = *tint;
  return spec;
}

std::optional<ColorSpec> read_rendered(const Workbook& wb, const xml_node& clr) {
  const std::string_view kind = io::local_name(clr);
  const std::string_view val = clr.attribute("val").value();
  if (kind == "srgbClr") {
    return rgb_spec(apply_luminance(io::parse_rgb_hex(val, 0U) | 0xFF000000U, clr));
  }
  if (kind == "schemeClr") {
    if (const auto index = scheme_index(val)) {
      ColorSpec spec;
      spec.kind = ColorSpec::Kind::kTheme;
      spec.theme = *index;
      return rgb_spec(apply_luminance(resolve_color(wb, spec, ColorContext::kFont).argb, clr));
    }
    return std::nullopt;
  }
  if (kind == "sysClr") {
    const std::string_view last = clr.attribute("lastClr").value();
    if (!last.empty()) {
      return rgb_spec(apply_luminance(io::parse_rgb_hex(last, 0U) | 0xFF000000U, clr));
    }
    return rgb_spec(val == "window" ? 0xFFFFFFFFU : 0xFF000000U);
  }
  return std::nullopt;
}

std::optional<ColorSpec> read_color(const Workbook& wb, const xml_node& clr, bool remapped) {
  if (auto theme = read_theme(clr, remapped)) {
    return theme;
  }
  return read_rendered(wb, clr);
}

std::optional<ColorSpec> read_exact(const xml_node& clr, bool remapped) {
  if (io::local_name(clr) != "srgbClr") {
    return read_theme(clr, remapped);
  }
  const std::string_view val = clr.attribute("val").value();
  if (val.size() != 6U || val.find_first_not_of("0123456789abcdefABCDEF") != std::string_view::npos) {
    return std::nullopt;
  }
  for (xml_node child = clr.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element && io::local_name(child) != "alpha") {
      return std::nullopt;
    }
  }
  return rgb_spec(io::parse_rgb_hex(val, 0U) | 0xFF000000U);
}

std::string alpha_of(const xml_node& clr) {
  return io::child_local(clr, "alpha").attribute("val").value();
}

bool same_color(const ColorSpec& a, const ColorSpec& b) {
  if (a.kind != b.kind) {
    return false;
  }
  if (a.kind == ColorSpec::Kind::kRgb) {
    return (a.rgb & 0xFFFFFFU) == (b.rgb & 0xFFFFFFU);
  }
  if (a.kind == ColorSpec::Kind::kTheme) {
    const Luminance la = luminance_of(a.tint);
    const Luminance lb = luminance_of(b.tint);
    return a.theme == b.theme && la.mod == lb.mod && la.off == lb.off;
  }
  return true;
}

void write_color(xml_node parent, const std::string& prefix, const ColorSpec& color, const std::string& alpha) {
  xml_node clr;
  if (color.kind == ColorSpec::Kind::kTheme) {
    clr = parent.append_child((prefix + "schemeClr").c_str());
    clr.append_attribute("val").set_value(std::string(kSchemeNames[std::min<std::uint32_t>(color.theme, 11U)]).c_str());
    const Luminance lum = luminance_of(color.tint);
    if (lum.mod != kUnit) {
      clr.append_child((prefix + "lumMod").c_str()).append_attribute("val").set_value(std::to_string(lum.mod).c_str());
    }
    if (lum.off != 0) {
      clr.append_child((prefix + "lumOff").c_str()).append_attribute("val").set_value(std::to_string(lum.off).c_str());
    }
  } else {
    char hex[8];
    format_hex(hex, sizeof hex, color.rgb & 0xFFFFFFU, 6, true);
    clr = parent.append_child((prefix + "srgbClr").c_str());
    clr.append_attribute("val").set_value(hex);
  }
  if (!alpha.empty()) {
    clr.append_child((prefix + "alpha").c_str()).append_attribute("val").set_value(alpha.c_str());
  }
}

}  // namespace drawing_color
}  // namespace formulon
