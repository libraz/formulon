#include "drawing/drawing_color.h"

#include <algorithm>
#include <cctype>
#include <climits>
#include <cmath>
#include <cstdint>
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
  const char* text = val.value();
  out = std::strtol(text, &end, 10);
  return end != text && *end == '\0';
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

// --- Rendering. Excel holds each channel in linear light on the 1/100000 `scrgbClr` grid between
// transforms. Hue, saturation and luminance transforms run on float HSL of the sRGB value (a run of
// them in one pass); shade, tint, inversion and channel transforms run on the linear value.

double to_linear(double v) {
  return v <= 0.04045 ? v / 12.92 : std::pow((v + 0.055) / 1.055, 2.4);
}

double from_linear(double v) {
  return v <= 0.0031308 ? v * 12.92 : 1.055 * std::pow(v, 1.0 / 2.4) - 0.055;
}

double on_grid(double v) {
  return std::floor(v * static_cast<double>(kUnit) + 0.5) / static_cast<double>(kUnit);
}

struct Hls {
  double h = 0.0;
  double l = 0.0;
  double s = 0.0;
};

Hls to_hls(double r, double g, double b) {
  const double hi = std::max({r, g, b});
  const double lo = std::min({r, g, b});
  Hls out;
  out.l = (lo + hi) / 2.0;
  if (hi == lo) {
    return out;
  }
  const double span = hi - lo;
  out.s = out.l <= 0.5 ? span / (hi + lo) : span / (2.0 - hi - lo);
  const double rc = (hi - r) / span;
  const double gc = (hi - g) / span;
  const double bc = (hi - b) / span;
  double h = r == hi ? bc - gc : g == hi ? 2.0 + rc - bc : 4.0 + gc - rc;
  h /= 6.0;
  out.h = h - std::floor(h);
  return out;
}

double hue_channel(double m1, double m2, double hue) {
  hue -= std::floor(hue);
  if (hue < 1.0 / 6.0) {
    return m1 + (m2 - m1) * hue * 6.0;
  }
  if (hue < 0.5) {
    return m2;
  }
  if (hue < 2.0 / 3.0) {
    return m1 + (m2 - m1) * (2.0 / 3.0 - hue) * 6.0;
  }
  return m1;
}

void from_hls(const Hls& c, double rgb[3]) {
  if (c.s == 0.0) {
    rgb[0] = rgb[1] = rgb[2] = c.l;
    return;
  }
  const double m2 = c.l <= 0.5 ? c.l * (1.0 + c.s) : c.l + c.s - c.l * c.s;
  const double m1 = 2.0 * c.l - m2;
  rgb[0] = hue_channel(m1, m2, c.h + 1.0 / 3.0);
  rgb[1] = hue_channel(m1, m2, c.h);
  rgb[2] = hue_channel(m1, m2, c.h - 1.0 / 3.0);
}

bool is_hls_transform(std::string_view name) {
  return name == "lum" || name == "lumMod" || name == "lumOff" || name == "sat" || name == "satMod" ||
         name == "satOff" || name == "hue" || name == "hueMod" || name == "hueOff" || name == "comp";
}

// Applies one HSL transform; false when it needs a value it lacks. Luminance clamps to 0..1; saturation does
// not, and hue wraps.
bool apply_hls(const ColorTransform& t, Hls& c) {
  if (t.name == "comp") {
    c.h += 0.5;
    c.h -= std::floor(c.h);
    return true;
  }
  if (!t.value) {
    return false;
  }
  const double v = static_cast<double>(*t.value) / static_cast<double>(kUnit);
  constexpr double kAngleUnit = 21600000.0;  // 60000ths of a degree in a turn.
  const std::string_view name = t.name;
  if (name == "lum") {
    c.l = std::clamp(v, 0.0, 1.0);
  } else if (name == "lumMod") {
    c.l = std::clamp(c.l * v, 0.0, 1.0);
  } else if (name == "lumOff") {
    c.l = std::clamp(c.l + v, 0.0, 1.0);
  } else if (name == "sat") {
    c.s = v;
  } else if (name == "satMod") {
    c.s *= v;
  } else if (name == "satOff") {
    c.s += v;
  } else {
    c.h = name == "hue"      ? static_cast<double>(*t.value) / kAngleUnit
          : name == "hueMod" ? c.h * v
                             : c.h + static_cast<double>(*t.value) / kAngleUnit;
    c.h -= std::floor(c.h);
  }
  return true;
}

// The channel (0 red, 1 green, 2 blue) a `red`/`redMod`/`redOff`-style transform names; -1 otherwise.
int channel_of(std::string_view name) {
  constexpr std::string_view kChannels[] = {"red", "green", "blue"};
  for (int i = 0; i < 3; ++i) {
    const std::string_view ch = kChannels[i];
    if (strings::starts_with(name, ch)) {
      const std::string_view rest = name.substr(ch.size());
      return rest.empty() || rest == "Mod" || rest == "Off" ? i : -1;
    }
  }
  return -1;
}

bool valid_transform_name(std::string_view name) {
  if (name.empty() || name.size() > 32U || !std::isalpha(static_cast<unsigned char>(name[0]))) {
    return false;
  }
  return std::all_of(name.begin(), name.end(),
                     [](char ch) { return std::isalnum(static_cast<unsigned char>(ch)) != 0; });
}

/// The namespace prefix of an element's qualified name; empty when it has none.
std::string_view prefix_of(const xml_node& node) {
  const std::string_view qname = node.name();
  const std::size_t colon = qname.find(':');
  return colon == std::string_view::npos ? std::string_view() : qname.substr(0, colon);
}

std::string_view scheme_name_of(std::uint32_t theme) {
  return kSchemeNames[std::min<std::uint32_t>(theme, 11U)];
}

}  // namespace

ResolvedPaint resolve_transforms(std::uint32_t base_argb, const std::vector<ColorTransform>& all) {
  ResolvedPaint out;
  const double unit = static_cast<double>(kUnit);
  // Opacity is a pass of its own over the alpha transforms, in their written order; the colour never sees them.
  double opacity = 1.0;
  std::vector<ColorTransform> transforms;
  for (const ColorTransform& t : all) {
    const std::string_view name = t.name;
    if (name != "alpha" && name != "alphaMod" && name != "alphaOff") {
      transforms.push_back(t);
      continue;
    }
    if (!t.value) {
      return out;
    }
    const double v = static_cast<double>(*t.value) / unit;
    opacity = std::clamp(name == "alpha" ? v : name == "alphaMod" ? opacity * v : opacity + v, 0.0, 1.0);
  }
  double lin[3];
  for (int i = 0; i < 3; ++i) {
    const auto byte = static_cast<double>((base_argb >> (16 - 8 * i)) & 0xFFU);
    lin[i] = on_grid(to_linear(byte / 255.0));
  }
  for (std::size_t i = 0; i < transforms.size();) {
    const ColorTransform& t = transforms[i];
    if (is_hls_transform(t.name)) {
      // A run of hue, saturation and luminance transforms shares one HSL pass.
      Hls c = to_hls(from_linear(lin[0]), from_linear(lin[1]), from_linear(lin[2]));
      for (; i < transforms.size() && is_hls_transform(transforms[i].name); ++i) {
        if (!apply_hls(transforms[i], c)) {
          return out;
        }
      }
      double rgb[3];
      from_hls(c, rgb);
      for (int k = 0; k < 3; ++k) {
        lin[k] = on_grid(to_linear(rgb[k]));
      }
      continue;
    }
    const std::string_view name = t.name;
    const double v = t.value ? static_cast<double>(*t.value) / unit : 0.0;
    const bool needs_value = name != "inv" && name != "gray" && name != "gamma" && name != "invGamma";
    if (needs_value && !t.value) {
      return out;
    }
    if (name == "shade") {
      for (double& x : lin) {
        x = on_grid(x * v);
      }
    } else if (name == "tint") {
      for (double& x : lin) {
        x = on_grid(1.0 - (1.0 - x) * v);
      }
    } else if (name == "inv") {
      for (double& x : lin) {
        x = on_grid(1.0 - x);
      }
    } else if (name == "gray") {
      const double g = 0.2126 * from_linear(lin[0]) + 0.7152 * from_linear(lin[1]) + 0.0722 * from_linear(lin[2]);
      lin[0] = lin[1] = lin[2] = on_grid(to_linear(g));
    } else if (name == "gamma") {
      // Off the grid: the next transform or the output quantises it.
      for (double& x : lin) {
        x = from_linear(x);
      }
    } else if (name == "invGamma") {
      for (double& x : lin) {
        x = on_grid(to_linear(x));
      }
    } else if (const int ch = channel_of(name); ch >= 0) {
      const std::string_view op = name.substr(name[0] == 'r' ? 3 : name[0] == 'g' ? 5 : 4);
      lin[ch] = on_grid(op.empty() ? v : op == "Mod" ? lin[ch] * v : lin[ch] + v);
    } else {
      return out;
    }
    ++i;
  }
  std::uint32_t argb = 0xFF000000U;
  for (int i = 0; i < 3; ++i) {
    const double v = std::clamp(from_linear(std::clamp(lin[i], 0.0, 1.0)) * 255.0, 0.0, 255.0);
    argb |= static_cast<std::uint32_t>(std::floor(v + 0.5)) << (16 - 8 * i);
  }
  out.argb = argb;
  out.alpha = static_cast<std::uint32_t>(std::floor(opacity * unit + 0.5));
  return out;
}

DrawingColor expand_tint(const DrawingColor& color) {
  if (color.base.kind != ColorSpec::Kind::kTheme || color.base.tint == 0.0) {
    return color;
  }
  DrawingColor out = color;
  const Luminance lum = luminance_of(color.base.tint);
  out.base.tint = 0.0;
  std::vector<ColorTransform> pair;
  if (lum.mod != kUnit) {
    pair.push_back({"lumMod", static_cast<std::int32_t>(lum.mod)});
  }
  if (lum.off != 0) {
    pair.push_back({"lumOff", static_cast<std::int32_t>(lum.off)});
  }
  out.transforms.insert(out.transforms.begin(), pair.begin(), pair.end());
  return out;
}

Expected<void, Error> validate_input(const DrawingColor& color) {
  const ColorSpec::Kind kind = color.base.kind;
  if (kind != ColorSpec::Kind::kRgb && kind != ColorSpec::Kind::kTheme) {
    return make_error(FormulonErrorCode::kInvalidArgument, "A drawing colour takes an Rgb or Theme base");
  }
  if (kind == ColorSpec::Kind::kTheme && color.base.theme > 11U) {
    return make_error(FormulonErrorCode::kInvalidArgument, "A theme colour index is 0..11");
  }
  if (kind == ColorSpec::Kind::kTheme && color.base.tint != 0.0 && !color.transforms.empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "A tint and transforms cannot be given together");
  }
  if (color.transforms.size() > kMaxTransforms) {
    return make_error(FormulonErrorCode::kSecurityResourceLimit, "A drawing colour takes at most 64 transforms");
  }
  for (const ColorTransform& t : color.transforms) {
    if (!valid_transform_name(t.name)) {
      return make_error(FormulonErrorCode::kInvalidArgument, "A colour transform name is 1..32 letters and digits");
    }
  }
  return Expected<void, Error>::Ok();
}

bool same_drawing_color(const DrawingColor& a, const DrawingColor& b) {
  const DrawingColor ea = expand_tint(a);
  const DrawingColor eb = expand_tint(b);
  return same_color(ea.base, eb.base) && ea.transforms == eb.transforms;
}

void write_drawing_color(xml_node parent, const std::string& prefix, const DrawingColor& color,
                         const std::string& alpha) {
  const DrawingColor c = expand_tint(color);
  xml_node clr;
  if (c.base.kind == ColorSpec::Kind::kTheme) {
    clr = parent.append_child((prefix + "schemeClr").c_str());
    clr.append_attribute("val").set_value(std::string(scheme_name_of(c.base.theme)).c_str());
  } else {
    char hex[8];
    format_hex(hex, sizeof hex, c.base.rgb & 0xFFFFFFU, 6, true);
    clr = parent.append_child((prefix + "srgbClr").c_str());
    clr.append_attribute("val").set_value(hex);
  }
  for (const ColorTransform& t : c.transforms) {
    xml_node node = clr.append_child((prefix + t.name).c_str());
    if (t.value) {
      node.append_attribute("val").set_value(std::to_string(*t.value).c_str());
    }
  }
  if (!alpha.empty()) {
    clr.append_child((prefix + "alpha").c_str()).append_attribute("val").set_value(alpha.c_str());
  }
}

namespace {

/// `read_drawing_color`, rendering only when `wb` is given.
std::optional<DrawingColor> parse_drawing_color(const Workbook* wb, const xml_node& clr, const xml_node& clr_map_ovr) {
  const std::string_view kind = io::local_name(clr);
  const std::string_view val = clr.attribute("val").value();
  DrawingColor color;
  std::uint32_t base_argb = 0;
  if (kind == "srgbClr") {
    if (val.size() != 6U || val.find_first_not_of("0123456789abcdefABCDEF") != std::string_view::npos) {
      return std::nullopt;
    }
    base_argb = io::parse_rgb_hex(val, 0U) | 0xFF000000U;
    color.base = rgb_spec(base_argb);
  } else if (kind == "schemeClr") {
    auto index = scheme_index(val);
    if (!index) {
      return std::nullopt;
    }
    bool mapped = false;
    if (*index <= 3U && (val == "bg1" || val == "tx1" || val == "bg2" || val == "tx2")) {
      const std::string_view to = clr_map_ovr.attribute(std::string(val).c_str()).value();
      if (!to.empty()) {
        const auto target = scheme_index(to);
        if (!target || *target > 3U) {
          return std::nullopt;
        }
        mapped = *target != *index;
        index = target;
      }
    }
    ColorSpec theme;
    theme.kind = ColorSpec::Kind::kTheme;
    theme.theme = *index;
    base_argb = wb ? resolve_color(*wb, theme, ColorContext::kFont).argb | 0xFF000000U : 0U;
    color.base = mapped ? rgb_spec(base_argb) : theme;
  } else if (kind == "sysClr") {
    const std::string_view last = clr.attribute("lastClr").value();
    base_argb =
        !last.empty() ? io::parse_rgb_hex(last, 0U) | 0xFF000000U : (val == "window" ? 0xFFFFFFFFU : 0xFF000000U);
    color.base = rgb_spec(base_argb);
  } else {
    return std::nullopt;
  }
  const std::string_view own_prefix = prefix_of(clr);
  for (xml_node child = clr.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string_view name = io::local_name(child);
    if (prefix_of(child) != own_prefix) {
      return std::nullopt;
    }
    ColorTransform t;
    t.name = std::string(name);
    if (const pugi::xml_attribute v = child.attribute("val")) {
      long parsed = 0;
      if (!val_of(child, parsed) || parsed < INT32_MIN || parsed > INT32_MAX) {
        return std::nullopt;
      }
      t.value = static_cast<std::int32_t>(parsed);
    }
    color.transforms.push_back(std::move(t));
    if (color.transforms.size() > kMaxTransforms) {
      return std::nullopt;
    }
  }
  if (wb) {
    const ResolvedPaint resolved = resolve_transforms(base_argb, color.transforms);
    color.resolved_argb = resolved.argb;
    color.resolved_alpha = resolved.alpha;
  }
  return color;
}

}  // namespace

std::optional<DrawingColor> read_drawing_color(const Workbook& wb, const xml_node& clr, const xml_node& clr_map_ovr) {
  return parse_drawing_color(&wb, clr, clr_map_ovr);
}

std::optional<DrawingColor> read_written_color(const xml_node& clr) {
  return parse_drawing_color(nullptr, clr, xml_node());
}

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
