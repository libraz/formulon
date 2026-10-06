#include "io/color_spec_xml.h"

#include "io/xml_utils.h"
#include "utils/number_text.h"

namespace formulon {
namespace io {

ColorSpec read_color_spec(const pugi::xml_node& node) {
  ColorSpec spec;
  if (!node) {
    return spec;
  }
  if (pugi::xml_attribute rgb = node.attribute("rgb"); rgb) {
    spec.kind = ColorSpec::Kind::kRgb;
    spec.rgb = parse_rgb_hex(rgb.value(), 0xFF000000U);
    return spec;
  }
  if (pugi::xml_attribute theme = node.attribute("theme"); theme) {
    spec.kind = ColorSpec::Kind::kTheme;
    spec.theme = theme.as_uint(0U);
    // Signed by design (ECMA-376 bounds it to [-1.0, 1.0]); attr_f64 rejects NaN and infinity.
    spec.tint = attr_f64(node, "tint", 0.0);
    return spec;
  }
  if (pugi::xml_attribute indexed = node.attribute("indexed"); indexed) {
    spec.kind = ColorSpec::Kind::kIndexed;
    spec.indexed = indexed.as_uint(0U);
    return spec;
  }
  if (node.attribute("auto")) {
    spec.kind = ColorSpec::Kind::kAuto;
  }
  return spec;
}

void append_color_spec_attrs(std::string& out, const ColorSpec& spec, std::uint32_t fallback_argb) {
  char buf[24];
  switch (spec.kind) {
    case ColorSpec::Kind::kTheme:
      format_unsigned(buf, sizeof(buf), spec.theme);
      out.append(" theme=\"");
      out.append(buf);
      out.push_back('"');
      if (spec.tint != 0.0) {
        out.append(" tint=\"");
        append_xml_number(out, spec.tint);
        out.push_back('"');
      }
      return;
    case ColorSpec::Kind::kIndexed:
      format_unsigned(buf, sizeof(buf), spec.indexed);
      out.append(" indexed=\"");
      out.append(buf);
      out.push_back('"');
      return;
    case ColorSpec::Kind::kAuto:
      out.append(" auto=\"1\"");
      return;
    case ColorSpec::Kind::kRgb:
    case ColorSpec::Kind::kNone:
      break;
  }
  format_hex(buf, sizeof(buf), spec.kind == ColorSpec::Kind::kRgb ? spec.rgb : fallback_argb, 8, true);
  out.append(" rgb=\"");
  out.append(buf);
  out.push_back('"');
}

}  // namespace io
}  // namespace formulon
