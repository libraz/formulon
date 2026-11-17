#include "drawing/drawing_color.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "pugixml.hpp"
#include "workbook.h"

namespace formulon {
namespace {

pugi::xml_node Parse(pugi::xml_document& doc, const std::string& xml) {
  EXPECT_TRUE(doc.load_string(xml.c_str()));
  return doc.first_child();
}

std::optional<ColorSpec> ReadTheme(const std::string& clr, bool remapped = false) {
  pugi::xml_document doc;
  return drawing_color::read_theme(Parse(doc, "<a:schemeClr xmlns:a=\"urn:a\" " + clr.substr(0, clr.find('>')) + ">" +
                                                  clr.substr(clr.find('>') + 1)),
                                   remapped);
}

ColorSpec Theme(std::uint32_t theme, double tint) {
  ColorSpec spec;
  spec.kind = ColorSpec::Kind::kTheme;
  spec.theme = theme;
  spec.tint = tint;
  return spec;
}

std::string Written(const ColorSpec& color, const std::string& alpha = "") {
  pugi::xml_document doc;
  pugi::xml_node fill = doc.append_child("a:solidFill");
  drawing_color::write_color(fill, "a:", color, alpha);
  std::string out;
  struct Writer : pugi::xml_writer {
    std::string* sink;
    void write(const void* data, std::size_t size) override { sink->append(static_cast<const char*>(data), size); }
  } writer;
  writer.sink = &out;
  fill.first_child().print(writer, "", pugi::format_raw);
  return out;
}

TEST(DrawingColor, SchemeNumbering) {
  EXPECT_EQ(drawing_color::scheme_index("bg1"), 0U);
  EXPECT_EQ(drawing_color::scheme_index("tx1"), 1U);
  EXPECT_EQ(drawing_color::scheme_index("lt2"), 2U);
  EXPECT_EQ(drawing_color::scheme_index("tx2"), 3U);
  EXPECT_EQ(drawing_color::scheme_index("accent6"), 9U);
  EXPECT_EQ(drawing_color::scheme_index("hlink"), 10U);
  EXPECT_EQ(drawing_color::scheme_index("folHlink"), 11U);
  EXPECT_FALSE(drawing_color::scheme_index("phClr").has_value());
}

TEST(DrawingColor, WritesPaletteForms) {
  EXPECT_EQ(Written(Theme(4, 0.4)),
            "<a:schemeClr val=\"accent1\"><a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/></a:schemeClr>");
  EXPECT_EQ(Written(Theme(1, -0.25)), "<a:schemeClr val=\"tx1\"><a:lumMod val=\"75000\"/></a:schemeClr>");
  EXPECT_EQ(Written(Theme(0, 0.0)), "<a:schemeClr val=\"bg1\"/>");
  EXPECT_EQ(Written(Theme(2, 0.0)), "<a:schemeClr val=\"bg2\"/>");
  EXPECT_EQ(Written(Theme(3, 0.0)), "<a:schemeClr val=\"tx2\"/>");
  EXPECT_EQ(Written(Theme(11, 0.0)), "<a:schemeClr val=\"folHlink\"/>");
  EXPECT_EQ(Written(Theme(5, 1e-7)), "<a:schemeClr val=\"accent2\"/>");
  EXPECT_EQ(Written(Theme(5, 0.4), "30000"),
            "<a:schemeClr val=\"accent2\"><a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/><a:alpha val=\"30000\"/>"
            "</a:schemeClr>");
  ColorSpec rgb;
  rgb.kind = ColorSpec::Kind::kRgb;
  rgb.rgb = 0x80123456U;
  EXPECT_EQ(Written(rgb, "1000"), "<a:srgbClr val=\"123456\"><a:alpha val=\"1000\"/></a:srgbClr>");
}

TEST(DrawingColor, ReadsOnlyTheTintForms) {
  const auto plain = ReadTheme("val=\"accent3\"/>");
  ASSERT_TRUE(plain.has_value());
  EXPECT_EQ(plain->theme, 6U);
  EXPECT_EQ(plain->tint, 0.0);
  EXPECT_FALSE(ReadTheme("val=\"accent3\"><a:lumOff val=\"40000\"/></a:schemeClr>").has_value());
  EXPECT_FALSE(
      ReadTheme("val=\"accent3\"><a:lumMod val=\"50000\"/><a:lumOff val=\"40000\"/></a:schemeClr>").has_value());
  EXPECT_FALSE(ReadTheme("val=\"accent3\"><a:lumMod val=\"100000\"/><a:lumOff val=\"0\"/></a:schemeClr>").has_value());
  EXPECT_FALSE(ReadTheme("val=\"accent3\"><a:lumMod val=\"120000\"/></a:schemeClr>").has_value());
  EXPECT_FALSE(ReadTheme("val=\"accent3\"><a:shade val=\"50000\"/></a:schemeClr>").has_value());
  EXPECT_FALSE(ReadTheme("val=\"accent3\"><a:lumMod val=\"50000\"/><a:satMod val=\"1\"/></a:schemeClr>").has_value());
  EXPECT_FALSE(ReadTheme("val=\"phClr\"/>").has_value());
  const auto alpha = ReadTheme("val=\"tx2\"><a:lumMod val=\"75000\"/><a:alpha val=\"20000\"/></a:schemeClr>");
  ASSERT_TRUE(alpha.has_value());
  EXPECT_EQ(alpha->theme, 3U);
  EXPECT_DOUBLE_EQ(alpha->tint, -0.25);
}

TEST(DrawingColor, RemappedMapWithholdsBaseColours) {
  EXPECT_FALSE(ReadTheme("val=\"bg1\"/>", true).has_value());
  EXPECT_FALSE(ReadTheme("val=\"tx2\"/>", true).has_value());
  EXPECT_TRUE(ReadTheme("val=\"lt1\"/>", true).has_value());
  EXPECT_TRUE(ReadTheme("val=\"accent1\"/>", true).has_value());

  pugi::xml_document identity;
  EXPECT_FALSE(drawing_color::remaps_color_map(
      Parse(identity, "<c:clrMapOvr bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\"/>")));
  pugi::xml_document swapped;
  EXPECT_TRUE(drawing_color::remaps_color_map(Parse(swapped, "<c:clrMapOvr bg1=\"dk1\" tx1=\"lt1\"/>")));
  pugi::xml_document none;
  EXPECT_FALSE(drawing_color::remaps_color_map(none.first_child()));
}

TEST(DrawingColor, TintsCompareAtStoredPrecision) {
  EXPECT_TRUE(drawing_color::same_color(Theme(4, 0.123456), Theme(4, 0.12346)));
  EXPECT_FALSE(drawing_color::same_color(Theme(4, 0.4), Theme(4, 0.41)));
  EXPECT_FALSE(drawing_color::same_color(Theme(4, 0.4), Theme(5, 0.4)));
  ColorSpec rgb;
  rgb.kind = ColorSpec::Kind::kRgb;
  EXPECT_FALSE(drawing_color::same_color(rgb, Theme(4, 0.0)));
}

std::optional<DrawingColor> ReadDrawing(const Workbook& wb, const std::string& clr, const std::string& map = "") {
  pugi::xml_document doc;
  pugi::xml_document map_doc;
  const pugi::xml_node map_node = map.empty() ? pugi::xml_node() : Parse(map_doc, map);
  return drawing_color::read_drawing_color(wb, Parse(doc, clr), map_node);
}

std::uint32_t Resolved(std::uint32_t rgb, const std::vector<ColorTransform>& transforms) {
  return drawing_color::resolve_transforms(0xFF000000U | rgb, transforms).argb.value_or(0U);
}

std::string WrittenDrawing(const DrawingColor& color, const std::string& alpha = "") {
  pugi::xml_document doc;
  pugi::xml_node fill = doc.append_child("a:solidFill");
  drawing_color::write_drawing_color(fill, "a:", color, alpha);
  std::string out;
  struct Writer : pugi::xml_writer {
    std::string* sink;
    void write(const void* data, std::size_t size) override { sink->append(static_cast<const char*>(data), size); }
  } writer;
  writer.sink = &out;
  fill.first_child().print(writer, "", pugi::format_raw);
  return out;
}

DrawingColor ThemeDrawing(std::uint32_t theme, std::vector<ColorTransform> transforms, double tint = 0.0) {
  DrawingColor color;
  color.base = Theme(theme, tint);
  color.transforms = std::move(transforms);
  return color;
}

// Excel 16.115 renders, for the Office theme colours named.
TEST(DrawingColorResolve, LuminanceMatchesExcel) {
  EXPECT_EQ(Resolved(0xE97132, {{"lumMod", 60000}, {"lumOff", 40000}}), 0xFFF2AA84U);  // accent2 -> 242,170,132
  EXPECT_EQ(Resolved(0x156082, {{"lumMod", 75000}}), 0xFF104862U);                     // accent1 -> 16,72,98
  EXPECT_EQ(Resolved(0xE97132, {{"lumMod", 20000}}), 0xFF331506U);                     // 51,21,6
  EXPECT_EQ(Resolved(0xFFFFFF, {{"lumMod", 50000}}), 0xFF7F7F7FU);                     // 127, not 128
  EXPECT_EQ(Resolved(0x000000, {{"lumMod", 50000}, {"lumOff", 50000}}), 0xFF7F7F7FU);
  EXPECT_EQ(Resolved(0x0E2841, {{"lumMod", 25000}}), 0xFF030A10U);                     // 3,10,16
  EXPECT_EQ(Resolved(0x156082, {{"lumMod", 95000}, {"lumOff", 5000}}), 0xFF186B91U);   // 24,107,145
  EXPECT_EQ(Resolved(0xE8E8E8, {{"lumMod", 50000}, {"lumOff", 50000}}), 0xFFF4F4F4U);  // 244
  EXPECT_EQ(Resolved(0x4EA72E, {{"lumMod", 25000}}), 0xFF132A0CU);                     // 19,42,12
  EXPECT_EQ(Resolved(0xA02B93, {{"lumMod", 1000}}), 0xFF020001U);                      // 2,0,1
  EXPECT_EQ(Resolved(0x123456, {}), 0xFF123456U);
}

TEST(DrawingColorResolve, LinearLightTransformsMatchExcel) {
  EXPECT_EQ(Resolved(0xFFFFFF, {{"shade", 50000}}), 0xFFBCBCBCU);  // lt1 shade 50000 -> 188
  EXPECT_EQ(Resolved(0xE97132, {{"tint", 50000}}), 0xFFF4C9BEU);
  EXPECT_EQ(Resolved(0xE97132, {{"inv", std::nullopt}}), 0xFF77ECFBU);
  EXPECT_EQ(Resolved(0xE97132, {{"redMod", 50000}}), 0xFFAB7132U);
  EXPECT_EQ(Resolved(0xE97132, {{"greenMod", 50000}}), 0xFFE95132U);
  EXPECT_EQ(Resolved(0xE97132, {{"red", 50000}}), 0xFFBC7132U);
  EXPECT_EQ(Resolved(0xE97132, {{"redOff", 20000}}), 0xFFFF7132U);  // clamps at the output
  EXPECT_EQ(Resolved(0xE97132, {{"gamma", std::nullopt}}), 0xFFF5B27AU);
  EXPECT_EQ(Resolved(0xE97132, {{"invGamma", std::nullopt}}), 0xFFD02A08U);
  // Gray weighs the sRGB channels 0.2126 / 0.7152 / 0.0722.
  EXPECT_EQ(Resolved(0xE97132, {{"gray", std::nullopt}}), 0xFF868686U);
  EXPECT_EQ(Resolved(0x7F3F9F, {{"gray", std::nullopt}}), 0xFF545454U);
}

TEST(DrawingColorResolve, HueAndSaturationMatchExcel) {
  EXPECT_EQ(Resolved(0xE97132, {{"hueOff", 5400000}}), 0xFF4FE932U);
  EXPECT_EQ(Resolved(0xE97132, {{"hueOff", -5400000}}), 0xFFCC32E9U);  // wraps
  EXPECT_EQ(Resolved(0xE97132, {{"hueMod", 200000}}), 0xFFE9B032U);
  EXPECT_EQ(Resolved(0xE97132, {{"comp", std::nullopt}}), 0xFF32AAE9U);
  EXPECT_EQ(Resolved(0xE97132, {{"sat", 0}}), 0xFF8E8E8EU);
  // Saturation is not clamped: past 1 the channels clip, below 0 the hue turns to its complement.
  EXPECT_EQ(Resolved(0xE97132, {{"satMod", 150000}}), 0xFFFF6304U);
  EXPECT_EQ(Resolved(0x7F3F9F, {{"satOff", -50000}}), 0xFF6D7767U);
  EXPECT_EQ(Resolved(0x7F3F9F, {{"lum", 0}}), 0xFF000000U);
}

TEST(DrawingColorResolve, TransformsApplyInTheirWrittenOrder) {
  EXPECT_EQ(Resolved(0xE97132, {{"shade", 50000}, {"lumMod", 75000}}), 0xFF803D19U);
  EXPECT_EQ(Resolved(0xE97132, {{"lumMod", 75000}, {"shade", 50000}}), 0xFF8C380CU);
  EXPECT_EQ(Resolved(0xE97132, {{"tint", 50000}, {"satMod", 150000}}), 0xFFFFC1B1U);
  EXPECT_EQ(Resolved(0xE97132, {{"lumMod", 60000}, {"lumOff", 40000}, {"shade", 75000}}), 0xFFD59574U);
  EXPECT_EQ(Resolved(0xE97132, {{"gamma", std::nullopt}, {"shade", 50000}, {"invGamma", std::nullopt}}), 0xFF743919U);
}

TEST(DrawingColorResolve, UnmodelledTransformsHaveNoColour) {
  const std::vector<std::vector<ColorTransform>> cases = {
      {{"alphaMod", 50000}},      {{"fooBar", 1}},           {{"redModx", 1}},
      {{"lumMod", std::nullopt}}, {{"shade", std::nullopt}}, {{"tint", 50000}, {"hueMod", std::nullopt}},
  };
  for (const auto& t : cases) {
    const ResolvedPaint got = drawing_color::resolve_transforms(0xFF123456U, t);
    EXPECT_FALSE(got.argb.has_value()) << t.front().name;
    EXPECT_FALSE(got.alpha.has_value()) << t.front().name;
  }
}

TEST(DrawingColorResolve, AlphaKeepsTheColourAndLeavesTheOpacityToTheModel) {
  const ResolvedPaint plain = drawing_color::resolve_transforms(0xFF123456U, {{"shade", 50000}});
  EXPECT_EQ(plain.alpha, std::optional<std::uint32_t>(100000U));
  const ResolvedPaint alpha = drawing_color::resolve_transforms(0xFF123456U, {{"alpha", 40000}, {"shade", 50000}});
  EXPECT_EQ(alpha.argb, plain.argb);
  EXPECT_FALSE(alpha.alpha.has_value());
}

TEST(DrawingColorRead, BaseTransformsAndRenderedColour) {
  const Workbook wb = Workbook::create();
  const auto shade =
      ReadDrawing(wb, "<a:schemeClr xmlns:a=\"urn:a\" val=\"lt1\"><a:shade val=\"50000\"/></a:schemeClr>");
  ASSERT_TRUE(shade.has_value());
  EXPECT_EQ(shade->base.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(shade->base.theme, 0U);
  EXPECT_EQ(shade->transforms, (std::vector<ColorTransform>{{"shade", 50000}}));
  EXPECT_EQ(shade->resolved_argb, 0xFFBCBCBCU);
  EXPECT_EQ(shade->resolved_alpha, std::optional<std::uint32_t>(100000U));

  const auto kept = ReadDrawing(wb,
                                "<a:srgbClr xmlns:a=\"urn:a\" val=\"7f3f9f\"><a:satMod val=\"-5\"/><a:alpha val=\"1\"/>"
                                "<a:comp/><a:fooBar val=\"7\"/></a:srgbClr>");
  ASSERT_TRUE(kept.has_value());
  EXPECT_EQ(kept->base.kind, ColorSpec::Kind::kRgb);
  EXPECT_EQ(kept->base.rgb & 0xFFFFFFU, 0x7F3F9FU);
  EXPECT_EQ(kept->transforms,
            (std::vector<ColorTransform>{{"satMod", -5}, {"alpha", 1}, {"comp", std::nullopt}, {"fooBar", 7}}));
  EXPECT_FALSE(kept->resolved_argb.has_value());
  EXPECT_FALSE(kept->resolved_alpha.has_value());

  const auto sys = ReadDrawing(wb, "<a:sysClr xmlns:a=\"urn:a\" val=\"windowText\" lastClr=\"101010\"/>");
  ASSERT_TRUE(sys.has_value());
  EXPECT_EQ(sys->base.rgb, 0xFF101010U);
  EXPECT_EQ(sys->resolved_argb, 0xFF101010U);
}

TEST(DrawingColorRead, RemappedSlotsReadAsTheMappedRgb) {
  const Workbook wb = Workbook::create();
  const auto mapped =
      ReadDrawing(wb, "<a:schemeClr xmlns:a=\"urn:a\" val=\"bg1\"><a:lumMod val=\"50000\"/></a:schemeClr>",
                  "<c:clrMapOvr bg1=\"dk1\" tx1=\"lt1\" bg2=\"lt2\" tx2=\"dk2\"/>");
  ASSERT_TRUE(mapped.has_value());
  EXPECT_EQ(mapped->base.kind, ColorSpec::Kind::kRgb);
  EXPECT_EQ(mapped->base.rgb, 0xFF000000U);
  EXPECT_EQ(mapped->transforms.size(), 1U);
  const auto identity = ReadDrawing(wb, "<a:schemeClr xmlns:a=\"urn:a\" val=\"tx2\"/>",
                                    "<c:clrMapOvr bg1=\"dk1\" tx1=\"lt1\" bg2=\"lt2\" tx2=\"dk2\"/>");
  ASSERT_TRUE(identity.has_value());
  EXPECT_EQ(identity->base.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(identity->base.theme, 3U);
}

TEST(DrawingColorRead, UnreadableColours) {
  const Workbook wb = Workbook::create();
  for (const std::string clr : {
           "<a:schemeClr xmlns:a=\"urn:a\" val=\"phClr\"/>",
           "<a:prstClr xmlns:a=\"urn:a\" val=\"red\"/>",
           "<a:srgbClr xmlns:a=\"urn:a\" val=\"12345\"/>",
           "<a:srgbClr xmlns:a=\"urn:a\" val=\"123456\"><a:lumMod val=\"1.5\"/></a:srgbClr>",
           "<a:srgbClr xmlns:a=\"urn:a\" val=\"123456\"><a:lumMod val=\"2147483648\"/></a:srgbClr>",
           "<a:srgbClr xmlns:a=\"urn:a\" xmlns:x=\"urn:x\" val=\"123456\"><x:lumMod val=\"1\"/></a:srgbClr>",
       }) {
    EXPECT_FALSE(ReadDrawing(wb, clr).has_value()) << clr;
  }
  std::string many = "<a:srgbClr xmlns:a=\"urn:a\" val=\"123456\">";
  for (int i = 0; i < 65; ++i) {
    many += "<a:comp/>";
  }
  EXPECT_FALSE(ReadDrawing(wb, many + "</a:srgbClr>").has_value());
}

TEST(DrawingColorWrite, TransformsInOrderAndTintShorthand) {
  EXPECT_EQ(WrittenDrawing(ThemeDrawing(0, {{"shade", 50000}}), "30000"),
            "<a:schemeClr val=\"bg1\"><a:shade val=\"50000\"/><a:alpha val=\"30000\"/></a:schemeClr>");
  EXPECT_EQ(WrittenDrawing(ThemeDrawing(4, {}, 0.4)),
            "<a:schemeClr val=\"accent1\"><a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/></a:schemeClr>");
  DrawingColor rgb;
  rgb.base.kind = ColorSpec::Kind::kRgb;
  rgb.base.rgb = 0x80ABCDEFU;
  rgb.transforms = {{"comp", std::nullopt}, {"lumMod", 0}};
  EXPECT_EQ(WrittenDrawing(rgb), "<a:srgbClr val=\"ABCDEF\"><a:comp/><a:lumMod val=\"0\"/></a:srgbClr>");
}

TEST(DrawingColorWrite, SameColourComparesTheExpandedForm) {
  EXPECT_TRUE(drawing_color::same_drawing_color(ThemeDrawing(4, {}, 0.4),
                                                ThemeDrawing(4, {{"lumMod", 60000}, {"lumOff", 40000}})));
  EXPECT_FALSE(drawing_color::same_drawing_color(ThemeDrawing(4, {{"lumOff", 40000}, {"lumMod", 60000}}),
                                                 ThemeDrawing(4, {{"lumMod", 60000}, {"lumOff", 40000}})));
  EXPECT_FALSE(drawing_color::same_drawing_color(ThemeDrawing(4, {{"shade", 1}}), ThemeDrawing(4, {})));
}

TEST(DrawingColorInput, Refusals) {
  EXPECT_TRUE(static_cast<bool>(drawing_color::validate_input(ThemeDrawing(4, {{"shade", 1}, {"x9", std::nullopt}}))));
  DrawingColor none;
  EXPECT_EQ(drawing_color::validate_input(none).error().code, FormulonErrorCode::kInvalidArgument);
  DrawingColor indexed;
  indexed.base.kind = ColorSpec::Kind::kIndexed;
  EXPECT_EQ(drawing_color::validate_input(indexed).error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(drawing_color::validate_input(ThemeDrawing(12, {})).error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(drawing_color::validate_input(ThemeDrawing(4, {{"shade", 1}}, 0.4)).error().code,
            FormulonErrorCode::kInvalidArgument);
  for (const std::string& name : std::vector<std::string>{"", "9a", "a-b", std::string(33, 'a')}) {
    EXPECT_EQ(drawing_color::validate_input(ThemeDrawing(4, {{name, 1}})).error().code,
              FormulonErrorCode::kInvalidArgument)
        << name;
  }
  EXPECT_TRUE(static_cast<bool>(drawing_color::validate_input(ThemeDrawing(4, {{std::string(32, 'a'), 1}}))));
  EXPECT_EQ(drawing_color::validate_input(ThemeDrawing(4, std::vector<ColorTransform>(65, {"comp", std::nullopt})))
                .error()
                .code,
            FormulonErrorCode::kSecurityResourceLimit);
}

}  // namespace
}  // namespace formulon
