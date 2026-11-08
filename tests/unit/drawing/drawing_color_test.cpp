#include "drawing/drawing_color.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>

#include "gtest/gtest.h"
#include "pugixml.hpp"

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

}  // namespace
}  // namespace formulon
