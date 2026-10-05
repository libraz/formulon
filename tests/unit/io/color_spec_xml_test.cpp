#include "io/color_spec_xml.h"

#include <cstdint>
#include <string>

#include "gtest/gtest.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

ColorSpec Read(const char* xml) {
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_string(xml));
  return read_color_spec(doc.first_child());
}

std::string Write(const ColorSpec& spec, std::uint32_t fallback = 0xFF000000U) {
  std::string out;
  append_color_spec_attrs(out, spec, fallback);
  return out;
}

TEST(ColorSpecXml, RgbRoundTrips) {
  const ColorSpec spec = Read("<color rgb=\"FF336699\"/>");
  EXPECT_EQ(spec.kind, ColorSpec::Kind::kRgb);
  EXPECT_EQ(spec.rgb, 0xFF336699U);
  EXPECT_EQ(Write(spec), " rgb=\"FF336699\"");
}

TEST(ColorSpecXml, ThemeWithTintRoundTrips) {
  const ColorSpec spec = Read("<color theme=\"4\" tint=\"-0.25\"/>");
  EXPECT_EQ(spec.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(spec.theme, 4U);
  EXPECT_DOUBLE_EQ(spec.tint, -0.25);
  EXPECT_EQ(Write(spec), " theme=\"4\" tint=\"-0.25\"");
}

TEST(ColorSpecXml, ThemeWithoutTintOmitsTint) {
  EXPECT_EQ(Write(Read("<color theme=\"1\"/>")), " theme=\"1\"");
}

TEST(ColorSpecXml, IndexedRoundTrips) {
  const ColorSpec spec = Read("<color indexed=\"10\"/>");
  EXPECT_EQ(spec.kind, ColorSpec::Kind::kIndexed);
  EXPECT_EQ(spec.indexed, 10U);
  EXPECT_EQ(Write(spec), " indexed=\"10\"");
}

TEST(ColorSpecXml, AutoRoundTrips) {
  const ColorSpec spec = Read("<color auto=\"1\"/>");
  EXPECT_EQ(spec.kind, ColorSpec::Kind::kAuto);
  EXPECT_EQ(Write(spec), " auto=\"1\"");
}

TEST(ColorSpecXml, AbsentOrEmptyIsNoneAndWritesFallback) {
  EXPECT_EQ(Read("<color/>").kind, ColorSpec::Kind::kNone);
  EXPECT_EQ(read_color_spec(pugi::xml_node()).kind, ColorSpec::Kind::kNone);
  EXPECT_EQ(Write(ColorSpec{}, 0x80112233U), " rgb=\"80112233\"");
}

}  // namespace
}  // namespace formulon::io
