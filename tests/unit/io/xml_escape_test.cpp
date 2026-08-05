//
// Unit tests for the element-text vs. attribute-value XML escaping split.

#include "io/xml_escape.h"

#include <string>

#include "gtest/gtest.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

TEST(XmlEscape, ElementTextEscapesCriticalCharsAndControls) {
  std::string out;
  AppendXmlEscaped(out, "a&b<c>\"d'e\n\tf\r\x01");
  EXPECT_EQ(out, "a&amp;b&lt;c&gt;&quot;d&apos;e_x000A__x0009_f_x000D__x0001_");
}

TEST(XmlEscape, AttrEscapesCriticalCharsAndWhitespaceControls) {
  std::string out;
  AppendXmlAttrEscaped(out, "a&b<c>\"d'e\n\tf\r");
  EXPECT_EQ(out, "a&amp;b&lt;c&gt;&quot;d&apos;e_x000A__x0009_f_x000D_");
}

TEST(XmlEscape, OoxmlControlAndLiteralEscapeRoundTrip) {
  const std::string original = "a\r\x01_x000D_\xF0\x9F\x98\x80";
  std::string escaped;
  AppendXmlEscaped(escaped, original);
  EXPECT_EQ(escaped, "a_x000D__x0001__x005F_x000D_\xF0\x9F\x98\x80");

  std::string unescaped;
  AppendOoxmlTextUnescaped(unescaped, escaped);
  EXPECT_EQ(unescaped, original);
}

TEST(XmlEscape, ReplacesInvalidUtf8BeforeWritingXml) {
  std::string escaped;
  AppendXmlEscaped(escaped, "a\xC3\x28\x80");
  EXPECT_EQ(escaped, "a\xEF\xBF\xBD(\xEF\xBF\xBD");

  pugi::xml_document doc;
  const std::string xml = "<t>" + escaped + "</t>";
  ASSERT_TRUE(doc.load_buffer(xml.data(), xml.size()));
  EXPECT_EQ(doc.child("t").text().get(), escaped);
}

TEST(XmlEscape, RichTextReaderDecodesOoxmlEscapes) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<si><t>a_x000D__x005F_x000D_</t></si>"));

  std::string text;
  EXPECT_EQ(append_rich_text(doc.child("si"), text), 1U);
  EXPECT_EQ(text, "a\r_x000D_");
}

TEST(XmlEscape, AttrEscapedNewlineSurvivesPugixmlRoundtrip) {
  // Attribute-value normalisation (mandatory for any conforming XML
  // parser, including pugixml) replaces a *literal* newline inside an
  // attribute value with a space. Only a character reference for it
  // survives intact -- this is the whole point of AppendXmlAttrEscaped.
  std::string attr_value;
  AppendXmlAttrEscaped(attr_value, "line one\nline two\ttabbed");

  std::string xml = "<r name=\"";
  xml += attr_value;
  xml += "\"/>";

  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml.c_str()));
  const std::string_view roundtripped = doc.child("r").attribute("name").value();
  EXPECT_EQ(roundtripped, "line one\nline two\ttabbed");
}

TEST(XmlEscape, PlainEscapeNewlineDoesNotSurviveAttributeRoundtrip) {
  // Contrast case: using the element-text escaper for an attribute value
  // loses the newline, because the literal character gets normalised
  // away by the XML parser instead of being preserved as `&#10;`.
  std::string attr_value;
  AppendXmlEscaped(attr_value, "line one\nline two");

  std::string xml = "<r name=\"";
  xml += attr_value;
  xml += "\"/>";

  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml.c_str()));
  const std::string_view roundtripped = doc.child("r").attribute("name").value();
  EXPECT_NE(roundtripped, "line one\nline two");
}

TEST(XmlEscape, AttrEscapeNonAsciiPassesThroughVerbatim) {
  std::string out;
  AppendXmlAttrEscaped(out, "\xE6\x97\xA5\xE6\x9C\xAC");  // "日本" in UTF-8.
  EXPECT_EQ(out, "\xE6\x97\xA5\xE6\x9C\xAC");
}

}  // namespace
}  // namespace formulon::io
