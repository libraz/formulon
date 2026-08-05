//
// Unit tests for the element-text vs. attribute-value XML escaping split.

#include "io/xml_escape.h"

#include <string>

#include "gtest/gtest.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

TEST(XmlEscape, ElementTextEscapesCriticalCharsOnly) {
  std::string out;
  AppendXmlEscaped(out, "a&b<c>\"d'e\n\tf\r");
  // TAB / LF / CR pass through verbatim in element-text context; only the
  // five XML-critical characters are entity-escaped.
  EXPECT_EQ(out, "a&amp;b&lt;c&gt;&quot;d&apos;e\n\tf\r");
}

TEST(XmlEscape, AttrEscapesCriticalCharsAndWhitespaceControls) {
  std::string out;
  AppendXmlAttrEscaped(out, "a&b<c>\"d'e\n\tf\r");
  EXPECT_EQ(out, "a&amp;b&lt;c&gt;&quot;d&apos;e&#10;&#9;f&#13;");
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
