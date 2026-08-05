//
// Unit tests for the typed node-attribute accessors in `io/xml_utils.h`.
// The legacy `parse_xml_*_attr` helpers (which take a `pugi::xml_attribute`
// directly) are exercised indirectly by the reader unit tests; this file
// targets the newer `attr_*(node, name, def)` API the readers are
// migrating onto.

#include "io/xml_utils.h"

#include <cstdint>
#include <string_view>

#include "gtest/gtest.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

pugi::xml_document Load(const char* xml) {
  pugi::xml_document doc;
  pugi::xml_parse_result rc = doc.load_string(xml);
  EXPECT_TRUE(rc) << rc.description();
  return doc;
}

TEST(XmlUtilsAttr, StrPresentReturnsAttributeValue) {
  pugi::xml_document doc = Load(R"(<r name="hello"/>)");
  EXPECT_EQ(attr_str(doc.child("r"), "name"), std::string_view("hello"));
}

TEST(XmlUtilsAttr, StrMissingReturnsDefault) {
  pugi::xml_document doc = Load(R"(<r/>)");
  EXPECT_EQ(attr_str(doc.child("r"), "name"), std::string_view());
  EXPECT_EQ(attr_str(doc.child("r"), "name", "fallback"), std::string_view("fallback"));
}

TEST(XmlUtilsAttr, StrEmptyValueReturnsDefault) {
  // An empty `name=""` attribute is treated as absent: callers that
  // care about presence-without-value should still go through
  // `node.attribute(name)` directly.
  pugi::xml_document doc = Load(R"(<r name=""/>)");
  EXPECT_EQ(attr_str(doc.child("r"), "name", "fallback"), std::string_view("fallback"));
}

TEST(XmlUtilsAttr, U32Roundtrip) {
  pugi::xml_document doc = Load(R"(<r id="42" zero="0" big="4294967290"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_EQ(attr_u32(n, "id"), 42U);
  EXPECT_EQ(attr_u32(n, "zero"), 0U);
  EXPECT_EQ(attr_u32(n, "big"), 4294967290U);
  EXPECT_EQ(attr_u32(n, "missing"), 0U);
  EXPECT_EQ(attr_u32(n, "missing", 99U), 99U);
}

TEST(XmlUtilsAttr, U32MalformedFallsBackToDefault) {
  pugi::xml_document doc = Load(R"(<r empty="" junk="abc" neg="-1"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_EQ(attr_u32(n, "empty", 7U), 7U);
  EXPECT_EQ(attr_u32(n, "junk", 7U), 7U);
  // pugi's as_uint() rejects negative input by returning the supplied
  // default — this matches the legacy `parse_xml_u32_attr` behaviour.
  EXPECT_EQ(attr_u32(n, "neg", 7U), 7U);
}

TEST(XmlUtilsAttr, I32Roundtrip) {
  pugi::xml_document doc = Load(R"(<r a="42" b="-1" c="0" d="-2147483648"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_EQ(attr_i32(n, "a"), 42);
  EXPECT_EQ(attr_i32(n, "b"), -1);
  EXPECT_EQ(attr_i32(n, "c"), 0);
  EXPECT_EQ(attr_i32(n, "d"), INT32_MIN);
  EXPECT_EQ(attr_i32(n, "missing", -42), -42);
}

TEST(XmlUtilsAttr, F64Roundtrip) {
  pugi::xml_document doc = Load(R"(<r a="3.14" b="-2.5" c="0" d="1e3"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_DOUBLE_EQ(attr_f64(n, "a"), 3.14);
  EXPECT_DOUBLE_EQ(attr_f64(n, "b"), -2.5);
  EXPECT_DOUBLE_EQ(attr_f64(n, "c"), 0.0);
  EXPECT_DOUBLE_EQ(attr_f64(n, "d"), 1000.0);
  EXPECT_DOUBLE_EQ(attr_f64(n, "missing", 1.5), 1.5);
}

TEST(XmlUtilsAttr, BoolOneAndZero) {
  pugi::xml_document doc = Load(R"(<r a="1" b="0"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_TRUE(attr_bool(n, "a"));
  EXPECT_FALSE(attr_bool(n, "b"));
}

TEST(XmlUtilsAttr, BoolTrueFalseCaseInsensitive) {
  pugi::xml_document doc = Load(R"(<r a="true" b="True" c="TRUE" d="false" e="False" f="FALSE"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_TRUE(attr_bool(n, "a"));
  EXPECT_TRUE(attr_bool(n, "b"));
  EXPECT_TRUE(attr_bool(n, "c"));
  EXPECT_FALSE(attr_bool(n, "d"));
  EXPECT_FALSE(attr_bool(n, "e"));
  EXPECT_FALSE(attr_bool(n, "f"));
}

TEST(XmlUtilsAttr, BoolMissingUsesDefault) {
  pugi::xml_document doc = Load(R"(<r/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_FALSE(attr_bool(n, "absent"));
  EXPECT_TRUE(attr_bool(n, "absent", true));
}

TEST(XmlUtilsAttr, BoolUnknownTokensMapToFalse) {
  // Anything outside {1, 0, true(any case), false(any case)} maps to
  // false. Matches the legacy `parse_xml_bool_attr` lenient lexicon.
  pugi::xml_document doc = Load(R"(<r a="yes" b="2" c="trues" d="tru"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_FALSE(attr_bool(n, "a"));
  EXPECT_FALSE(attr_bool(n, "b"));
  EXPECT_FALSE(attr_bool(n, "c"));
  EXPECT_FALSE(attr_bool(n, "d"));
}

TEST(XmlUtilsAttr, EmptyNodeReturnsDefaults) {
  pugi::xml_node empty;
  EXPECT_EQ(attr_str(empty, "x", "fb"), std::string_view("fb"));
  EXPECT_EQ(attr_u32(empty, "x", 11U), 11U);
  EXPECT_EQ(attr_i32(empty, "x", -3), -3);
  EXPECT_DOUBLE_EQ(attr_f64(empty, "x", 2.5), 2.5);
  EXPECT_TRUE(attr_bool(empty, "x", true));
  EXPECT_FALSE(attr_bool(empty, "x", false));
}

TEST(XmlUtilsAttr, RootExtraAttributesEscapeValuesForReEmission) {
  pugi::xml_document doc =
      Load(R"(<worksheet xmlns="urn:main" xmlns:r="urn:rel" custom="one &amp; two &quot;three&quot; &lt;four&gt;"/>)");

  EXPECT_EQ(capture_root_extra_ns_attrs(doc.document_element()),
            " custom=\"one &amp; two &quot;three&quot; &lt;four&gt;\"");
}

}  // namespace
}  // namespace formulon::io
