//
// Unit tests for the tri-state XSD boolean helpers. The lexical parser is
// the interesting part: "absent" must fall back to the caller's default,
// the four lexical forms (`true` / `false` / `1` / `0`) must parse
// case-insensitively, and anything outside the lexical space must also
// fall back rather than guess.

#include "io/xsd_bool.h"

#include <string>

#include "gtest/gtest.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace {

TEST(XsdBool, ParsesCanonicalForms) {
  EXPECT_TRUE(parse_xsd_bool("1", false));
  EXPECT_TRUE(parse_xsd_bool("true", false));
  EXPECT_FALSE(parse_xsd_bool("0", true));
  EXPECT_FALSE(parse_xsd_bool("false", true));
}

TEST(XsdBool, AlphabeticFormsAreCaseInsensitive) {
  EXPECT_TRUE(parse_xsd_bool("True", false));
  EXPECT_TRUE(parse_xsd_bool("TRUE", false));
  EXPECT_FALSE(parse_xsd_bool("False", true));
  EXPECT_FALSE(parse_xsd_bool("FALSE", true));
}

TEST(XsdBool, EmptyAndMalformedFallBackToDefault) {
  EXPECT_FALSE(parse_xsd_bool("", false));
  EXPECT_TRUE(parse_xsd_bool("", true));
  // Out of the lexical space -> default, no guessing.
  EXPECT_TRUE(parse_xsd_bool("yes", true));
  EXPECT_FALSE(parse_xsd_bool("yes", false));
  EXPECT_TRUE(parse_xsd_bool("2", true));
  EXPECT_FALSE(parse_xsd_bool("2", false));
}

TEST(XsdBool, ReadFromNodeHonoursAbsentDefault) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<e a=\"1\" b=\"0\" c=\"true\" d=\"false\"/>"));
  const pugi::xml_node e = doc.child("e");
  // Present attributes take their explicit value regardless of default.
  EXPECT_TRUE(read_xsd_bool(e, "a", false));
  EXPECT_FALSE(read_xsd_bool(e, "b", true));
  EXPECT_TRUE(read_xsd_bool(e, "c", false));
  EXPECT_FALSE(read_xsd_bool(e, "d", true));
  // Absent attribute -> the caller's default (both polarities).
  EXPECT_TRUE(read_xsd_bool(e, "missing", true));
  EXPECT_FALSE(read_xsd_bool(e, "missing", false));
}

TEST(XsdBool, EmitAlwaysExplicit) {
  std::string out;
  emit_xsd_bool_attr(out, "lockStructure", true);
  emit_xsd_bool_attr(out, "lockWindows", false);
  // Leading space on each; always the explicit 0 / 1 form so a
  // default-true attribute set to false survives the round trip.
  EXPECT_EQ(out, " lockStructure=\"1\" lockWindows=\"0\"");
}

}  // namespace
}  // namespace io
}  // namespace formulon
