// pugixml_smoke_test.cpp
//
// Verifies that the FetchContent-vendored pugixml library is reachable from
// Formulon, that PUGIXML_NO_EXCEPTIONS is in effect (bad XML yields a
// failing xml_parse_result rather than throwing), and that the compact
// build flag is picked up.

#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "pugixml.hpp"

namespace {

TEST(PugixmlSmoke, ParsesSimpleDocument) {
  constexpr std::string_view kXml = "<root><child attr=\"v\">text</child></root>";

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_buffer(kXml.data(), kXml.size());
  ASSERT_TRUE(static_cast<bool>(result)) << "pugixml parse failed: " << result.description();
  EXPECT_EQ(result.status, pugi::status_ok);

  pugi::xml_node root = doc.child("root");
  ASSERT_TRUE(static_cast<bool>(root));

  pugi::xml_node child = root.child("child");
  ASSERT_TRUE(static_cast<bool>(child));

  EXPECT_STREQ(child.attribute("attr").value(), "v");
  EXPECT_STREQ(child.child_value(), "text");
}

TEST(PugixmlSmoke, NoExceptionsOnMalformedInput) {
  // Malformed XML must surface via xml_parse_result::status rather than an
  // exception — Formulon builds with -fno-exceptions / PUGIXML_NO_EXCEPTIONS.
  constexpr std::string_view kBadXml = "<root><child></root>";

  pugi::xml_document doc;
  pugi::xml_parse_result result = doc.load_buffer(kBadXml.data(), kBadXml.size());
  EXPECT_FALSE(static_cast<bool>(result));
  EXPECT_NE(result.status, pugi::status_ok);
}

TEST(PugixmlSmoke, CompactBuildDefinitionVisible) {
  // When FetchPugixml.cmake is picked up correctly we see PUGIXML_COMPACT
  // both as a cache var and as a public compile definition. This guard
  // catches a regression where the alias target forgets to propagate it.
#ifndef PUGIXML_COMPACT
  FAIL() << "PUGIXML_COMPACT must be defined for Formulon builds";
#endif
#ifndef PUGIXML_NO_EXCEPTIONS
  FAIL() << "PUGIXML_NO_EXCEPTIONS must be defined for Formulon builds";
#endif
  SUCCEED();
}

}  // namespace
