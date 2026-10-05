#include "io/ooxml/part_dom.h"

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace ooxml {
namespace {

constexpr const char* kPath = "xl/custom/part1.xml";

std::vector<std::uint8_t> Bytes(const std::string& s) {
  return std::vector<std::uint8_t>(s.begin(), s.end());
}

std::string Text(const std::vector<std::uint8_t>& bytes) {
  return std::string(bytes.begin(), bytes.end());
}

Workbook WorkbookWithPart(const std::string& xml) {
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(kPath, "application/xml", Bytes(xml)))));
  return wb;
}

TEST(PartDom, FindPassthroughPartByPath) {
  Workbook wb = WorkbookWithPart("<r/>");
  const PassthroughPart* part = find_passthrough_part(wb, kPath);
  ASSERT_NE(part, nullptr);
  EXPECT_EQ(part->path, kPath);
  EXPECT_EQ(find_passthrough_part(wb, "xl/custom/other.xml"), nullptr);
}

TEST(PartDom, EditRoundTripsThroughOneDocument) {
  Workbook wb = WorkbookWithPart(
      "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<x:r xmlns:x=\"urn:x\">\n  <x:a v=\"1\"/>\n  <x:t> </x:t>\n</x:r>");
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(load_part_dom(wb, kPath, "test", doc)));
  doc.document_element().child("x:a").attribute("v").set_value("2");
  doc.document_element().append_child("x:b");
  ASSERT_TRUE(static_cast<bool>(store_part_dom(wb, kPath, doc)));
  // Indentation-only text is dropped, a single-space leaf is kept, and the
  // declaration is rewritten in the standalone form.
  EXPECT_EQ(Text(find_passthrough_part(wb, kPath)->bytes),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            "<x:r xmlns:x=\"urn:x\"><x:a v=\"2\"/><x:t> </x:t><x:b/></x:r>");
}

TEST(PartDom, StoreKeepsContentType) {
  Workbook wb = WorkbookWithPart("<r/>");
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(load_part_dom(wb, kPath, "test", doc)));
  ASSERT_TRUE(static_cast<bool>(store_part_dom(wb, kPath, doc)));
  EXPECT_EQ(find_passthrough_part(wb, kPath)->content_type, "application/xml");
}

TEST(PartDom, MissingPartIsInvalidArgument) {
  Workbook wb = Workbook::create();
  pugi::xml_document doc;
  auto loaded = load_part_dom(wb, kPath, "test", doc);
  ASSERT_FALSE(static_cast<bool>(loaded));
  EXPECT_EQ(loaded.error().code, FormulonErrorCode::kInvalidArgument);
  auto stored = store_part_dom(wb, kPath, doc);
  ASSERT_FALSE(static_cast<bool>(stored));
  EXPECT_EQ(stored.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(PartDom, MalformedPartIsXmlParseError) {
  Workbook wb = WorkbookWithPart("<r><unclosed></r>");
  pugi::xml_document doc;
  auto loaded = load_part_dom(wb, kPath, "test", doc);
  ASSERT_FALSE(static_cast<bool>(loaded));
  EXPECT_EQ(loaded.error().code, FormulonErrorCode::kIoXmlParse);
  EXPECT_NE(loaded.error().context.find("context=test"), std::string::npos);
  EXPECT_NE(loaded.error().context.find(kPath), std::string::npos);
}

}  // namespace
}  // namespace ooxml
}  // namespace io
}  // namespace formulon
