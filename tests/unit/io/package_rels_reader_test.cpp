//
// Unit tests for `formulon::io::ooxml::read_unknown_package_rels`, the
// package-level `_rels/.rels` reader shared by the OOXML and XLSB readers.

#include "io/ooxml/package_rels_reader.h"

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"

namespace formulon::io::ooxml {
namespace {

constexpr const char* kRelsOpen =
    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";

std::vector<std::uint8_t> Bytes(const std::string& body) {
  const std::string xml = std::string(kRelsOpen) + body + "</Relationships>";
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

TEST(PackageRelsReader, ExternalTargetKeepsALeadingSlashForEveryReader) {
  const auto bytes =
      Bytes("<Relationship Id=\"rId9\" Type=\"urn:custom\" Target=\"/share/doc.xlsx\" TargetMode=\"External\"/>");
  for (const char* module : {"ooxml_reader", "xlsb_reader"}) {
    auto rels = read_unknown_package_rels(bytes, module);
    ASSERT_TRUE(static_cast<bool>(rels)) << module << ": " << rels.error().message;
    ASSERT_EQ(rels.value().size(), 1U);
    EXPECT_TRUE(rels.value()[0].target_external);
    EXPECT_EQ(rels.value()[0].target, "/share/doc.xlsx") << module;
  }
}

TEST(PackageRelsReader, InternalTargetLosesItsLeadingSlash) {
  auto rels = read_unknown_package_rels(
      Bytes("<Relationship Id=\"rId9\" Type=\"urn:custom\" Target=\"/docProps/x.xml\"/>"), "xlsb_reader");
  ASSERT_TRUE(static_cast<bool>(rels)) << rels.error().message;
  ASSERT_EQ(rels.value().size(), 1U);
  EXPECT_FALSE(rels.value()[0].target_external);
  EXPECT_EQ(rels.value()[0].target, "docProps/x.xml");
}

TEST(PackageRelsReader, WellKnownRootsAreNotUnknown) {
  auto rels = read_unknown_package_rels(
      Bytes("<Relationship Id=\"rId1\" "
            "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
            "Target=\"xl/workbook.xml\"/>"
            "<Relationship Id=\"rId2\" Type=\"urn:custom\" Target=\"custom/a.xml\"/>"),
      "ooxml_reader");
  ASSERT_TRUE(static_cast<bool>(rels)) << rels.error().message;
  ASSERT_EQ(rels.value().size(), 1U);
  EXPECT_EQ(rels.value()[0].id, "rId2");
}

TEST(PackageRelsReader, InternalTargetEscapingThePackageIsRefused) {
  auto rels = read_unknown_package_rels(Bytes("<Relationship Id=\"rId9\" Type=\"urn:custom\" Target=\"../../x\"/>"),
                                        "xlsb_reader");
  ASSERT_FALSE(static_cast<bool>(rels));
  EXPECT_EQ(rels.error().code, FormulonErrorCode::kIoZipSlip);
  EXPECT_NE(rels.error().context.find("context=xlsb_reader"), std::string::npos);
}

}  // namespace
}  // namespace formulon::io::ooxml
