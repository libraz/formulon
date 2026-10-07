#include "io/ooxml/part_graph.h"

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "passthrough_part.h"
#include "unknown_relationship.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kRels =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" "
    "Target=\"../charts/chart1.xml\"/>"
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
    "Target=\"../media/image1.png\"/>"
    "<Relationship Id=\"rId3\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
    "Target=\"https://example.com/a\" TargetMode=\"External\"/>"
    "</Relationships>";

Bytes ToBytes(const std::string& s) {
  return Bytes(s.begin(), s.end());
}

std::string RelsTo(const std::string& target) {
  return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
         "<Relationship Id=\"rId1\" Type=\"urn:t\" Target=\"" +
         target + "\"/></Relationships>";
}

bool Has(const Workbook& wb, const std::string& path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return true;
    }
  }
  return false;
}

TEST(PartGraph, RelsResolveAgainstOwnerDirectory) {
  auto rels = io::ooxml::parse_part_rels(ToBytes(kRels), "xl/drawings");
  ASSERT_TRUE(static_cast<bool>(rels));
  ASSERT_EQ(rels.value().size(), 3U);
  EXPECT_EQ(rels.value()[0].id, "rId1");
  EXPECT_EQ(rels.value()[0].target, "xl/charts/chart1.xml");
  EXPECT_FALSE(rels.value()[0].external);
  EXPECT_EQ(rels.value()[1].target, "xl/media/image1.png");
  EXPECT_EQ(rels.value()[2].target, "https://example.com/a");
  EXPECT_TRUE(rels.value()[2].external);
}

TEST(PartGraph, MalformedRelsAreRefused) {
  auto bad = io::ooxml::parse_part_rels(ToBytes("<Relationships><Relationship"), "xl/drawings");
  ASSERT_FALSE(static_cast<bool>(bad));
  EXPECT_EQ(bad.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  auto escaping = io::ooxml::parse_part_rels(ToBytes(RelsTo("../../../x.xml")), "xl/drawings");
  ASSERT_FALSE(static_cast<bool>(escaping));
  EXPECT_EQ(escaping.error().code, FormulonErrorCode::kIoDrawingUnparseable);
}

TEST(PartGraph, ReferencedCountsSheetWorkbookAndRelsLinks) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_drawing_rel_target("xl/drawings/drawing1.xml");
  wb.set_unknown_workbook_rels({UnknownRelationship{"rId9", "urn:t", "theme/theme1.xml", false}});
  ASSERT_TRUE(static_cast<bool>(
      wb.add_passthrough_part(PassthroughPart("xl/drawings/_rels/drawing1.xml.rels", "", ToBytes(kRels)))));
  const std::vector<PassthroughPart>& parts = wb.passthrough_parts();
  EXPECT_TRUE(io::ooxml::referenced(wb, parts, "xl/drawings/drawing1.xml"));
  EXPECT_TRUE(io::ooxml::referenced(wb, parts, "xl/theme/theme1.xml"));
  EXPECT_TRUE(io::ooxml::referenced(wb, parts, "xl/media/image1.png"));
  EXPECT_FALSE(io::ooxml::referenced(wb, parts, "xl/media/image2.png"));
  EXPECT_FALSE(io::ooxml::referenced(wb, parts, "https://example.com/a"));
}

TEST(PartGraph, RemoveOrphansFollowsTheChain) {
  // a.xml -> (its rels) -> b.xml -> (its rels) -> c.xml; kept.xml is also
  // reached from a rels part that stays.
  Workbook wb = Workbook::create();
  for (const char* path :
       {"xl/parts/a.xml", "xl/parts/b.xml", "xl/parts/c.xml", "xl/parts/kept.xml", "xl/parts/holder.xml"}) {
    ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(path, "application/x", ToBytes("<x/>")))));
  }
  const auto add_rels = [&wb](const char* path, const std::string& target) {
    ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(path, "", ToBytes(RelsTo(target))))));
  };
  add_rels("xl/parts/_rels/a.xml.rels", "b.xml");
  add_rels("xl/parts/_rels/b.xml.rels", "c.xml");
  add_rels("xl/parts/_rels/holder.xml.rels", "kept.xml");
  wb.set_unknown_workbook_rels({UnknownRelationship{"rId9", "urn:t", "parts/holder.xml", false}});

  io::ooxml::remove_orphans(wb, {"xl/parts/a.xml", "xl/parts/kept.xml", "xl/parts/holder.xml"});
  for (const char* gone : {"xl/parts/a.xml", "xl/parts/_rels/a.xml.rels", "xl/parts/b.xml", "xl/parts/_rels/b.xml.rels",
                           "xl/parts/c.xml"}) {
    EXPECT_FALSE(Has(wb, gone)) << gone;
  }
  for (const char* kept : {"xl/parts/kept.xml", "xl/parts/holder.xml", "xl/parts/_rels/holder.xml.rels"}) {
    EXPECT_TRUE(Has(wb, kept)) << kept;
  }
}

TEST(PartGraph, RemoveOrphansKeepsBytesWhenNothingGoes) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_drawing_rel_target("xl/drawings/drawing1.xml");
  ASSERT_TRUE(static_cast<bool>(
      wb.add_passthrough_part(PassthroughPart("xl/drawings/drawing1.xml", "application/x", ToBytes("<x/>")))));
  io::ooxml::remove_orphans(wb, {"xl/drawings/drawing1.xml", "xl/missing.xml"});
  EXPECT_TRUE(Has(wb, "xl/drawings/drawing1.xml"));
  EXPECT_EQ(wb.passthrough_parts().size(), 1U);
}

}  // namespace
}  // namespace formulon
