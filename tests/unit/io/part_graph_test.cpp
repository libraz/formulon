#include "io/ooxml/part_graph.h"

#include <cstddef>
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

TEST(PartGraph, ParsesPrefixedRelationshipsAndIgnoresForeignNamespaceSiblings) {
  const std::string xml =
      "<p:Relationships xmlns:p=\"http://schemas.openxmlformats.org/package/2006/relationships\" "
      "xmlns:f=\"urn:foreign\"><p:Relationship Id=\"good\" Type=\"urn:t\" Target=\"child.xml\"/>"
      "<f:Relationship Id=\"foreign\" Type=\"urn:t\" Target=\"ignored.xml\"/>"
      "<xml:Relationship Id=\"builtin-foreign\" Type=\"urn:t\" Target=\"ignored.xml\"/>"
      "</p:Relationships>";

  auto rels = io::ooxml::parse_part_rels(ToBytes(xml), "xl/parts");

  ASSERT_TRUE(static_cast<bool>(rels));
  ASSERT_EQ(rels.value().size(), 1U);
  EXPECT_EQ(rels.value()[0].id, "good");
  EXPECT_EQ(rels.value()[0].target, "xl/parts/child.xml");
}

TEST(PartGraph, RejectsRelationshipsRootInAnUnrelatedNamespace) {
  const std::string xml =
      "<p:Relationships xmlns:p=\"urn:not-opc\"><p:Relationship Id=\"rId1\" Type=\"urn:t\" "
      "Target=\"child.xml\"/></p:Relationships>";

  auto rels = io::ooxml::parse_part_rels(ToBytes(xml), "xl/parts");

  ASSERT_FALSE(static_cast<bool>(rels));
  EXPECT_EQ(rels.error().code, FormulonErrorCode::kIoDrawingUnparseable);
}

TEST(PartGraph, RejectsUnboundRelationshipPrefixAndMalformedPrefixedRoot) {
  for (const std::string& xml :
       {"<Relationships><p:Relationship Id=\"rId1\" Type=\"urn:t\" Target=\"child.xml\"/>"
        "</Relationships>",
        "<p:Relationships><p:Relationship Id=\"rId1\" Type=\"urn:t\" Target=\"child.xml\"/>"
        "</p:Relationships>",
        "<:Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>"}) {
    auto rels = io::ooxml::parse_part_rels(ToBytes(xml), "xl/parts");
    ASSERT_FALSE(static_cast<bool>(rels)) << xml;
    EXPECT_EQ(rels.error().code, FormulonErrorCode::kIoDrawingUnparseable) << xml;
  }
}

TEST(PartGraph, UnboundRelationshipMakesOrphanRemovalConservative) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({PassthroughPart("xl/parts/orphan.xml", "", ToBytes("orphan")),
                            PassthroughPart("xl/_rels/broken.xml.rels", "",
                                            ToBytes("<Relationships><p:Relationship Target=\"orphan.xml\"/>"
                                                    "</Relationships>"))});

  io::ooxml::remove_orphans(wb, {"xl/parts/orphan.xml"});

  EXPECT_TRUE(Has(wb, "xl/parts/orphan.xml"));
  EXPECT_TRUE(Has(wb, "xl/_rels/broken.xml.rels"));
}

TEST(PartGraph, InvalidRelationshipsRootMakesOrphanRemovalConservative) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts(
      {PassthroughPart("xl/parts/orphan.xml", "", ToBytes("orphan")),
       PassthroughPart("xl/_rels/wrong-root.xml.rels", "", ToBytes("<p:Relationships xmlns:p=\"urn:not-opc\"/>"))});

  io::ooxml::remove_orphans(wb, {"xl/parts/orphan.xml"});

  EXPECT_TRUE(Has(wb, "xl/parts/orphan.xml"));
  EXPECT_TRUE(Has(wb, "xl/_rels/wrong-root.xml.rels"));
}

TEST(PartGraph, PrefixedLiveRelationshipKeepsTargetAndItsCascade) {
  const auto prefixed_rels = [](const char* target) {
    return std::string("<p:Relationships xmlns:p=\"http://schemas.openxmlformats.org/package/2006/relationships\">") +
           "<p:Relationship Id=\"rId1\" Type=\"urn:t\" Target=\"" + target + "\"/></p:Relationships>";
  };
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({
      PassthroughPart("xl/parts/remove.xml", "", ToBytes("remove")),
      PassthroughPart("xl/parts/child.xml", "", ToBytes("child")),
      PassthroughPart("xl/parts/grandchild.xml", "", ToBytes("grandchild")),
      PassthroughPart("xl/parts/holder.xml", "", ToBytes("holder")),
      PassthroughPart("xl/parts/_rels/remove.xml.rels", "", ToBytes(prefixed_rels("child.xml"))),
      PassthroughPart("xl/parts/_rels/child.xml.rels", "", ToBytes(prefixed_rels("grandchild.xml"))),
      PassthroughPart("xl/parts/_rels/holder.xml.rels", "", ToBytes(prefixed_rels("child.xml"))),
  });

  // Queue the former owner's target too; processing the owner first removes
  // its inbound edge, leaving only the live prefixed holder to protect child.
  io::ooxml::remove_orphans(wb, {"xl/parts/child.xml", "xl/parts/remove.xml"});

  EXPECT_FALSE(Has(wb, "xl/parts/remove.xml"));
  EXPECT_FALSE(Has(wb, "xl/parts/_rels/remove.xml.rels"));
  EXPECT_TRUE(Has(wb, "xl/parts/child.xml"));
  EXPECT_TRUE(Has(wb, "xl/parts/_rels/child.xml.rels"));
  EXPECT_TRUE(Has(wb, "xl/parts/grandchild.xml"));
  EXPECT_TRUE(Has(wb, "xl/parts/holder.xml"));
  EXPECT_TRUE(Has(wb, "xl/parts/_rels/holder.xml.rels"));
}

TEST(PartGraph, MalformedRelsAreRefused) {
  auto bad = io::ooxml::parse_part_rels(ToBytes("<Relationships><Relationship"), "xl/drawings");
  ASSERT_FALSE(static_cast<bool>(bad));
  EXPECT_EQ(bad.error().code, FormulonErrorCode::kIoDrawingUnparseable);
  auto escaping = io::ooxml::parse_part_rels(ToBytes(RelsTo("../../../x.xml")), "xl/drawings");
  ASSERT_FALSE(static_cast<bool>(escaping));
  EXPECT_EQ(escaping.error().code, FormulonErrorCode::kIoDrawingUnparseable);
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

TEST(PartGraph, MalformedRelsMakesRemovalConservativeForUnrelatedParts) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({PassthroughPart("xl/orphan.xml", "", ToBytes("orphan")),
                            PassthroughPart("xl/other.xml", "", ToBytes("other")),
                            PassthroughPart("xl/_rels/broken.xml.rels", "", ToBytes("<Relationships><Relationship"))});

  io::ooxml::remove_orphans(wb, {"xl/orphan.xml"});

  ASSERT_EQ(wb.passthrough_parts().size(), 3U);
  EXPECT_TRUE(Has(wb, "xl/orphan.xml"));
  EXPECT_TRUE(Has(wb, "xl/other.xml"));
  EXPECT_TRUE(Has(wb, "xl/_rels/broken.xml.rels"));
}

TEST(PartGraph, SharedChildRemainsUntilAllRelationshipPartsAreRemoved) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({
      PassthroughPart("xl/parts/a.xml", "", ToBytes("a")),
      PassthroughPart("xl/parts/b.xml", "", ToBytes("b")),
      PassthroughPart("xl/parts/shared.xml", "", ToBytes("shared")),
      PassthroughPart("xl/parts/_rels/a.xml.rels", "", ToBytes(RelsTo("shared.xml"))),
      PassthroughPart("xl/parts/_rels/b.xml.rels", "", ToBytes(RelsTo("shared.xml"))),
  });

  io::ooxml::remove_orphans(wb, {"xl/parts/a.xml", "xl/parts/b.xml"});

  EXPECT_FALSE(Has(wb, "xl/parts/a.xml"));
  EXPECT_FALSE(Has(wb, "xl/parts/b.xml"));
  EXPECT_FALSE(Has(wb, "xl/parts/_rels/a.xml.rels"));
  EXPECT_FALSE(Has(wb, "xl/parts/_rels/b.xml.rels"));
  EXPECT_FALSE(Has(wb, "xl/parts/shared.xml"));
}

TEST(PartGraph, MutuallyReferencingCandidatesBothSurvive) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({
      PassthroughPart("xl/parts/a.xml", "", ToBytes("a")),
      PassthroughPart("xl/parts/b.xml", "", ToBytes("b")),
      PassthroughPart("xl/parts/_rels/a.xml.rels", "", ToBytes(RelsTo("b.xml"))),
      PassthroughPart("xl/parts/_rels/b.xml.rels", "", ToBytes(RelsTo("a.xml"))),
  });

  io::ooxml::remove_orphans(wb, {"xl/parts/a.xml", "xl/parts/b.xml"});

  ASSERT_EQ(wb.passthrough_parts().size(), 4U);
  EXPECT_TRUE(Has(wb, "xl/parts/a.xml"));
  EXPECT_TRUE(Has(wb, "xl/parts/b.xml"));
}

TEST(PartGraph, ModelRootsAndExternalRelationshipsKeepOnlyTheirInternalTargets) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_drawing_rel_target("xl/parts/root.xml");
  wb.set_unknown_workbook_rels({UnknownRelationship{"rId3", "urn:t", "theme/theme.xml", false}});
  const std::string root_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" Type=\"urn:t\" Target=\"root-child.xml\"/>"
      "<Relationship Id=\"rId2\" Type=\"urn:t\" Target=\"xl/parts/external.xml\" "
      "TargetMode=\"External\"/>"
      "</Relationships>";
  wb.set_passthrough_parts({
      PassthroughPart("xl/parts/root.xml", "", ToBytes("root")),
      PassthroughPart("xl/parts/root-child.xml", "", ToBytes("root-child")),
      PassthroughPart("xl/parts/external.xml", "", ToBytes("external")),
      PassthroughPart("theme/theme.xml", "", ToBytes("theme")),
      PassthroughPart("xl/theme/theme.xml", "", ToBytes("xl-theme")),
      PassthroughPart("xl/parts/_rels/root.xml.rels", "", ToBytes(root_rels)),
  });

  io::ooxml::remove_orphans(wb, {"xl/parts/root.xml", "xl/parts/root-child.xml", "xl/parts/external.xml",
                                 "theme/theme.xml", "xl/theme/theme.xml"});

  EXPECT_TRUE(Has(wb, "xl/parts/root.xml"));
  EXPECT_TRUE(Has(wb, "xl/parts/root-child.xml"));
  EXPECT_FALSE(Has(wb, "xl/parts/external.xml"));
  EXPECT_TRUE(Has(wb, "xl/parts/_rels/root.xml.rels"));
  EXPECT_TRUE(Has(wb, "theme/theme.xml"));
  EXPECT_TRUE(Has(wb, "xl/theme/theme.xml"));
}

TEST(PartGraph, DuplicatePathsUseFirstLivePartAndMatchingFirstSidecar) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({
      PassthroughPart("xl/parts/duplicate.xml", "", ToBytes("first")),
      PassthroughPart("xl/parts/duplicate.xml", "", ToBytes("second")),
      PassthroughPart("xl/parts/child1.xml", "", ToBytes("child1")),
      PassthroughPart("xl/parts/child2.xml", "", ToBytes("child2")),
      PassthroughPart("xl/parts/_rels/duplicate.xml.rels", "", ToBytes(RelsTo("child1.xml"))),
      PassthroughPart("xl/parts/_rels/duplicate.xml.rels", "", ToBytes(RelsTo("child2.xml"))),
  });

  io::ooxml::remove_orphans(wb, {"xl/parts/duplicate.xml"});

  ASSERT_EQ(wb.passthrough_parts().size(), 3U);
  EXPECT_EQ(wb.passthrough_parts()[0].path, "xl/parts/duplicate.xml");
  EXPECT_EQ(std::string(wb.passthrough_parts()[0].bytes.begin(), wb.passthrough_parts()[0].bytes.end()), "second");
  EXPECT_EQ(wb.passthrough_parts()[1].path, "xl/parts/child2.xml");
  EXPECT_EQ(wb.passthrough_parts()[2].path, "xl/parts/_rels/duplicate.xml.rels");
  EXPECT_EQ(std::string(wb.passthrough_parts()[2].bytes.begin(), wb.passthrough_parts()[2].bytes.end()),
            RelsTo("child2.xml"));
}

TEST(PartGraph, CascadeKeepsLifoCandidateProcessing) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({
      PassthroughPart("xl/parts/root.xml", "", ToBytes("root")),
      PassthroughPart("xl/parts/first.xml", "", ToBytes("first")),
      PassthroughPart("xl/parts/second.xml", "", ToBytes("second")),
      PassthroughPart("xl/parts/second-child.xml", "", ToBytes("second-child")),
      PassthroughPart("xl/parts/_rels/root.xml.rels", "",
                      ToBytes("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                              "<Relationship Id=\"rId1\" Type=\"urn:t\" Target=\"first.xml\"/>"
                              "<Relationship Id=\"rId2\" Type=\"urn:t\" Target=\"second.xml\"/>"
                              "</Relationships>")),
      PassthroughPart("xl/parts/_rels/second.xml.rels", "", ToBytes(RelsTo("second-child.xml"))),
  });

  io::ooxml::remove_orphans(wb, {"xl/parts/root.xml"});

  EXPECT_FALSE(Has(wb, "xl/parts/root.xml"));
  EXPECT_FALSE(Has(wb, "xl/parts/first.xml"));
  EXPECT_FALSE(Has(wb, "xl/parts/second.xml"));
  EXPECT_FALSE(Has(wb, "xl/parts/second-child.xml"));
  EXPECT_FALSE(Has(wb, "xl/parts/_rels/root.xml.rels"));
  EXPECT_FALSE(Has(wb, "xl/parts/_rels/second.xml.rels"));
}

TEST(PartGraph, StableCompactionPreservesOrderAndPayloadBuffers) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({PassthroughPart("xl/remove.xml", "", ToBytes("remove")),
                            PassthroughPart("xl/keep.xml", "", ToBytes("keep")),
                            PassthroughPart("xl/tail.xml", "", ToBytes("tail"))});
  const std::uint8_t* keep_bytes = wb.passthrough_parts()[1].bytes.data();

  io::ooxml::remove_orphans(wb, {"xl/remove.xml"});

  ASSERT_EQ(wb.passthrough_parts().size(), 2U);
  EXPECT_EQ(wb.passthrough_parts()[0].path, "xl/keep.xml");
  EXPECT_EQ(wb.passthrough_parts()[1].path, "xl/tail.xml");
  EXPECT_EQ(wb.passthrough_parts()[0].bytes.data(), keep_bytes);
  EXPECT_EQ(std::string(wb.passthrough_parts()[0].bytes.begin(), wb.passthrough_parts()[0].bytes.end()), "keep");
}

TEST(PartGraph, ErasePassthroughPartsCanBorrowMaskFromAnExistingPayload) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({PassthroughPart("a", "", Bytes{1, 0, 0}), PassthroughPart("b", "", Bytes{42}),
                            PassthroughPart("c", "", Bytes{43})});
  ASSERT_TRUE(wb.erase_passthrough_parts(wb.passthrough_parts()[0].bytes));
  ASSERT_EQ(wb.passthrough_parts().size(), 2U);
  EXPECT_EQ(wb.passthrough_parts()[0].path, "b");
  EXPECT_EQ(wb.passthrough_parts()[1].path, "c");
  EXPECT_EQ(wb.passthrough_parts()[0].bytes, Bytes{42});
  EXPECT_EQ(wb.passthrough_parts()[1].bytes, Bytes{43});
}

TEST(PartGraph, ErasePassthroughPartsRejectsMismatchedMask) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({PassthroughPart("xl/one.xml", "", ToBytes("one"))});

  auto result = wb.erase_passthrough_parts({});

  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kInvalidArgument);
  ASSERT_EQ(wb.passthrough_parts().size(), 1U);
  EXPECT_EQ(wb.passthrough_parts()[0].path, "xl/one.xml");
}

}  // namespace
}  // namespace formulon
