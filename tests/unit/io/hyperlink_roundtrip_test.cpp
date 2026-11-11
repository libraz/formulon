//
// Reader-side tests for `<hyperlinks>` and the rid-to-target join. The
// writer side is exercised by `tests/integration/sheet_features_roundtrip_test.cpp`.

#include <cstdint>
#include <string>
#include <unordered_map>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/sheet_overlay_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

pugi::xml_node ParseWorksheet(pugi::xml_document& doc, const std::string& body) {
  const std::string xml = "<worksheet>" + body + "</worksheet>";
  EXPECT_TRUE(doc.load_string(xml.c_str()));
  return doc.child("worksheet");
}

TEST(HyperlinkRoundTrip, EmptyHyperlinks) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc, "<sheetData/>");
  auto out = read_hyperlinks(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  EXPECT_TRUE(out.value().empty());
}

TEST(HyperlinkRoundTrip, ExternalLinkPreservesRid) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                           "<hyperlinks>"
                           "<hyperlink xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                           " ref=\"A1\" r:id=\"rId7\" tooltip=\"open site\" display=\"Click\"/>"
                           "</hyperlinks>");
  auto out = read_hyperlinks(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 1U);
  const Hyperlink& h = out.value()[0];
  EXPECT_EQ(h.row, 0U);
  EXPECT_EQ(h.col, 0U);
  EXPECT_EQ(h.rid, "rId7");
  EXPECT_EQ(h.tooltip, "open site");
  EXPECT_EQ(h.display, "Click");
  EXPECT_TRUE(h.target.empty());
  EXPECT_TRUE(h.location.empty());
}

TEST(HyperlinkRoundTrip, InternalLinkUsesLocation) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc, "<hyperlinks><hyperlink ref=\"B5\" location=\"Sheet2!A1\"/></hyperlinks>");
  auto out = read_hyperlinks(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 1U);
  EXPECT_EQ(out.value()[0].location, "Sheet2!A1");
  EXPECT_TRUE(out.value()[0].rid.empty());
}

TEST(HyperlinkRoundTrip, ApplyRelsFillsTarget) {
  std::vector<Hyperlink> hls;
  Hyperlink h;
  h.row = 0;
  h.col = 0;
  h.rid = "rId3";
  hls.push_back(std::move(h));
  std::unordered_map<std::string, std::string> rid_to_target = {
      {"rId3", "https://example.com"},
      {"rId4", "https://other.example"},
  };
  apply_hyperlink_rels(hls, rid_to_target);
  EXPECT_EQ(hls[0].target, "https://example.com");
}

TEST(HyperlinkRoundTrip, ApplyRelsLeavesUnknownAlone) {
  std::vector<Hyperlink> hls(1);
  hls[0].rid = "rIdMissing";
  hls[0].target = "preserved";
  std::unordered_map<std::string, std::string> rid_to_target;
  apply_hyperlink_rels(hls, rid_to_target);
  EXPECT_EQ(hls[0].target, "preserved");
}

TEST(HyperlinkRoundTrip, RangeRefAccepted) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc, "<hyperlinks><hyperlink ref=\"A1:B2\" location=\"Sheet2!A1\"/></hyperlinks>");
  auto out = read_hyperlinks(ws);
  ASSERT_TRUE(static_cast<bool>(out)) << out.error().message;
  ASSERT_EQ(out.value().size(), 1U);
  const Hyperlink& h = out.value().front();
  // Both corners are decoded into the numeric model rectangle.
  EXPECT_EQ(h.row, 0U);
  EXPECT_EQ(h.col, 0U);
  EXPECT_EQ(h.last_row, 1U);
  EXPECT_EQ(h.last_col, 1U);
  EXPECT_EQ(h.location, "Sheet2!A1");
}

// Same invalid-reference policy the sibling overlay readers apply: drop
// the entry, keep the sheet.
TEST(HyperlinkRoundTrip, MalformedRefSkipsOnlyThatEntry) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                           "<hyperlinks>"
                           "<hyperlink ref=\"\" location=\"Sheet2!A1\"/>"
                           "<hyperlink ref=\"!!\" location=\"Sheet2!A2\"/>"
                           "<hyperlink ref=\"C3\" location=\"Sheet2!A3\"/>"
                           "</hyperlinks>");
  auto out = read_hyperlinks(ws);
  ASSERT_TRUE(static_cast<bool>(out)) << out.error().message;
  ASSERT_EQ(out.value().size(), 1U);
  EXPECT_EQ(out.value()[0].row, 2U);
  EXPECT_EQ(out.value()[0].col, 2U);
  EXPECT_EQ(out.value()[0].location, "Sheet2!A3");
}

// Identity of a link is its exact ref, but a file may hold the same ref twice
// and overlapping refs: the reader keeps every entry in file order.
TEST(HyperlinkRoundTrip, DuplicateAndOverlappingRefsKeepFileOrder) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                           "<hyperlinks>"
                           "<hyperlink ref=\"B2\" location=\"Sheet2!A1\"/>"
                           "<hyperlink ref=\"A1:C3\" location=\"Sheet2!A2\"/>"
                           "<hyperlink ref=\"B2\" location=\"Sheet2!A3\"/>"
                           "<hyperlink ref=\"B2:D4\" location=\"Sheet2!A4\"/>"
                           "</hyperlinks>");
  auto out = read_hyperlinks(ws);
  ASSERT_TRUE(static_cast<bool>(out)) << out.error().message;
  ASSERT_EQ(out.value().size(), 4U);
  const std::vector<std::string> locations{"Sheet2!A1", "Sheet2!A2", "Sheet2!A3", "Sheet2!A4"};
  for (std::size_t i = 0; i < locations.size(); ++i) {
    EXPECT_EQ(out.value()[i].location, locations[i]) << i;
  }
  EXPECT_TRUE(out.value()[0].same_range(out.value()[2]));
  EXPECT_FALSE(out.value()[0].same_range(out.value()[1]));
  EXPECT_EQ(out.value()[1].last_row, 2U);
  EXPECT_EQ(out.value()[3].row, 1U);
}

TEST(HyperlinkRoundTrip, DuplicateAndOverlappingRefsSurviveSaveInListOrder) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  const auto make = [](std::uint32_t row, std::uint32_t col, std::uint32_t last_row, std::uint32_t last_col,
                       const char* target, const char* display) {
    Hyperlink link;
    link.row = row;
    link.col = col;
    link.last_row = last_row;
    link.last_col = last_col;
    link.target = target;
    link.display = display;
    return link;
  };
  sheet.mutable_hyperlinks().push_back(make(1, 1, 1, 1, "https://one.example/", "one"));
  sheet.mutable_hyperlinks().push_back(make(0, 0, 2, 2, "https://two.example/", "two"));
  sheet.mutable_hyperlinks().push_back(make(1, 1, 1, 1, "https://three.example/", "three"));
  sheet.mutable_hyperlinks().push_back(make(1, 1, 3, 3, "https://four.example/", "four"));

  auto saved = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto loaded = read_ooxml(ByteSpan{saved.value().data(), saved.value().size()});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const auto& links = loaded.value().workbook.sheet(0).hyperlinks();
  ASSERT_EQ(links.size(), 4U);
  const std::vector<std::string> targets{"https://one.example/", "https://two.example/", "https://three.example/",
                                         "https://four.example/"};
  const std::vector<std::string> displays{"one", "two", "three", "four"};
  for (std::size_t i = 0; i < links.size(); ++i) {
    EXPECT_EQ(links[i].target, targets[i]) << i;
    EXPECT_EQ(links[i].display, displays[i]) << i;
  }
  EXPECT_TRUE(links[0].same_range(links[2]));
  EXPECT_EQ(links[1].last_col, 2U);
  EXPECT_EQ(links[3].last_row, 3U);
  // The first entry of the pair is the one a read of the shared cell returns.
  EXPECT_EQ(loaded.value().workbook.sheet(0).hyperlink_at(1, 1)->target, "https://one.example/");
}

TEST(HyperlinkRoundTrip, SavedExternalTargetEncodesSpacesWhileTheModelKeepsThem) {
  Workbook wb = Workbook::create();
  Hyperlink link;
  link.target = "https://e.x/a b";
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, link)));
  EXPECT_EQ(wb.sheet(0).hyperlinks()[0].target, "https://e.x/a b");

  auto saved = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto loaded = read_ooxml(ByteSpan{saved.value().data(), saved.value().size()});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ASSERT_EQ(loaded.value().workbook.sheet(0).hyperlinks().size(), 1U);
  EXPECT_EQ(loaded.value().workbook.sheet(0).hyperlinks()[0].target, "https://e.x/a%20b");

  // A target that already holds %20 is written as it is.
  auto resaved = write_ooxml(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(resaved)) << resaved.error().message;
  auto reloaded = read_ooxml(ByteSpan{resaved.value().data(), resaved.value().size()});
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  EXPECT_EQ(reloaded.value().workbook.sheet(0).hyperlinks()[0].target, "https://e.x/a%20b");
}

}  // namespace
}  // namespace io
}  // namespace formulon
