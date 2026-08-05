//
// Reader-side tests for `<hyperlinks>` and the rid-to-target join. The
// writer side is exercised by `tests/integration/sheet_features_roundtrip_test.cpp`.

#include <string>
#include <unordered_map>
#include <vector>

#include "gtest/gtest.h"
#include "io/sheet_reader.h"
#include "pugixml.hpp"
#include "sheet.h"

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
  // Anchor is the range's top-left; the full span is preserved verbatim.
  EXPECT_EQ(h.row, 0U);
  EXPECT_EQ(h.col, 0U);
  EXPECT_EQ(h.ref_span, "A1:B2");
  EXPECT_EQ(h.location, "Sheet2!A1");
}

}  // namespace
}  // namespace io
}  // namespace formulon
