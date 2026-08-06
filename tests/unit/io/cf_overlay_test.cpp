//
// Unit tests for the x14 conditional-formatting overlay reconciliation
// (`io/cf_overlay.h`). Each test feeds a raw `<extLst>` capture plus a
// CF model into `reconcile_x14_cf_overlay` and asserts the pruned
// overlay: removed model rules must lose their `<x14:cfRule id>` entry,
// surviving rules must keep theirs, and unrecoverable overlay shapes
// must fold to a full drop rather than leave a dangling GUID.

#include "io/cf_overlay.h"

#include <initializer_list>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"

namespace formulon::io {
namespace {

constexpr const char* kIdA = "{11111111-1111-1111-1111-111111111111}";
constexpr const char* kIdB = "{22222222-2222-2222-2222-222222222222}";

/// One `<ext>` block carrying two dataBar `<x14:cfRule>` entries (ids A
/// and B) in a single `<x14:conditionalFormatting>`, in the shape Excel
/// emits (uri + x14/xm namespace declarations, xm:sqref trailer).
std::string OverlayWithRules(std::initializer_list<const char*> ids) {
  std::string rules;
  for (const char* id : ids) {
    rules.append("<x14:cfRule type=\"dataBar\" id=\"");
    rules.append(id);
    rules.append(
        "\"><x14:dataBar minLength=\"0\" maxLength=\"100\"><x14:cfvo type=\"autoMin\"/>"
        "<x14:cfvo type=\"autoMax\"/><x14:negativeFillColor rgb=\"FFFF0000\"/></x14:dataBar></x14:cfRule>");
  }
  return "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
         "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
         "<x14:conditionalFormattings>"
         "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
         rules + "<xm:sqref>A1:A10</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";
}

/// Builds a single-block CF model holding one dataBar rule per id.
std::vector<cf::ConditionalFormat> ModelWithIds(std::initializer_list<const char*> ids) {
  std::vector<cf::ConditionalFormat> out;
  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  for (const char* id : ids) {
    cf::CFRule rule;
    rule.type = cf::RuleType::DataBar;
    rule.id = id;
    rule.data_bar = cf::DataBarSpec{};
    block.rules.push_back(std::move(rule));
  }
  out.push_back(std::move(block));
  return out;
}

bool Contains(const std::string& haystack, const std::string& needle) {
  return haystack.find(needle) != std::string::npos;
}

TEST(CfOverlay, EmptyOverlayStaysEmpty) {
  EXPECT_TRUE(reconcile_x14_cf_overlay(std::string(), ModelWithIds({kIdA})).empty());
}

TEST(CfOverlay, RemovedRuleIsStrippedSurvivorKept) {
  const std::string overlay = OverlayWithRules({kIdA, kIdB});
  const std::string out = reconcile_x14_cf_overlay(overlay, ModelWithIds({kIdB}));
  EXPECT_FALSE(Contains(out, kIdA));
  EXPECT_TRUE(Contains(out, kIdB));
  // The surviving rule keeps its container chain and payload.
  EXPECT_TRUE(Contains(out, "<extLst>"));
  EXPECT_TRUE(Contains(out, "x14:conditionalFormattings"));
  EXPECT_TRUE(Contains(out, "x14:negativeFillColor"));
}

TEST(CfOverlay, UntouchedOverlayRoundTripsByteForByte) {
  const std::string overlay = OverlayWithRules({kIdA, kIdB});
  EXPECT_EQ(reconcile_x14_cf_overlay(overlay, ModelWithIds({kIdA, kIdB})), overlay);
}

TEST(CfOverlay, ClearAllEmptiesOverlay) {
  const std::string overlay = OverlayWithRules({kIdA, kIdB});
  EXPECT_TRUE(reconcile_x14_cf_overlay(overlay, {}).empty());
}

TEST(CfOverlay, MalformedOverlayFallsBackToFullDrop) {
  // Unbalanced element: pugixml rejects it, so the referenced ids are
  // unknowable and the conservative path must drop everything.
  const std::string malformed = "<extLst><ext><x14:conditionalFormattings>";
  EXPECT_TRUE(reconcile_x14_cf_overlay(malformed, ModelWithIds({kIdA})).empty());
}

TEST(CfOverlay, UnexpectedRootFallsBackToFullDrop) {
  EXPECT_TRUE(reconcile_x14_cf_overlay("<notExtLst/>", ModelWithIds({kIdA})).empty());
}

TEST(CfOverlay, ForeignExtensionSurvivesCfPruning) {
  const std::string overlay =
      "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:conditionalFormattings><x14:conditionalFormatting>"
      "<x14:cfRule type=\"dataBar\" id=\"" +
      std::string(kIdA) +
      "\"/><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>"
      "<ext uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:dataValidations count=\"1\"/></ext></extLst>";
  const std::string out = reconcile_x14_cf_overlay(overlay, {});
  EXPECT_FALSE(Contains(out, kIdA));
  EXPECT_FALSE(Contains(out, "x14:conditionalFormattings"));
  EXPECT_TRUE(Contains(out, "x14:dataValidations"));
}

TEST(CfOverlay, RuleWithoutIdIsKept) {
  const std::string overlay =
      "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\">"
      "<x14:conditionalFormattings><x14:conditionalFormatting>"
      "<x14:cfRule type=\"dataBar\" id=\"" +
      std::string(kIdA) +
      "\"/><x14:cfRule type=\"iconSet\"/></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";
  const std::string out = reconcile_x14_cf_overlay(overlay, {});
  EXPECT_FALSE(Contains(out, kIdA));
  EXPECT_TRUE(Contains(out, "<x14:cfRule type=\"iconSet\"/>"));
}

}  // namespace
}  // namespace formulon::io
