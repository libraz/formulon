//
// Unit tests for the x14 conditional-formatting overlay reconciliation
// (`io/cf_overlay.h`). Each test feeds a raw `<extLst>` capture plus a
// CF model into `reconcile_x14_cf_overlay` and asserts the pruned
// overlay: removed model rules must lose their `<x14:cfRule id>` entry,
// surviving rules must keep theirs, and unrecoverable overlay shapes
// must fold to a full drop rather than leave a dangling GUID.

#include "io/cf_overlay.h"

#include <cstddef>
#include <initializer_list>
#include <string>
#include <utility>
#include <vector>

#include "../small_stack.h"
#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/cf_writer.h"

namespace formulon::io {
namespace {

constexpr const char* kIdA = "{11111111-1111-1111-1111-111111111111}";
constexpr const char* kIdB = "{22222222-2222-2222-2222-222222222222}";

/// One `<ext>` block carrying two dataBar `<x14:cfRule>` entries (ids A
/// and B) in a single `<x14:conditionalFormatting>`, in the shape this
/// engine's writer emits (uri + x14/xm namespace declarations, xm:sqref
/// trailer). No Excel-authored capture of an x14 overlay exists in this
/// repository, so what Excel writes here is unobserved; these tests
/// exercise reconciliation against our own emission.
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

std::string DeepXml(std::size_t depth, const std::string& body) {
  std::string xml;
  xml.reserve(depth * 20U + body.size());
  for (std::size_t i = 0; i < depth; ++i) {
    xml.append("<x:deep>");
  }
  xml.append(body);
  for (std::size_t i = 0; i < depth; ++i) {
    xml.append("</x:deep>");
  }
  return xml;
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

TEST(CfOverlay, RulesTheModelCannotHoldSurviveAMutationOfOthers) {
  // A 3Stars icon set and a formula rule live only in the extension and
  // never enter the model, so their absence from it is no removal.
  constexpr const char* kStars = "{33333333-3333-3333-3333-333333333333}";
  constexpr const char* kFormula = "{44444444-4444-4444-4444-444444444444}";
  std::string overlay = OverlayWithRules({kIdA, kIdB});
  const std::string extra = std::string("<x14:cfRule type=\"iconSet\" priority=\"3\" id=\"") + kStars +
                            "\"><x14:iconSet iconSet=\"3Stars\"><x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo>"
                            "</x14:iconSet></x14:cfRule><x14:cfRule type=\"expression\" priority=\"4\" id=\"" +
                            kFormula + "\"><xm:f>Sheet2!$A$1&gt;0</xm:f></x14:cfRule>";
  overlay.insert(overlay.find("<xm:sqref>"), extra);
  const std::string out = reconcile_x14_cf_overlay(overlay, ModelWithIds({kIdB}));
  EXPECT_FALSE(Contains(out, kIdA)) << out;
  EXPECT_TRUE(Contains(out, kIdB)) << out;
  EXPECT_TRUE(Contains(out, kStars)) << out;
  EXPECT_TRUE(Contains(out, kFormula)) << out;
  // With every modelled rule cleared they still remain.
  const std::string cleared = reconcile_x14_cf_overlay(overlay, {});
  EXPECT_TRUE(Contains(cleared, kStars)) << cleared;
  EXPECT_TRUE(Contains(cleared, kFormula)) << cleared;
}

TEST(CfOverlay, UntouchedOverlayRoundTripsByteForByte) {
  const std::string overlay = OverlayWithRules({kIdA, kIdB});
  EXPECT_EQ(reconcile_x14_cf_overlay(overlay, ModelWithIds({kIdA, kIdB})), overlay);
}

TEST(CfOverlay, DeeplyClaimedRuleDoesNotRecurseOrDuplicate) {
  const std::string overlay =
      "<extLst><ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" +
      DeepXml(10000U, "<x14:cfRule id=\"" + std::string(kIdA) + "\"/>") + "</ext></extLst>";
  std::string merged;
  testing_support::RunOnSmallStack([&] { merged = merge_x14_cf_entries(overlay, ModelWithIds({kIdA})); });
  EXPECT_EQ(merged, overlay);
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

constexpr const char* kExcelId1 = "{00000000-000E-0000-0100-000001000000}";

/// Excel's Sheet2 extension from structural_ref_m10b: one cross-sheet
/// formula rule and a data-validation extension after it.
std::string ExcelCrossSheetOverlay() {
  return std::string(
             "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
             "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:conditionalFormattings>"
             "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
             "<x14:cfRule type=\"expression\" priority=\"1\" id=\"") +
         kExcelId1 +
         "\"><xm:f>Sheet1!$A$1&gt;0</xm:f><x14:dxf/></x14:cfRule><xm:sqref>C1:C3</xm:sqref>"
         "</x14:conditionalFormatting></x14:conditionalFormattings></ext>"
         "<ext uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\" "
         "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
         "<x14:dataValidations count=\"1\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/>"
         "</ext></extLst>";
}

TEST(CfOverlay, StripRemovesAdoptedRulesAndTheirEmptyContainers) {
  const std::string out = strip_x14_cf_rules(ExcelCrossSheetOverlay(), {kExcelId1});
  EXPECT_FALSE(Contains(out, kExcelId1)) << out;
  EXPECT_FALSE(Contains(out, "x14:conditionalFormattings")) << out;
  EXPECT_FALSE(Contains(out, "{78C0D931-6437-407d-A8EE-F0AAD7539E65}")) << out;
  EXPECT_TRUE(Contains(out, "x14:dataValidations")) << out;
}

TEST(CfOverlay, StripLeavesAnOverlayWithoutTheIdsByteIdentical) {
  const std::string overlay = ExcelCrossSheetOverlay();
  EXPECT_EQ(strip_x14_cf_rules(overlay, {kIdA}), overlay);
  EXPECT_EQ(strip_x14_cf_rules(overlay, {}), overlay);
}

TEST(CfOverlay, StripOfEveryRuleEmptiesACfOnlyOverlay) {
  EXPECT_TRUE(strip_x14_cf_rules(OverlayWithRules({kIdA, kIdB}), {kIdA, kIdB}).empty());
}

TEST(CfOverlay, CrossSheetRuleJoinsTheOverlayAheadOfOtherExtensions) {
  // The captured CF extension was stripped on load; the rebuilt one goes
  // first, where Excel writes it.
  const std::string captured = strip_x14_cf_rules(ExcelCrossSheetOverlay(), {kExcelId1});
  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 2}, {2, 2}});
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = 1;
  rule.formula1 = "Sheet1!$A$1>0";
  rule.id = kExcelId1;
  block.rules.push_back(rule);
  const std::string merged = merge_x14_cf_entries(captured, {block}, nullptr, CfSheetContext{"Sheet2", 1, nullptr});
  const std::size_t cf_ext = merged.find("{78C0D931-6437-407d-A8EE-F0AAD7539E65}");
  const std::size_t dv_ext = merged.find("{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}");
  ASSERT_NE(cf_ext, std::string::npos) << merged;
  ASSERT_NE(dv_ext, std::string::npos) << merged;
  EXPECT_LT(cf_ext, dv_ext) << merged;
  EXPECT_TRUE(Contains(merged, "<x14:cfRule type=\"expression\" priority=\"1\" id=\"" + std::string(kExcelId1) +
                                   "\"><xm:f>Sheet1!$A$1&gt;0</xm:f><x14:dxf/></x14:cfRule>"))
      << merged;
}

TEST(CfOverlay, MintedIdAvoidsAnIdTheOverlayAlreadyHolds) {
  // A rule left in the extension holds Excel's id for priority 1; a new
  // cross-sheet rule at priority 1 must not take it.
  const std::string captured = ExcelCrossSheetOverlay();
  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 3}, {2, 3}});
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = 1;
  rule.formula1 = "Sheet3!$A$1>0";
  block.rules.push_back(rule);
  const std::string merged = merge_x14_cf_entries(captured, {block}, nullptr, CfSheetContext{"Sheet2", 1, nullptr});
  EXPECT_TRUE(Contains(merged, "id=\"{00000000-000E-0000-0100-000002000000}\"><xm:f>Sheet3!$A$1&gt;0</xm:f>"))
      << merged;
}

}  // namespace
}  // namespace formulon::io
