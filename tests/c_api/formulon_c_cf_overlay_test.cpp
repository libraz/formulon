// Stable C ABI conditional-format extended-overlay tests.

#include <initializer_list>
#include <utility>

#include "formulon_c_cf_test_helpers.h"

namespace {

constexpr const char* kX14IdA = "{11111111-1111-1111-1111-111111111111}";
constexpr const char* kX14IdB = "{22222222-2222-2222-2222-222222222222}";

// Builds an Excel-shaped worksheet-level x14 overlay carrying one
// dataBar `<x14:cfRule>` per id.
std::string X14OverlayFor(std::initializer_list<const char*> ids) {
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

// Builds the nested `<extLst>` link a legacy `<cfRule>` carries to reach
// its x14 counterpart. This, not an `id` attribute on the `<cfRule>`, is
// how Excel spells the cross-reference, so a rule loaded from an
// x14-bearing file always has one.
std::string X14RuleLinkFor(const char* id) {
  return std::string(
             "<extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" "
             "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>") +
         id + "</x14:id></ext></extLst>";
}

// Seeds sheet 0 with one CF block holding two id-bearing dataBar rules
// plus the matching two-entry x14 overlay, mirroring the state produced
// by loading an Excel 2010+ file with extended data bars.
void SeedDataBarRulesWithOverlay(fm_workbook_t* handle) {
  auto& sheet = handle->workbook().sheet(0);
  formulon::cf::ConditionalFormat block{};
  block.sqref.push_back(MakeRange(0, 0, 9, 0));
  for (const char* id : {kX14IdA, kX14IdB}) {
    formulon::cf::CFRule rule;
    rule.type = formulon::cf::RuleType::DataBar;
    rule.id = id;
    rule.ext_lst_raw = X14RuleLinkFor(id);
    rule.data_bar = formulon::cf::DataBarSpec{};
    block.rules.push_back(std::move(rule));
  }
  sheet.mutable_conditional_formats().push_back(std::move(block));
  sheet.set_ext_lst_xml(X14OverlayFor({kX14IdA, kX14IdB}));
}

bool ContainsSubstring(const std::string& haystack, const std::string& needle) {
  return haystack.find(needle) != std::string::npos;
}

}  // namespace

TEST(FormulonCApiCfMutate, RemoveAtPrunesRemovedRuleFromX14Overlay) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDataBarRulesWithOverlay(wb.handle);

  ASSERT_EQ(fm_sheet_cf_remove_at(wb.handle, 0, 0), 0);

  const std::string& overlay = wb.handle->workbook().sheet(0).ext_lst_xml();
  EXPECT_FALSE(ContainsSubstring(overlay, kX14IdA));
  EXPECT_TRUE(ContainsSubstring(overlay, kX14IdB));
}

TEST(FormulonCApiCfMutate, ClearEmptiesX14Overlay) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDataBarRulesWithOverlay(wb.handle);

  ASSERT_EQ(fm_sheet_cf_clear(wb.handle, 0), 0);
  EXPECT_TRUE(wb.handle->workbook().sheet(0).ext_lst_xml().empty());
}

TEST(FormulonCApiCfMutate, MalformedOverlayDroppedWhollyOnRemove) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDataBarRulesWithOverlay(wb.handle);
  wb.handle->workbook().sheet(0).set_ext_lst_xml("<extLst><ext><x14:conditionalFormattings>");

  ASSERT_EQ(fm_sheet_cf_remove_at(wb.handle, 0, 0), 0);
  EXPECT_TRUE(wb.handle->workbook().sheet(0).ext_lst_xml().empty());
}

TEST(FormulonCApiCfMutate, RemovedRuleDoesNotResurfaceThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDataBarRulesWithOverlay(wb.handle);

  ASSERT_EQ(fm_sheet_cf_remove_at(wb.handle, 0, 0), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  const auto& sheet = reloaded.handle->workbook().sheet(0);
  ASSERT_EQ(sheet.conditional_formats().size(), 1U);
  ASSERT_EQ(sheet.conditional_formats()[0].rules.size(), 1U);
  EXPECT_EQ(sheet.conditional_formats()[0].rules[0].id, kX14IdB);
  EXPECT_FALSE(ContainsSubstring(sheet.ext_lst_xml(), kX14IdA));
  EXPECT_TRUE(ContainsSubstring(sheet.ext_lst_xml(), kX14IdB));
}

TEST(FormulonCApiCfMutate, DataBarExtensionFieldsSurviveSaveAndLoad) {
  // Reaching the model is not enough: axis, gradient and the negative
  // colours have no legacy `<dataBar>` attribute, so a save that does not
  // build the x14 extension loses every one of them and the caller gets a
  // plain bar back. This is the only test on this path that crosses the
  // ABI boundary in both directions.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule = MakeDataBarRule(&sqref);
  rule.data_bar_gradient_engaged = 1;
  rule.data_bar_gradient = 0;
  rule.data_bar_axis_position_engaged = 1;
  rule.data_bar_axis_position = 1;  // middle
  rule.data_bar_negative_fill_engaged = 1;
  rule.data_bar_negative_fill = fm_cf_color_t{255, 0, 0, 255};
  rule.data_bar_border_engaged = 1;
  rule.data_bar_border = fm_cf_color_t{1, 2, 3, 255};
  rule.data_bar_negative_border_engaged = 1;
  rule.data_bar_negative_border = fm_cf_color_t{4, 5, 6, 255};
  rule.data_bar_axis_color_engaged = 1;
  rule.data_bar_axis_color = fm_cf_color_t{7, 8, 9, 255};
  rule.data_bar_direction = 2;  // right to left
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);

  fm_cf_rule_t back{};
  ASSERT_EQ(fm_sheet_cf_get_at(reloaded.handle, 0, 0, &back), 0) << fm_last_error_message();
  EXPECT_EQ(back.data_bar_gradient, 0);
  EXPECT_EQ(back.data_bar_axis_position, 1U);
  EXPECT_EQ(back.data_bar_negative_fill.r, 255U);
  EXPECT_EQ(back.data_bar_negative_fill.g, 0U);
  EXPECT_EQ(back.data_bar_border_engaged, 1);
  EXPECT_EQ(back.data_bar_border.b, 3U);
  EXPECT_EQ(back.data_bar_negative_border_engaged, 1);
  EXPECT_EQ(back.data_bar_negative_border.b, 6U);
  EXPECT_EQ(back.data_bar_axis_color.r, 7U);
  EXPECT_EQ(back.data_bar_axis_color.b, 9U);
  EXPECT_EQ(back.data_bar_direction, 2U);
}

TEST(FormulonCApiCfMutate, IconSetFloorSurvivesSaveAndLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 4;  // IconSet
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  rule.icon_set_engaged = 1;
  rule.icon_set_name = 2;  // Three_Flags
  fm_cfvo_t thresholds[2]{};
  thresholds[0].type = 0;  // Number
  thresholds[0].value = "7";
  thresholds[0].gte = 1;
  thresholds[1].type = 0;
  thresholds[1].value = "9";
  thresholds[1].gte = 1;
  rule.icon_set_thresholds = thresholds;
  rule.icon_set_threshold_count = 2;
  rule.icon_set_show_value = 1;
  rule.icon_set_percent = 1;
  rule.icon_set_floor_engaged = 1;
  rule.icon_set_floor.type = 0;  // Number
  rule.icon_set_floor.value = "5";
  rule.icon_set_floor.gte = 0;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  fm_cf_rule_t back{};
  ASSERT_EQ(fm_sheet_cf_get_at(reloaded.handle, 0, 0, &back), 0) << fm_last_error_message();
  EXPECT_EQ(back.icon_set_floor_engaged, 1);
  EXPECT_EQ(back.icon_set_floor.type, 0U);
  EXPECT_STREQ(back.icon_set_floor.value, "5");
  EXPECT_EQ(back.icon_set_floor.gte, 0);
}

TEST(FormulonCApiCfMutate, UnengagedIconSetFloorIsExcelsDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 4;  // IconSet
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  rule.icon_set_engaged = 1;
  fm_cfvo_t thresholds[2]{};
  rule.icon_set_thresholds = thresholds;
  rule.icon_set_threshold_count = 2;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();
  const auto& floor = wb.handle->workbook().sheet(0).conditional_formats()[0].rules[0].icon_set->floor;
  EXPECT_EQ(floor.type, formulon::cf::CfvoType::Percent);
  EXPECT_EQ(floor.value, "0");
  EXPECT_TRUE(floor.gte);
}
