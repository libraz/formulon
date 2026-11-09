// Stable C ABI conditional-format authoring tests.

#include <iterator>
#include <limits>

#include "formulon_c_cf_test_helpers.h"

TEST(FormulonCApiCfMutate, AddCellIsRuleRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  SeedDxfs(wb.handle, 1);

  fm_cell_rect_t sqref{};
  sqref.first_row = 0;
  sqref.first_col = 0;
  sqref.last_row = 9;
  sqref.last_col = 0;

  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 1;  // CellIs
  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_GREATER_THAN;
  std::string formula = "50";
  rule.formula1 = formula.c_str();
  rule.has_dxf_index = 1;
  rule.dxf_index = 0;
  rule.ranges = &sqref;
  rule.range_count = 1;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0);
  EXPECT_EQ(rule_index, 0U);

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.type, 1U);
  EXPECT_EQ(out.has_op, 1);
  EXPECT_EQ(out.op, FM_COMPARISON_OPERATOR_GREATER_THAN);
  EXPECT_EQ(out.priority, 1);  // auto-assigned since input <= 0
  EXPECT_EQ(out.has_dxf_index, 1);
  EXPECT_EQ(out.dxf_index, 0U);
  ASSERT_NE(out.formula1, nullptr);
  EXPECT_STREQ(out.formula1, "=50");
  ASSERT_NE(out.ranges, nullptr);
  EXPECT_EQ(out.range_count, 1U);
  EXPECT_EQ(out.ranges[0].first_row, 0U);
  EXPECT_EQ(out.ranges[0].last_row, 9U);
  ASSERT_NE(out.id, nullptr);
  EXPECT_FALSE(std::string(out.id).empty());
}

// A `cellIs` rule takes the shared comparison table; each member reaches the
// model as the operator it names, and reads back as the same member.
TEST(FormulonCApiCfMutate, CellIsRuleTakesEachComparisonOperatorAsNamed) {
  using formulon::cf::CellIsOperator;
  const struct {
    int32_t op;
    CellIsOperator model;
  } kCases[] = {
      {FM_COMPARISON_OPERATOR_BETWEEN, CellIsOperator::Between},
      {FM_COMPARISON_OPERATOR_NOT_BETWEEN, CellIsOperator::NotBetween},
      {FM_COMPARISON_OPERATOR_EQUAL, CellIsOperator::Equal},
      {FM_COMPARISON_OPERATOR_NOT_EQUAL, CellIsOperator::NotEqual},
      {FM_COMPARISON_OPERATOR_LESS_THAN, CellIsOperator::LessThan},
      {FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL, CellIsOperator::LessThanOrEqual},
      {FM_COMPARISON_OPERATOR_GREATER_THAN, CellIsOperator::GreaterThan},
      {FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL, CellIsOperator::GreaterThanOrEqual},
  };
  for (const auto& c : kCases) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    fm_cell_rect_t sqref{0, 0, 9, 0};
    fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
    rule.type = 1;  // CellIs
    rule.has_op = 1;
    rule.op = static_cast<std::uint8_t>(c.op);
    rule.formula1 = "1";
    rule.formula2 = "10";
    rule.ranges = &sqref;
    rule.range_count = 1;
    std::size_t index = 0;
    ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
    const auto& blocks = wb.handle->workbook().sheet(0).conditional_formats();
    ASSERT_EQ(blocks.size(), 1U);
    ASSERT_TRUE(blocks[0].rules[0].op.has_value()) << c.op;
    EXPECT_EQ(blocks[0].rules[0].op.value(), c.model) << c.op;
    fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
    ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
    EXPECT_EQ(out.op, c.op);
  }
}

TEST(FormulonCApiCfMutate, AddRuleRejectsDxfIdPastTheStylesTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  SeedDxfs(wb.handle, 2);

  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 1;  // CellIs
  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_GREATER_THAN;
  rule.formula1 = "50";
  rule.has_dxf_index = 1;
  rule.dxf_index = 2;  // one past the last registered `<dxf>`
  rule.ranges = &sqref;
  rule.range_count = 1;

  std::size_t rule_index = 99U;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_STREQ(fm_get_last_error_message(), "dxf_index out of range");

  // The rejected rule is not appended, so a workbook saved after the failed
  // call carries no `<cfRule>` whose `dxfId` Excel cannot resolve.
  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  // The last in-range index is still accepted.
  rule.dxf_index = 1;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiCfMutate, AddRuleWithoutDxfIdIgnoresTheEmptyStylesTable) {
  // A rule that engages no differential format must stay addable on a
  // workbook that has registered no `<dxf>` at all.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 1;  // CellIs
  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_GREATER_THAN;
  rule.formula1 = "50";
  rule.ranges = &sqref;
  rule.range_count = 1;

  std::size_t rule_index = 99U;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(rule_index, 0U);

  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.has_dxf_index, 0);
}

TEST(FormulonCApiCfMutate, AddMultipleRulesAutoIncrementsPriority) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_cell_rect_t sqref{0, 0, 0, 0};
  for (int i = 0; i < 3; ++i) {
    fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
    rule.type = 0;  // Expression
    std::string f = "TRUE";
    rule.formula1 = f.c_str();
    rule.ranges = &sqref;
    rule.range_count = 1;
    std::size_t rule_index = 0;
    ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0);
    EXPECT_EQ(rule_index, static_cast<std::size_t>(i));
  }

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 3U);
  for (std::size_t i = 0; i < 3; ++i) {
    fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
    ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, i, &out), 0);
    EXPECT_EQ(out.priority, static_cast<int32_t>(i + 1));
  }
}

TEST(FormulonCApiCfMutate, AutoPriorityRejectsExhaustedInt32RangeAtomically) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_cell_rect_t sqref{0, 0, 0, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;  // Expression
  rule.formula1 = "TRUE";
  rule.ranges = &sqref;
  rule.range_count = 1;
  rule.priority = std::numeric_limits<std::int32_t>::max();

  std::size_t rule_index = 0U;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(rule_index, 0U);

  std::size_t count = 0U;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  for (const std::int32_t auto_priority : {0, -1, std::numeric_limits<std::int32_t>::min()}) {
    rule.priority = auto_priority;
    rule_index = 99U;
    EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
    EXPECT_EQ(rule_index, 99U);
    ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
    EXPECT_EQ(count, 1U);
  }
}

TEST(FormulonCApiCfMutate, AutoPriorityMayReachMaxThenExplicitPriorityStillWorks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_cell_rect_t sqref{0, 0, 0, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;  // Expression
  rule.formula1 = "TRUE";
  rule.ranges = &sqref;
  rule.range_count = 1;
  rule.priority = std::numeric_limits<std::int32_t>::max() - 1;

  std::size_t rule_index = 0U;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(rule_index, 0U);

  rule.priority = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(rule_index, 1U);
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, rule_index, &out), 0);
  EXPECT_EQ(out.priority, std::numeric_limits<std::int32_t>::max());

  rule.priority = -1;
  rule_index = 99U;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(rule_index, 99U);

  std::size_t count = 0U;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  // Explicit positive priorities remain usable even when automatic assignment
  // has reached its limit.
  rule.priority = 7;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(rule_index, 2U);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, rule_index, &out), 0);
  EXPECT_EQ(out.priority, 7);
}

TEST(FormulonCApiCfMutate, RemoveAtFlattensIndices) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_cell_rect_t sqref{0, 0, 0, 0};
  std::vector<std::string> formulas{"A1", "A2", "A3"};
  std::size_t expected_index = 0;
  for (const auto& f : formulas) {
    fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
    rule.type = 0;  // Expression
    rule.formula1 = f.c_str();
    rule.ranges = &sqref;
    rule.range_count = 1;
    std::size_t rule_index = 0;
    ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0);
    EXPECT_EQ(rule_index, expected_index);
    ++expected_index;
  }
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 3U);

  ASSERT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, 1), 0);
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  // A record's storage belongs to the handle and the next get takes it
  // back, so each rule is read before the following one is fetched.
  fm_cf_rule_t out0 = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out0), 0);
  EXPECT_STREQ(out0.formula1, "=A1");
  fm_cf_rule_t out1 = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 1, &out1), 0);
  EXPECT_STREQ(out1.formula1, "=A3");
}

TEST(FormulonCApiCfMutate, ClearRemovesAllBlocks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 0, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;
  std::string f = "TRUE";
  rule.formula1 = f.c_str();
  rule.ranges = &sqref;
  rule.range_count = 1;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_index), 0);
  EXPECT_EQ(rule_index, 0U);

  ASSERT_EQ(fm_sheet_clear_cf_rules(wb.handle, 0), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiCfMutate, AddsVisualRuleTypesAndPreservesThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 0, 0};

  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 2;  // ColorScale
  rule.ranges = &sqref;
  rule.range_count = 1;
  fm_cfvo_t thresholds[3]{};
  thresholds[0].type = 3;  // Min
  thresholds[0].gte = 1;
  thresholds[1].type = 1;  // Percent
  thresholds[1].value = "50";
  thresholds[1].gte = 1;
  thresholds[2].type = 4;  // Max
  thresholds[2].gte = 1;
  fm_color_t colors[3]{CfRgb(255, 0, 0, 255), CfRgb(255, 255, 0, 255), CfRgb(0, 255, 0, 255)};
  rule.color_scale_thresholds = thresholds;
  rule.color_scale_colors = colors;
  rule.color_scale_count = 3;
  std::size_t color_rule_index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &color_rule_index), 0);
  EXPECT_EQ(color_rule_index, 0U);

  fm_cell_rect_t db_sqref{1, 0, 1, 0};
  fm_cf_rule_t db_rule = FM_STRUCT_INIT(fm_cf_rule_t);
  db_rule.type = 3;  // DataBar
  db_rule.ranges = &db_sqref;
  db_rule.range_count = 1;
  db_rule.has_data_bar = 1;
  db_rule.data_bar_min.type = 3;  // Min
  db_rule.data_bar_min.gte = 1;
  db_rule.data_bar_max.type = 4;  // Max
  db_rule.data_bar_max.gte = 1;
  db_rule.data_bar_fill = CfRgb(99, 142, 198, 255);
  db_rule.data_bar_show_value = 1;
  db_rule.data_bar_min_length_pct = 10;
  db_rule.data_bar_max_length_pct = 90;
  std::size_t data_bar_rule_index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &db_rule, &data_bar_rule_index), 0);
  EXPECT_EQ(data_bar_rule_index, 1U);

  fm_cell_rect_t icon_sqref{2, 0, 2, 0};
  fm_cf_rule_t icon_rule = FM_STRUCT_INIT(fm_cf_rule_t);
  icon_rule.type = 4;  // IconSet
  icon_rule.ranges = &icon_sqref;
  icon_rule.range_count = 1;
  icon_rule.has_icon_set = 1;
  icon_rule.icon_set_name = 0;  // Three_Arrows
  fm_cfvo_t icon_thresholds[2]{};
  icon_thresholds[0].type = 1;  // Percent
  icon_thresholds[0].value = "33";
  icon_thresholds[0].gte = 1;
  icon_thresholds[1].type = 1;  // Percent
  icon_thresholds[1].value = "67";
  icon_thresholds[1].gte = 1;
  icon_rule.icon_set_thresholds = icon_thresholds;
  icon_rule.icon_set_threshold_count = 2;
  icon_rule.icon_set_show_value = 1;
  icon_rule.icon_set_percent = 1;
  std::size_t icon_rule_index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &icon_rule, &icon_rule_index), 0);
  EXPECT_EQ(icon_rule_index, 2U);

  const auto& blocks = wb.handle->workbook().sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 3U);
  ASSERT_EQ(blocks[0].rules.size(), 1U);
  ASSERT_TRUE(blocks[0].rules[0].color_scale.has_value());
  ASSERT_EQ(blocks[0].rules[0].color_scale->thresholds.size(), 3U);
  EXPECT_EQ(blocks[0].rules[0].color_scale->thresholds[1].value, "50");
  EXPECT_EQ(blocks[0].rules[0].color_scale->colors[2].g, 255U);
  ASSERT_TRUE(blocks[1].rules[0].data_bar.has_value());
  EXPECT_EQ(blocks[1].rules[0].data_bar->fill.b, 198U);
  ASSERT_TRUE(blocks[2].rules[0].icon_set.has_value());
  EXPECT_EQ(blocks[2].rules[0].icon_set->thresholds[1].value, "67");

  fm_cf_rule_t out_color = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out_color), 0);
  ASSERT_EQ(out_color.color_scale_count, 3U);
  ASSERT_NE(out_color.color_scale_thresholds, nullptr);
  ASSERT_NE(out_color.color_scale_colors, nullptr);
  EXPECT_EQ(out_color.color_scale_thresholds[1].type, 1U);
  EXPECT_STREQ(out_color.color_scale_thresholds[1].value, "50");
  EXPECT_EQ(CfG(out_color.color_scale_colors[2]), 255U);

  fm_cf_rule_t out_bar = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 1, &out_bar), 0);
  EXPECT_EQ(out_bar.has_data_bar, 1);
  EXPECT_EQ(CfB(out_bar.data_bar_fill), 198U);
  EXPECT_EQ(out_bar.data_bar_min_length_pct, 10U);
  EXPECT_EQ(out_bar.data_bar_max_length_pct, 90U);

  fm_cf_rule_t out_icon = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 2, &out_icon), 0);
  EXPECT_EQ(out_icon.has_icon_set, 1);
  EXPECT_EQ(out_icon.icon_set_threshold_count, 2U);
  ASSERT_NE(out_icon.icon_set_thresholds, nullptr);
  EXPECT_STREQ(out_icon.icon_set_thresholds[1].value, "67");

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  const auto& reloaded_blocks = reloaded.handle->workbook().sheet(0).conditional_formats();
  ASSERT_EQ(reloaded_blocks.size(), 3U);
  ASSERT_EQ(reloaded_blocks[0].rules.size(), 1U);
  ASSERT_TRUE(reloaded_blocks[0].rules[0].color_scale.has_value());
  EXPECT_EQ(reloaded_blocks[0].rules[0].color_scale->colors[0].r, 255U);
  ASSERT_TRUE(reloaded_blocks[1].rules[0].data_bar.has_value());
  EXPECT_EQ(reloaded_blocks[1].rules[0].data_bar->fill.g, 142U);
  ASSERT_TRUE(reloaded_blocks[2].rules[0].icon_set.has_value());
  EXPECT_EQ(reloaded_blocks[2].rules[0].icon_set->thresholds.size(), 2U);
}

// One leading '=' is accepted and dropped, so the stored rule matches the
// one a reader builds from the saved file.
TEST(FormulonCApiCfMutate, AddRuleStoresFormulasWithoutTheLeadingEquals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 1;  // CellIs
  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_BETWEEN;
  rule.formula1 = "=$B1>100";
  rule.formula2 = "=SUM($C$1:$C$2)";
  rule.ranges = &sqref;
  rule.range_count = 1;
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();

  fm_cfvo_t thresholds[2]{};
  thresholds[0].type = 5;  // Formula
  thresholds[0].value = "=$D$1";
  thresholds[0].gte = 1;
  thresholds[1].type = 4;  // Max
  thresholds[1].gte = 1;
  fm_color_t colors[2]{CfRgb(255, 0, 0, 255), CfRgb(0, 255, 0, 255)};
  fm_cf_rule_t scale = FM_STRUCT_INIT(fm_cf_rule_t);
  scale.type = 2;  // ColorScale
  scale.ranges = &sqref;
  scale.range_count = 1;
  scale.color_scale_thresholds = thresholds;
  scale.color_scale_colors = colors;
  scale.color_scale_count = 2;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &scale, &index), 0) << fm_get_last_error_message();

  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.formula1, "=$B1>100");
  EXPECT_STREQ(out.formula2, "=SUM($C$1:$C$2)");
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 1, &out), 0);
  ASSERT_EQ(out.color_scale_count, 2U);
  EXPECT_STREQ(out.color_scale_thresholds[0].value, "=$D$1");

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(reloaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.formula1, "=$B1>100");
  EXPECT_STREQ(out.formula2, "=SUM($C$1:$C$2)");

  // Only one '=' is the entry marker; a formula that still starts with one
  // after it does not parse.
  rule.formula1 = "==$B1>100";
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken));
  thresholds[0].value = "=SUM(";
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &scale, &index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken));
}

// A formula threshold reads with a leading `=` and takes either spelling.
TEST(FormulonCApiCfMutate, FormulaThresholdsReadWithALeadingEquals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cfvo_t thresholds[2]{};
  thresholds[0].type = 5;  // Formula
  thresholds[0].value = "$D$1";
  thresholds[0].gte = 1;
  thresholds[1].type = 5;  // Formula
  thresholds[1].value = "=$D$2";
  thresholds[1].gte = 1;
  fm_color_t colors[2]{CfRgb(255, 0, 0, 255), CfRgb(0, 255, 0, 255)};
  fm_cf_rule_t scale = FM_STRUCT_INIT(fm_cf_rule_t);
  scale.type = 2;  // ColorScale
  scale.ranges = &sqref;
  scale.range_count = 1;
  scale.color_scale_thresholds = thresholds;
  scale.color_scale_colors = colors;
  scale.color_scale_count = 2;
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &scale, &index), 0) << fm_get_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  for (fm_workbook_t* book : {wb.handle, reloaded.handle}) {
    fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
    ASSERT_EQ(fm_sheet_get_cf_rule_at(book, 0, 0, &out), 0);
    ASSERT_EQ(out.color_scale_count, 2U);
    EXPECT_STREQ(out.color_scale_thresholds[0].value, "=$D$1");
    EXPECT_STREQ(out.color_scale_thresholds[1].value, "=$D$2");
  }
  // The model keeps the formula without the entry marker.
  const auto& blocks = wb.handle->workbook().sheet(0).conditional_formats();
  ASSERT_TRUE(blocks[0].rules[0].color_scale.has_value());
  EXPECT_EQ(blocks[0].rules[0].color_scale->thresholds[0].value, "$D$1");
  EXPECT_EQ(blocks[0].rules[0].color_scale->thresholds[1].value, "$D$2");
}

// An N-icon set carries exactly N-1 thresholds; the floor is separate.
TEST(FormulonCApiCfMutate, AddRuleRejectsIconThresholdCountsOtherThanIconsMinusOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 4, 0};
  fm_cfvo_t thresholds[4]{};
  for (fm_cfvo_t& t : thresholds) {
    t.type = 1;  // Percent
    t.value = "50";
    t.gte = 1;
  }
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 4;  // IconSet
  rule.ranges = &sqref;
  rule.range_count = 1;
  rule.has_icon_set = 1;
  rule.icon_set_name = 0;  // Three_Arrows
  rule.icon_set_thresholds = thresholds;
  std::size_t index = 0;
  for (std::uint32_t count : {0U, 1U, 3U, 4U}) {
    rule.icon_set_threshold_count = count;
    EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << "count=" << count;
  }
  rule.icon_set_threshold_count = 2;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  rule.icon_set_name = 8;  // Four_Arrows
  EXPECT_NE(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0);
  rule.icon_set_threshold_count = 3;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  rule.icon_set_name = 16;  // Five_Quarters
  rule.icon_set_threshold_count = 4;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 3U);
}

// The icon set table is in `ST_IconSetType` order: the 17 built-ins a rule
// draws, then the x14-only sets a custom icon level can name.
TEST(FormulonCApiCfMutate, IconSetNameOrdinalsFollowTheSchemaOrder) {
  EXPECT_EQ(FM_ICON_SET_NAME_3_ARROWS, 0);
  EXPECT_EQ(FM_ICON_SET_NAME_3_ARROWS_GRAY, 1);
  EXPECT_EQ(FM_ICON_SET_NAME_3_FLAGS, 2);
  EXPECT_EQ(FM_ICON_SET_NAME_3_TRAFFIC_LIGHTS_1, 3);
  EXPECT_EQ(FM_ICON_SET_NAME_3_TRAFFIC_LIGHTS_2, 4);
  EXPECT_EQ(FM_ICON_SET_NAME_3_SIGNS, 5);
  EXPECT_EQ(FM_ICON_SET_NAME_3_SYMBOLS, 6);
  EXPECT_EQ(FM_ICON_SET_NAME_3_SYMBOLS_2, 7);
  EXPECT_EQ(FM_ICON_SET_NAME_4_ARROWS, 8);
  EXPECT_EQ(FM_ICON_SET_NAME_4_ARROWS_GRAY, 9);
  EXPECT_EQ(FM_ICON_SET_NAME_4_RED_TO_BLACK, 10);
  EXPECT_EQ(FM_ICON_SET_NAME_4_RATING, 11);
  EXPECT_EQ(FM_ICON_SET_NAME_4_TRAFFIC_LIGHTS, 12);
  EXPECT_EQ(FM_ICON_SET_NAME_5_ARROWS, 13);
  EXPECT_EQ(FM_ICON_SET_NAME_5_ARROWS_GRAY, 14);
  EXPECT_EQ(FM_ICON_SET_NAME_5_RATING, 15);
  EXPECT_EQ(FM_ICON_SET_NAME_5_QUARTERS, 16);
  EXPECT_EQ(FM_ICON_SET_NAME_3_STARS, 17);
  EXPECT_EQ(FM_ICON_SET_NAME_3_TRIANGLES, 18);
  EXPECT_EQ(FM_ICON_SET_NAME_5_BOXES, 19);
  EXPECT_EQ(FM_ICON_SET_NAME_NO_ICONS, 20);
}

// A rule draws only a built-in set; the x14-only sets are for custom levels.
TEST(FormulonCApiCfMutate, RuleRejectsAnIconSetOnlyACustomLevelNames) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cfvo_t thresholds[4]{};
  for (fm_cfvo_t& t : thresholds) {
    t.type = 1;  // Percent
    t.value = "50";
    t.gte = 1;
  }
  const struct {
    int32_t name;
    std::uint32_t threshold_count;
  } kCases[] = {{FM_ICON_SET_NAME_3_STARS, 2U},
                {FM_ICON_SET_NAME_3_TRIANGLES, 2U},
                {FM_ICON_SET_NAME_5_BOXES, 4U},
                {FM_ICON_SET_NAME_NO_ICONS, 2U}};
  for (const auto& c : kCases) {
    fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
    rule.type = 4;  // IconSet
    rule.ranges = &sqref;
    rule.range_count = 1;
    rule.has_icon_set = 1;
    rule.icon_set_name = static_cast<std::uint8_t>(c.name);
    rule.icon_set_thresholds = thresholds;
    rule.icon_set_threshold_count = c.threshold_count;
    std::size_t index = 0;
    EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << c.name;
  }
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

// Text, blanks, errors and timePeriod rules authored without a formula carry
// the one Excel stores, anchored at the top-left corner of the sqref bounding box.
TEST(FormulonCApiCfMutate, AddRuleDerivesExcelsFormulaForSideFieldRules) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // B2:B10 and D5; the anchor is B2, the bounding box's top-left.
  fm_cell_rect_t ranges[2]{{1, 1, 9, 1}, {4, 3, 4, 3}};
  struct Case {
    std::uint8_t type;
    const char* text;
    int time_period;
    const char* formula1;
    const char* expected;
  };
  const Case cases[] = {
      {7, "ab", -1, nullptr, "NOT(ISERROR(SEARCH(\"ab\",B2)))"},
      {8, "q\"r", -1, "", "ISERROR(SEARCH(\"q\"\"r\",B2))"},
      {9, "x", -1, nullptr, "LEFT(B2,LEN(\"x\"))=\"x\""},
      {10, "y", -1, nullptr, "RIGHT(B2,LEN(\"y\"))=\"y\""},
      {11, nullptr, -1, nullptr, "LEN(TRIM(B2))=0"},
      {12, nullptr, -1, nullptr, "LEN(TRIM(B2))>0"},
      {13, nullptr, -1, nullptr, "ISERROR(B2)"},
      {14, nullptr, -1, nullptr, "NOT(ISERROR(B2))"},
      {15, nullptr, FM_DATE_PERIOD_YESTERDAY, nullptr, "FLOOR(B2,1)=TODAY()-1"},
      {15, nullptr, FM_DATE_PERIOD_LAST_MONTH, nullptr,
       "AND(MONTH(B2)=MONTH(EDATE(TODAY(),0-1)),YEAR(B2)=YEAR(EDATE(TODAY(),0-1)))"},
      // A caller-supplied formula is kept.
      {11, nullptr, -1, "=ISBLANK(B2)", "ISBLANK(B2)"},
  };
  for (const Case& c : cases) {
    fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
    rule.type = c.type;
    rule.text = c.text;
    rule.formula1 = c.formula1;
    if (c.time_period >= 0) {
      rule.has_time_period = 1;
      rule.time_period = static_cast<std::uint8_t>(c.time_period);
    }
    rule.ranges = ranges;
    rule.range_count = 2;
    std::size_t index = 0;
    ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
    fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
    ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, index, &out), 0);
    ASSERT_NE(out.formula1, nullptr) << "type=" << static_cast<int>(c.type);
    EXPECT_EQ(std::string(out.formula1), std::string("=") + c.expected);
  }

  // Both writers carry the derived formulas, and readers return them unchanged.
  for (std::int32_t format :
       {static_cast<std::int32_t>(FM_WORKBOOK_FORMAT_XLSX), static_cast<std::int32_t>(FM_WORKBOOK_FORMAT_XLSB)}) {
    BufferGuard saved;
    ASSERT_EQ(fm_workbook_save_as(wb.handle, format, &saved.data, &saved.len), 0)
        << "format=" << format << ": " << fm_get_last_error_message();
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
    std::size_t count = 0;
    ASSERT_EQ(fm_sheet_get_cf_rule_count(reloaded.handle, 0, &count), 0);
    ASSERT_EQ(count, std::size(cases));
    for (std::size_t i = 0; i < count; ++i) {
      fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
      ASSERT_EQ(fm_sheet_get_cf_rule_at(reloaded.handle, 0, i, &out), 0);
      ASSERT_NE(out.formula1, nullptr) << "format=" << format << " rule=" << i;
      EXPECT_EQ(std::string(out.formula1), std::string("=") + cases[i].expected)
          << "format=" << format << " rule=" << i;
    }
  }
}

// The CF time period table is in `ST_TimePeriod` order.
TEST(FormulonCApiCfMutate, DatePeriodOrdinalsFollowTheSchemaOrder) {
  EXPECT_EQ(FM_DATE_PERIOD_TODAY, 0);
  EXPECT_EQ(FM_DATE_PERIOD_YESTERDAY, 1);
  EXPECT_EQ(FM_DATE_PERIOD_TOMORROW, 2);
  EXPECT_EQ(FM_DATE_PERIOD_LAST_7_DAYS, 3);
  EXPECT_EQ(FM_DATE_PERIOD_THIS_MONTH, 4);
  EXPECT_EQ(FM_DATE_PERIOD_LAST_MONTH, 5);
  EXPECT_EQ(FM_DATE_PERIOD_NEXT_MONTH, 6);
  EXPECT_EQ(FM_DATE_PERIOD_THIS_WEEK, 7);
  EXPECT_EQ(FM_DATE_PERIOD_LAST_WEEK, 8);
  EXPECT_EQ(FM_DATE_PERIOD_NEXT_WEEK, 9);
}

// The anchor is the bounding box's top-left, not the first range's.
TEST(FormulonCApiCfMutate, DerivedFormulaAnchorsAtSqrefBoundingBox) {
  const fm_cell_rect_t d5_then_b2[2]{{4, 3, 5, 3}, {1, 1, 2, 1}};
  const fm_cell_rect_t b5_then_d2[2]{{4, 1, 5, 1}, {1, 3, 2, 3}};
  for (const fm_cell_rect_t* ranges : {d5_then_b2, b5_then_d2}) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
    rule.type = 7;
    rule.text = "x";
    rule.ranges = ranges;
    rule.range_count = 2;
    std::size_t index = 0;
    ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
    fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
    ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, index, &out), 0);
    ASSERT_NE(out.formula1, nullptr);
    EXPECT_STREQ(out.formula1, "=NOT(ISERROR(SEARCH(\"x\",B2)))");
  }
}

namespace {

fm_cf_rule_t ReadRule(const WorkbookGuard& wb, std::size_t index) {
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  EXPECT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, index, &out), 0) << fm_get_last_error_message();
  return out;
}

// Adds a rule over `range` and returns its index.
std::size_t AddRule(const WorkbookGuard& wb, fm_cf_rule_t rule, const fm_cell_rect_t* range) {
  rule.ranges = range;
  rule.range_count = 1;
  std::size_t index = 0;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  return index;
}

fm_cf_rule_t CellIsRule(const char* formula) {
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 1;  // CellIs
  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_GREATER_THAN;
  rule.formula1 = formula;
  return rule;
}

}  // namespace

TEST(FormulonCApiCfMutate, UpdateMergesIntoTheRuleAndKeepsIdAndPosition) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  SeedDxfs(wb.handle, 2);
  const fm_cell_rect_t a1_a10{0, 0, 9, 0};
  fm_cf_rule_t first = CellIsRule("50");
  first.has_dxf_index = 1;
  first.dxf_index = 1;
  first.stop_if_true = 1;
  ASSERT_EQ(AddRule(wb, first, &a1_a10), 0U);
  ASSERT_EQ(AddRule(wb, CellIsRule("7"), &a1_a10), 1U);
  const std::string id = ReadRule(wb, 0).id;
  const std::int32_t priority = ReadRule(wb, 0).priority;

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.formula1_op = FM_PATCH_OP_SET;
  patch.formula1 = "=60";
  patch.op_op = FM_PATCH_OP_SET;
  patch.op = FM_COMPARISON_OPERATOR_LESS_THAN;
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &patch), 0) << fm_get_last_error_message();

  const fm_cf_rule_t out = ReadRule(wb, 0);
  EXPECT_STREQ(out.id, id.c_str());
  EXPECT_EQ(out.priority, priority);
  EXPECT_STREQ(out.formula1, "=60");
  EXPECT_EQ(out.op, FM_COMPARISON_OPERATOR_LESS_THAN);
  EXPECT_EQ(out.has_dxf_index, 1);
  EXPECT_EQ(out.dxf_index, 1U);
  EXPECT_EQ(out.stop_if_true, 1);
  EXPECT_STREQ(ReadRule(wb, 1).formula1, "=7");
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 2U);
}

TEST(FormulonCApiCfMutate, UpdateClearsDxfIndexAndAcceptsAnEmptyPatch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  SeedDxfs(wb.handle, 1);
  const fm_cell_rect_t range{0, 0, 4, 0};
  fm_cf_rule_t rule = CellIsRule("1");
  rule.has_dxf_index = 1;
  ASSERT_EQ(AddRule(wb, rule, &range), 0U);

  fm_cf_rule_patch_t none = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &none), 0);
  EXPECT_EQ(ReadRule(wb, 0).has_dxf_index, 1);

  fm_cf_rule_patch_t clear = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  clear.dxf_index_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &clear), 0);
  EXPECT_EQ(ReadRule(wb, 0).has_dxf_index, 0);

  fm_cf_rule_patch_t past = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  past.dxf_index_op = FM_PATCH_OP_SET;
  past.dxf_index = 5;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &past), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ReadRule(wb, 0).has_dxf_index, 0);
}

TEST(FormulonCApiCfMutate, UpdateRefusesFieldsTheRuleTypeDoesNotUseAndChangesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 4, 0};
  ASSERT_EQ(AddRule(wb, CellIsRule("1"), &range), 0U);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.formula1_op = FM_PATCH_OP_SET;
  patch.formula1 = "=2";
  patch.rank_op = FM_PATCH_OP_SET;  // a Top10 field
  patch.rank = 3;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &patch), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_STREQ(ReadRule(wb, 0).formula1, "=1");

  fm_cf_rule_patch_t not_clearable = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  not_clearable.op_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &not_clearable), FM_STATUS_INVALID_ARGUMENT);

  fm_cf_rule_patch_t bad_op = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  bad_op.priority_op = static_cast<fm_patch_op_t>(9);
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &bad_op), FM_STATUS_INVALID_ARGUMENT);

  fm_cf_rule_patch_t null_text = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  null_text.formula1_op = FM_PATCH_OP_SET;
  null_text.formula1 = nullptr;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &null_text), FM_STATUS_INVALID_ARGUMENT);
}

TEST(FormulonCApiCfMutate, UpdateReportsMissingTargetsByStatus) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 4, 0};
  ASSERT_EQ(AddRule(wb, CellIsRule("1"), &range), 0U);
  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);

  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 1, &patch), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 9, 0, &patch), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_update_cf_rule(nullptr, 0, 0, &patch), FM_STATUS_BINDING_NULL_POINTER);

  fm_cf_rule_patch_t null_ranges = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  null_ranges.ranges_op = FM_PATCH_OP_SET;
  null_ranges.range_count = 1;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &null_ranges), FM_STATUS_BINDING_NULL_POINTER);
}

TEST(FormulonCApiCfMutate, UpdateRederivesTheFormulaOfATextRule) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t b2_b5{1, 1, 4, 1};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 7;  // ContainsText
  rule.text = "ab";
  ASSERT_EQ(AddRule(wb, rule, &b2_b5), 0U);
  EXPECT_STREQ(ReadRule(wb, 0).formula1, "=NOT(ISERROR(SEARCH(\"ab\",B2)))");

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.text_op = FM_PATCH_OP_SET;
  patch.text = "cd";
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &patch), 0) << fm_get_last_error_message();
  EXPECT_STREQ(ReadRule(wb, 0).formula1, "=NOT(ISERROR(SEARCH(\"cd\",B2)))");

  const fm_cell_rect_t d4_d6{3, 3, 5, 3};
  fm_cf_rule_patch_t move = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  move.ranges_op = FM_PATCH_OP_SET;
  move.ranges = &d4_d6;
  move.range_count = 1;
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &move), 0);
  const fm_cf_rule_t out = ReadRule(wb, 0);
  EXPECT_STREQ(out.formula1, "=NOT(ISERROR(SEARCH(\"cd\",D4)))");
  ASSERT_EQ(out.range_count, 1U);
  EXPECT_EQ(out.ranges[0].first_row, 3U);
  EXPECT_EQ(out.ranges[0].last_row, 5U);

  fm_cf_rule_patch_t empty_ranges = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  empty_ranges.ranges_op = FM_PATCH_OP_SET;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &empty_ranges), FM_STATUS_INVALID_ARGUMENT);
}

TEST(FormulonCApiCfMutate, UpdateOfOneRuleInASharedBlockKeepsTheFlattenedOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  formulon::cf::ConditionalFormat block;
  block.sqref = {MakeRange(0, 0, 4, 0)};
  for (int i = 0; i < 3; ++i) {
    formulon::cf::CFRule rule;
    rule.type = formulon::cf::RuleType::Expression;
    rule.priority = i + 1;
    rule.formula1 = std::to_string(i + 1) + ">0";
    block.rules.push_back(rule);
  }
  wb.handle->workbook().sheet(0).mutable_conditional_formats().push_back(block);

  const fm_cell_rect_t moved{0, 2, 4, 2};
  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.ranges_op = FM_PATCH_OP_SET;
  patch.ranges = &moved;
  patch.range_count = 1;
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 1, &patch), 0) << fm_get_last_error_message();

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 3U);
  const fm_cf_rule_t r0 = ReadRule(wb, 0);
  EXPECT_STREQ(r0.formula1, "=1>0");
  EXPECT_EQ(r0.ranges[0].first_col, 0U);
  const fm_cf_rule_t r1 = ReadRule(wb, 1);
  EXPECT_STREQ(r1.formula1, "=2>0");
  EXPECT_EQ(r1.ranges[0].first_col, 2U);
  const fm_cf_rule_t r2 = ReadRule(wb, 2);
  EXPECT_STREQ(r2.formula1, "=3>0");
  EXPECT_EQ(r2.ranges[0].first_col, 0U);
}

TEST(FormulonCApiCfMutate, UpdateMergesDataBarSettingsAndChecksTheMergedBar) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 9, 0};
  ASSERT_EQ(AddRule(wb, MakeDataBarRule(&range), &range), 0U);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.data_bar_fill_op = FM_PATCH_OP_SET;
  patch.data_bar_fill = CfRgb(1, 2, 3);
  patch.data_bar_gradient_op = FM_PATCH_OP_SET;
  patch.data_bar_gradient = 0;
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &patch), 0) << fm_get_last_error_message();
  fm_cf_rule_t out = ReadRule(wb, 0);
  EXPECT_EQ(CfR(out.data_bar_fill), 1U);
  EXPECT_EQ(CfB(out.data_bar_fill), 3U);
  EXPECT_EQ(out.data_bar_gradient, 0);
  EXPECT_EQ(out.data_bar_min_length_pct, 10U);
  EXPECT_EQ(out.data_bar_max_length_pct, 90U);

  // The merged lengths must stay ordered, and none may pass 100.
  fm_cf_rule_patch_t inverted = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  inverted.data_bar_min_length_pct_op = FM_PATCH_OP_SET;
  inverted.data_bar_min_length_pct = 95;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &inverted), FM_STATUS_INVALID_ARGUMENT);
  fm_cf_rule_patch_t too_long = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  too_long.data_bar_max_length_pct_op = FM_PATCH_OP_SET;
  too_long.data_bar_max_length_pct = 300;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &too_long), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ReadRule(wb, 0).data_bar_min_length_pct, 10U);
}

TEST(FormulonCApiCfMutate, UpdateColorScaleNeedsMatchingThresholdAndColorCounts) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 9, 0};
  fm_cf_rule_t scale = FM_STRUCT_INIT(fm_cf_rule_t);
  scale.type = 2;  // ColorScale
  fm_cfvo_t thresholds[2]{};
  thresholds[0].type = 3;  // Min
  thresholds[1].type = 4;  // Max
  fm_color_t colors[2] = {CfRgb(255, 0, 0), CfRgb(0, 255, 0)};
  scale.color_scale_thresholds = thresholds;
  scale.color_scale_colors = colors;
  scale.color_scale_count = 2;
  ASSERT_EQ(AddRule(wb, scale, &range), 0U);

  fm_color_t recolored[2] = {CfRgb(9, 9, 9), CfRgb(8, 8, 8)};
  fm_cf_rule_patch_t colors_only = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  colors_only.color_scale_colors_op = FM_PATCH_OP_SET;
  colors_only.color_scale_colors = recolored;
  colors_only.color_scale_colors_count = 2;
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &colors_only), 0) << fm_get_last_error_message();
  EXPECT_EQ(CfR(ReadRule(wb, 0).color_scale_colors[1]), 8U);

  fm_cf_rule_patch_t three_colors = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  fm_color_t more[3] = {CfRgb(1, 1, 1), CfRgb(2, 2, 2), CfRgb(3, 3, 3)};
  three_colors.color_scale_colors_op = FM_PATCH_OP_SET;
  three_colors.color_scale_colors = more;
  three_colors.color_scale_colors_count = 3;
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, 0, &three_colors), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ReadRule(wb, 0).color_scale_count, 2U);
}
