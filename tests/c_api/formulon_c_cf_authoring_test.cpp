// Stable C ABI conditional-format authoring tests.

#include <iterator>

#include "formulon_c_cf_test_helpers.h"

TEST(FormulonCApiCfMutate, AddCellIsRuleRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDxfs(wb.handle, 1);

  fm_cf_cell_range_t sqref{};
  sqref.first_row = 0;
  sqref.first_col = 0;
  sqref.last_row = 9;
  sqref.last_col = 0;

  fm_cf_rule_t rule{};
  rule.type = 1;  // CellIs
  rule.op_engaged = 1;
  rule.op = 5;  // GreaterThan
  std::string formula = "50";
  rule.formula1 = formula.c_str();
  rule.dxf_id_engaged = 1;
  rule.dxf_id = 0;
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0);
  EXPECT_EQ(rule_index, 0U);

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_cf_rule_t out{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.type, 1U);
  EXPECT_EQ(out.op_engaged, 1);
  EXPECT_EQ(out.op, 5U);
  EXPECT_EQ(out.priority, 1);  // auto-assigned since input <= 0
  EXPECT_EQ(out.dxf_id_engaged, 1);
  EXPECT_EQ(out.dxf_id, 0U);
  ASSERT_NE(out.formula1, nullptr);
  EXPECT_STREQ(out.formula1, "50");
  ASSERT_NE(out.sqref, nullptr);
  EXPECT_EQ(out.sqref_count, 1U);
  EXPECT_EQ(out.sqref[0].first_row, 0U);
  EXPECT_EQ(out.sqref[0].last_row, 9U);
  ASSERT_NE(out.id, nullptr);
  EXPECT_FALSE(std::string(out.id).empty());
}

TEST(FormulonCApiCfMutate, AddRuleRejectsDxfIdPastTheStylesTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDxfs(wb.handle, 2);

  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 1;  // CellIs
  rule.op_engaged = 1;
  rule.op = 5;  // GreaterThan
  rule.formula1 = "50";
  rule.dxf_id_engaged = 1;
  rule.dxf_id = 2;  // one past the last registered `<dxf>`
  rule.sqref = &sqref;
  rule.sqref_count = 1;

  std::size_t rule_index = 99U;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_STREQ(fm_last_error_message(), "fm_sheet_cf_add_rule: dxf_id out of range");

  // The rejected rule is not appended, so a workbook saved after the failed
  // call carries no `<cfRule>` whose `dxfId` Excel cannot resolve.
  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  // The last in-range index is still accepted.
  rule.dxf_id = 1;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiCfMutate, AddRuleWithoutDxfIdIgnoresTheEmptyStylesTable) {
  // A rule that engages no differential format must stay addable on a
  // workbook that has registered no `<dxf>` at all.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 1;  // CellIs
  rule.op_engaged = 1;
  rule.op = 5;  // GreaterThan
  rule.formula1 = "50";
  rule.sqref = &sqref;
  rule.sqref_count = 1;

  std::size_t rule_index = 99U;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();
  EXPECT_EQ(rule_index, 0U);

  fm_cf_rule_t out{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.dxf_id_engaged, 0);
}

TEST(FormulonCApiCfMutate, AddMultipleRulesAutoIncrementsPriority) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cf_cell_range_t sqref{0, 0, 0, 0};
  for (int i = 0; i < 3; ++i) {
    fm_cf_rule_t rule{};
    rule.type = 0;  // Expression
    std::string f = "TRUE";
    rule.formula1 = f.c_str();
    rule.sqref = &sqref;
    rule.sqref_count = 1;
    std::size_t rule_index = 0;
    ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0);
    EXPECT_EQ(rule_index, static_cast<std::size_t>(i));
  }

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 3U);
  for (std::size_t i = 0; i < 3; ++i) {
    fm_cf_rule_t out{};
    ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, i, &out), 0);
    EXPECT_EQ(out.priority, static_cast<int32_t>(i + 1));
  }
}

TEST(FormulonCApiCfMutate, RemoveAtFlattensIndices) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cf_cell_range_t sqref{0, 0, 0, 0};
  std::vector<std::string> formulas{"A1", "A2", "A3"};
  std::size_t expected_index = 0;
  for (const auto& f : formulas) {
    fm_cf_rule_t rule{};
    rule.type = 0;  // Expression
    rule.formula1 = f.c_str();
    rule.sqref = &sqref;
    rule.sqref_count = 1;
    std::size_t rule_index = 0;
    ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0);
    EXPECT_EQ(rule_index, expected_index);
    ++expected_index;
  }
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 3U);

  ASSERT_EQ(fm_sheet_cf_remove_at(wb.handle, 0, 1), 0);
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  // A record's storage belongs to the handle and the next get takes it
  // back, so each rule is read before the following one is fetched.
  fm_cf_rule_t out0{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out0), 0);
  EXPECT_STREQ(out0.formula1, "A1");
  fm_cf_rule_t out1{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 1, &out1), 0);
  EXPECT_STREQ(out1.formula1, "A3");
}

TEST(FormulonCApiCfMutate, ClearRemovesAllBlocks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 0, 0};
  fm_cf_rule_t rule{};
  rule.type = 0;
  std::string f = "TRUE";
  rule.formula1 = f.c_str();
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0);
  EXPECT_EQ(rule_index, 0U);

  ASSERT_EQ(fm_sheet_cf_clear(wb.handle, 0), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiCfMutate, AddsVisualRuleTypesAndPreservesThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 0, 0};

  fm_cf_rule_t rule{};
  rule.type = 2;  // ColorScale
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  fm_cfvo_t thresholds[3]{};
  thresholds[0].type = 3;  // Min
  thresholds[0].gte = 1;
  thresholds[1].type = 1;  // Percent
  thresholds[1].value = "50";
  thresholds[1].gte = 1;
  thresholds[2].type = 4;  // Max
  thresholds[2].gte = 1;
  fm_cf_color_t colors[3]{{255, 0, 0, 255}, {255, 255, 0, 255}, {0, 255, 0, 255}};
  rule.color_scale_thresholds = thresholds;
  rule.color_scale_colors = colors;
  rule.color_scale_count = 3;
  std::size_t color_rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &color_rule_index), 0);
  EXPECT_EQ(color_rule_index, 0U);

  fm_cf_cell_range_t db_sqref{1, 0, 1, 0};
  fm_cf_rule_t db_rule{};
  db_rule.type = 3;  // DataBar
  db_rule.sqref = &db_sqref;
  db_rule.sqref_count = 1;
  db_rule.data_bar_engaged = 1;
  db_rule.data_bar_min.type = 3;  // Min
  db_rule.data_bar_min.gte = 1;
  db_rule.data_bar_max.type = 4;  // Max
  db_rule.data_bar_max.gte = 1;
  db_rule.data_bar_fill = fm_cf_color_t{99, 142, 198, 255};
  db_rule.data_bar_show_value = 1;
  db_rule.data_bar_min_length_pct = 10;
  db_rule.data_bar_max_length_pct = 90;
  std::size_t data_bar_rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, db_rule, &data_bar_rule_index), 0);
  EXPECT_EQ(data_bar_rule_index, 1U);

  fm_cf_cell_range_t icon_sqref{2, 0, 2, 0};
  fm_cf_rule_t icon_rule{};
  icon_rule.type = 4;  // IconSet
  icon_rule.sqref = &icon_sqref;
  icon_rule.sqref_count = 1;
  icon_rule.icon_set_engaged = 1;
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
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, icon_rule, &icon_rule_index), 0);
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

  fm_cf_rule_t out_color{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out_color), 0);
  ASSERT_EQ(out_color.color_scale_count, 3U);
  ASSERT_NE(out_color.color_scale_thresholds, nullptr);
  ASSERT_NE(out_color.color_scale_colors, nullptr);
  EXPECT_EQ(out_color.color_scale_thresholds[1].type, 1U);
  EXPECT_STREQ(out_color.color_scale_thresholds[1].value, "50");
  EXPECT_EQ(out_color.color_scale_colors[2].g, 255U);

  fm_cf_rule_t out_bar{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 1, &out_bar), 0);
  EXPECT_EQ(out_bar.data_bar_engaged, 1);
  EXPECT_EQ(out_bar.data_bar_fill.b, 198U);
  EXPECT_EQ(out_bar.data_bar_min_length_pct, 10U);
  EXPECT_EQ(out_bar.data_bar_max_length_pct, 90U);

  fm_cf_rule_t out_icon{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 2, &out_icon), 0);
  EXPECT_EQ(out_icon.icon_set_engaged, 1);
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 1;  // CellIs
  rule.op_engaged = 1;
  rule.op = 6;  // Between
  rule.formula1 = "=$B1>100";
  rule.formula2 = "=SUM($C$1:$C$2)";
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();

  fm_cfvo_t thresholds[2]{};
  thresholds[0].type = 5;  // Formula
  thresholds[0].value = "=$D$1";
  thresholds[0].gte = 1;
  thresholds[1].type = 4;  // Max
  thresholds[1].gte = 1;
  fm_cf_color_t colors[2]{{255, 0, 0, 255}, {0, 255, 0, 255}};
  fm_cf_rule_t scale{};
  scale.type = 2;  // ColorScale
  scale.sqref = &sqref;
  scale.sqref_count = 1;
  scale.color_scale_thresholds = thresholds;
  scale.color_scale_colors = colors;
  scale.color_scale_count = 2;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, scale, &index), 0) << fm_last_error_message();

  fm_cf_rule_t out{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.formula1, "$B1>100");
  EXPECT_STREQ(out.formula2, "SUM($C$1:$C$2)");
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 1, &out), 0);
  ASSERT_EQ(out.color_scale_count, 2U);
  EXPECT_STREQ(out.color_scale_thresholds[0].value, "$D$1");

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  ASSERT_EQ(fm_sheet_cf_get_at(reloaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.formula1, "$B1>100");
  EXPECT_STREQ(out.formula2, "SUM($C$1:$C$2)");

  // Only one '=' is the entry marker; a formula that still starts with one
  // after it does not parse.
  rule.formula1 = "==$B1>100";
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken));
  thresholds[0].value = "=SUM(";
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, scale, &index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken));
}

// An N-icon set carries exactly N-1 thresholds; the floor is separate.
TEST(FormulonCApiCfMutate, AddRuleRejectsIconThresholdCountsOtherThanIconsMinusOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 4, 0};
  fm_cfvo_t thresholds[4]{};
  for (fm_cfvo_t& t : thresholds) {
    t.type = 1;  // Percent
    t.value = "50";
    t.gte = 1;
  }
  fm_cf_rule_t rule{};
  rule.type = 4;  // IconSet
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  rule.icon_set_engaged = 1;
  rule.icon_set_name = 0;  // Three_Arrows
  rule.icon_set_thresholds = thresholds;
  std::size_t index = 0;
  for (std::uint32_t count : {0U, 1U, 3U, 4U}) {
    rule.icon_set_threshold_count = count;
    EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << "count=" << count;
  }
  rule.icon_set_threshold_count = 2;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();
  rule.icon_set_name = 8;  // Four_Arrows
  EXPECT_NE(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0);
  rule.icon_set_threshold_count = 3;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();
  rule.icon_set_name = 16;  // Five_Quarters
  rule.icon_set_threshold_count = 4;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 3U);
}

// Text, blanks, errors and timePeriod rules authored without a formula carry
// the one Excel stores, anchored at the top-left corner of the sqref bounding box.
TEST(FormulonCApiCfMutate, AddRuleDerivesExcelsFormulaForSideFieldRules) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // B2:B10 and D5; the anchor is B2, the bounding box's top-left.
  fm_cf_cell_range_t ranges[2]{{1, 1, 9, 1}, {4, 3, 4, 3}};
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
      {15, nullptr, 1, nullptr, "FLOOR(B2,1)=TODAY()-1"},
      {15, nullptr, 8, nullptr, "AND(MONTH(B2)=MONTH(EDATE(TODAY(),0-1)),YEAR(B2)=YEAR(EDATE(TODAY(),0-1)))"},
      // A caller-supplied formula is kept.
      {11, nullptr, -1, "=ISBLANK(B2)", "ISBLANK(B2)"},
  };
  for (const Case& c : cases) {
    fm_cf_rule_t rule{};
    rule.type = c.type;
    rule.text = c.text;
    rule.formula1 = c.formula1;
    if (c.time_period >= 0) {
      rule.time_period_engaged = 1;
      rule.time_period = static_cast<std::uint8_t>(c.time_period);
    }
    rule.sqref = ranges;
    rule.sqref_count = 2;
    std::size_t index = 0;
    ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();
    fm_cf_rule_t out{};
    ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, index, &out), 0);
    ASSERT_NE(out.formula1, nullptr) << "type=" << static_cast<int>(c.type);
    EXPECT_STREQ(out.formula1, c.expected);
  }

  // Both writers carry the derived formulas, and readers return them unchanged.
  for (std::int32_t format :
       {static_cast<std::int32_t>(FM_WORKBOOK_FORMAT_XLSX), static_cast<std::int32_t>(FM_WORKBOOK_FORMAT_XLSB)}) {
    BufferGuard saved;
    ASSERT_EQ(fm_workbook_save_as(wb.handle, format, &saved.data, &saved.len), 0)
        << "format=" << format << ": " << fm_last_error_message();
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
    std::size_t count = 0;
    ASSERT_EQ(fm_sheet_cf_count(reloaded.handle, 0, &count), 0);
    ASSERT_EQ(count, std::size(cases));
    for (std::size_t i = 0; i < count; ++i) {
      fm_cf_rule_t out{};
      ASSERT_EQ(fm_sheet_cf_get_at(reloaded.handle, 0, i, &out), 0);
      ASSERT_NE(out.formula1, nullptr) << "format=" << format << " rule=" << i;
      EXPECT_STREQ(out.formula1, cases[i].expected) << "format=" << format << " rule=" << i;
    }
  }
}

// The anchor is the bounding box's top-left, not the first range's.
TEST(FormulonCApiCfMutate, DerivedFormulaAnchorsAtSqrefBoundingBox) {
  const fm_cf_cell_range_t d5_then_b2[2]{{4, 3, 5, 3}, {1, 1, 2, 1}};
  const fm_cf_cell_range_t b5_then_d2[2]{{4, 1, 5, 1}, {1, 3, 2, 3}};
  for (const fm_cf_cell_range_t* ranges : {d5_then_b2, b5_then_d2}) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cf_rule_t rule{};
    rule.type = 7;
    rule.text = "x";
    rule.sqref = ranges;
    rule.sqref_count = 2;
    std::size_t index = 0;
    ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();
    fm_cf_rule_t out{};
    ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, index, &out), 0);
    ASSERT_NE(out.formula1, nullptr);
    EXPECT_STREQ(out.formula1, "NOT(ISERROR(SEARCH(\"x\",B2)))");
  }
}
