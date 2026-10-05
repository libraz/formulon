// Stable C ABI conditional-format visual payload tests.

#include <utility>

#include "formulon_c_cf_test_helpers.h"

namespace {

const formulon::cf::DataBarSpec& OnlyDataBarSpec(fm_workbook_t* handle) {
  const auto& blocks = handle->workbook().sheet(0).conditional_formats();
  EXPECT_EQ(blocks.size(), 1U);
  EXPECT_EQ(blocks[0].rules.size(), 1U);
  EXPECT_TRUE(blocks[0].rules[0].data_bar.has_value());
  return *blocks[0].rules[0].data_bar;
}

}  // namespace

TEST(FormulonCApiCfMutate, ZeroInitializedDataBarExtensionKeepsTheModelDefaults) {
  // The extension fields were appended to a struct callers already
  // zero-initialize. Leaving every `*_engaged` flag at zero must produce
  // exactly the spec the ABI produced before those fields existed.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  const fm_cf_rule_t rule = MakeDataBarRule(&sqref);
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();

  const formulon::cf::DataBarSpec& spec = OnlyDataBarSpec(wb.handle);
  EXPECT_TRUE(spec.gradient);
  EXPECT_EQ(spec.axis_position, formulon::cf::DataBarAxisPosition::Automatic);
  // Unengaged negative fill still mirrors the positive fill.
  EXPECT_EQ(spec.negative_fill.r, spec.fill.r);
  EXPECT_EQ(spec.negative_fill.g, spec.fill.g);
  EXPECT_EQ(spec.negative_fill.b, spec.fill.b);
  EXPECT_EQ(spec.negative_fill.a, spec.fill.a);
  EXPECT_FALSE(spec.border.has_value());
  EXPECT_FALSE(spec.negative_border.has_value());
  EXPECT_EQ(spec.axis_color.r, 0U);
  EXPECT_EQ(spec.axis_color.g, 0U);
  EXPECT_EQ(spec.axis_color.b, 0U);
  EXPECT_EQ(spec.axis_color.a, 255U);
}

TEST(FormulonCApiCfMutate, DataBarExtensionFieldsReachTheModelWhenEngaged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule = MakeDataBarRule(&sqref);
  rule.data_bar_gradient_engaged = 1;
  rule.data_bar_gradient = 0;  // solid, not gradient
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

  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();

  const formulon::cf::DataBarSpec& spec = OnlyDataBarSpec(wb.handle);
  EXPECT_FALSE(spec.gradient);
  EXPECT_EQ(spec.axis_position, formulon::cf::DataBarAxisPosition::Middle);
  EXPECT_EQ(spec.negative_fill.r, 255U);
  EXPECT_EQ(spec.negative_fill.g, 0U);
  ASSERT_TRUE(spec.border.has_value());
  EXPECT_EQ(spec.border->r, 1U);
  EXPECT_EQ(spec.border->b, 3U);
  ASSERT_TRUE(spec.negative_border.has_value());
  EXPECT_EQ(spec.negative_border->r, 4U);
  EXPECT_EQ(spec.negative_border->b, 6U);
  EXPECT_EQ(spec.axis_color.r, 7U);
  EXPECT_EQ(spec.axis_color.b, 9U);
}

TEST(FormulonCApiCfMutate, DataBarRuleSurvivesGetAtFedBackIntoAddRule) {
  // The only way to edit a CF rule through this ABI is get -> remove ->
  // add, so that path has to be an identity on the rule model, and the
  // record has to stay readable across the removal without the caller
  // copying anything out of it. Before the extension fields existed the
  // sequence silently reset axis, gradient and the negative colours to
  // their defaults.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t seed = MakeDataBarRule(&sqref);
  seed.data_bar_gradient_engaged = 1;
  seed.data_bar_gradient = 0;
  seed.data_bar_axis_position_engaged = 1;
  seed.data_bar_axis_position = 1;  // middle
  seed.data_bar_negative_fill_engaged = 1;
  seed.data_bar_negative_fill = fm_cf_color_t{200, 30, 40, 255};
  seed.data_bar_border_engaged = 1;
  seed.data_bar_border = fm_cf_color_t{11, 22, 33, 255};
  seed.data_bar_axis_color_engaged = 1;
  seed.data_bar_axis_color = fm_cf_color_t{44, 55, 66, 255};
  std::size_t seed_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, seed, &seed_index), 0) << fm_last_error_message();

  // Read it back, change one unrelated field, and write it out again.
  fm_cf_rule_t round_trip{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &round_trip), 0);
  EXPECT_EQ(round_trip.data_bar_gradient_engaged, 1);
  EXPECT_EQ(round_trip.data_bar_gradient, 0);
  EXPECT_EQ(round_trip.data_bar_axis_position_engaged, 1);
  EXPECT_EQ(round_trip.data_bar_axis_position, 1U);
  EXPECT_EQ(round_trip.data_bar_negative_fill_engaged, 1);
  EXPECT_EQ(round_trip.data_bar_negative_fill.r, 200U);
  EXPECT_EQ(round_trip.data_bar_border_engaged, 1);
  EXPECT_EQ(round_trip.data_bar_border.g, 22U);
  // No negative border was set, so the getter must report it absent
  // rather than handing back an engaged black.
  EXPECT_EQ(round_trip.data_bar_negative_border_engaged, 0);
  EXPECT_EQ(round_trip.data_bar_axis_color_engaged, 1);
  EXPECT_EQ(round_trip.data_bar_axis_color.b, 66U);

  round_trip.stop_if_true = 1;
  ASSERT_EQ(fm_sheet_cf_remove_at(wb.handle, 0, 0), 0);
  std::size_t rewritten_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, round_trip, &rewritten_index), 0) << fm_last_error_message();

  const formulon::cf::DataBarSpec& spec = OnlyDataBarSpec(wb.handle);
  EXPECT_FALSE(spec.gradient);
  EXPECT_EQ(spec.axis_position, formulon::cf::DataBarAxisPosition::Middle);
  EXPECT_EQ(spec.negative_fill.r, 200U);
  EXPECT_EQ(spec.negative_fill.g, 30U);
  ASSERT_TRUE(spec.border.has_value());
  EXPECT_EQ(spec.border->b, 33U);
  EXPECT_FALSE(spec.negative_border.has_value());
  EXPECT_EQ(spec.axis_color.g, 55U);
}

TEST(FormulonCApiCfMutate, GetAtRecordStaysReadableUntilTheNextGet) {
  // The record the getter fills has to outlive the mutations of the ABI's
  // own edit procedure, and unrelated reads that recycle the handle's text
  // scratch must not disturb it either. Only the next successful CF get
  // takes the storage back.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 5, 5, "unrelated"), 0);

  fm_cf_cell_range_t sqref[2]{{0, 0, 9, 0}, {4, 3, 14, 3}};
  fm_cf_rule_t seed{};
  seed.id = "{AB000000-0000-0000-0000-000000000001}";
  seed.type = 2;  // ColorScale
  seed.sqref = sqref;
  seed.sqref_count = 2;
  seed.formula1 = "SEEDFORMULA";
  fm_cfvo_t thresholds[2]{};
  thresholds[0].type = 0;  // Number
  thresholds[0].value = "11";
  thresholds[1].type = 0;
  thresholds[1].value = "33";
  fm_cf_color_t colors[2]{{255, 0, 0, 255}, {0, 255, 0, 255}};
  seed.color_scale_thresholds = thresholds;
  seed.color_scale_colors = colors;
  seed.color_scale_count = 2;
  std::size_t seed_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, seed, &seed_index), 0) << fm_last_error_message();

  fm_cf_rule_t held{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &held), 0);
  ASSERT_NE(held.sqref, nullptr);
  ASSERT_EQ(held.sqref_count, 2U);
  ASSERT_EQ(held.color_scale_count, 2U);

  // A removal frees the block the rule lived in.
  ASSERT_EQ(fm_sheet_cf_remove_at(wb.handle, 0, 0), 0);
  std::size_t remaining = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &remaining), 0);
  ASSERT_EQ(remaining, 0U);
  EXPECT_STREQ(held.id, "{AB000000-0000-0000-0000-000000000001}");
  EXPECT_STREQ(held.formula1, "SEEDFORMULA");
  EXPECT_EQ(held.sqref[1].first_col, 3U);
  EXPECT_EQ(held.sqref[1].last_row, 14U);
  EXPECT_STREQ(held.color_scale_thresholds[1].value, "33");

  // An unrelated text read recycles `read_scratch`.
  fm_value_t unrelated{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 5, 5, &unrelated), 0);
  EXPECT_STREQ(held.id, "{AB000000-0000-0000-0000-000000000001}");
  EXPECT_STREQ(held.color_scale_thresholds[0].value, "11");

  // Feeding the held record straight back is the documented edit path; it
  // must not need a caller-side copy of anything.
  std::size_t rewritten_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, held, &rewritten_index), 0) << fm_last_error_message();
  const auto& blocks = wb.handle->workbook().sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 1U);
  ASSERT_EQ(blocks[0].sqref.size(), 2U);
  EXPECT_EQ(blocks[0].sqref[1], MakeRange(4, 3, 14, 3));
  EXPECT_EQ(blocks[0].rules[0].id, "{AB000000-0000-0000-0000-000000000001}");
  ASSERT_TRUE(blocks[0].rules[0].color_scale.has_value());
  EXPECT_EQ(blocks[0].rules[0].color_scale->thresholds[1].value, "33");
}

TEST(FormulonCApiCfMutate, DataBarRejectsInvertedLengthsAndUnknownAxisPosition) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_cell_range_t sqref{0, 0, 9, 0};

  fm_cf_rule_t inverted = MakeDataBarRule(&sqref);
  inverted.data_bar_min_length_pct = 90;
  inverted.data_bar_max_length_pct = 10;
  std::size_t rule_index = 0;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, inverted, &rule_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  fm_cf_rule_t bad_axis = MakeDataBarRule(&sqref);
  bad_axis.data_bar_axis_position_engaged = 1;
  bad_axis.data_bar_axis_position = 3;  // one past `none`
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, bad_axis, &rule_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // An unengaged out-of-domain axis ordinal is ignored, not rejected: the
  // flag is what makes the value meaningful.
  fm_cf_rule_t stale_axis = MakeDataBarRule(&sqref);
  stale_axis.data_bar_axis_position = 3;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, stale_axis, &rule_index), 0) << fm_last_error_message();
  EXPECT_EQ(OnlyDataBarSpec(wb.handle).axis_position, formulon::cf::DataBarAxisPosition::Automatic);
}

TEST(FormulonCApiCfMutate, ThresholdStringsStayLiveWhileLaterPayloadsArePulled) {
  // Mirrors what the JS bindings do for a rule object that carries more
  // than one payload key: every threshold string is pulled into one
  // caller-side store, and the record keeps a borrowed view of each. The
  // store must therefore not relocate bytes an earlier pull already
  // published into the record - which is what a growing vector does on
  // the second short string it holds.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  formulon::c_api::BorrowedStringArena strings;

  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 2;  // ColorScale
  rule.sqref = &sqref;
  rule.sqref_count = 1;

  fm_cfvo_t thresholds[3]{};
  const char* const kThresholdValues[3] = {"11", "22", "33"};
  for (std::size_t i = 0; i < 3; ++i) {
    thresholds[i].type = 0;  // Number
    thresholds[i].gte = 1;
    thresholds[i].value = strings.emplace(kThresholdValues[i]);
  }
  fm_cf_color_t colors[3]{{255, 0, 0, 255}, {255, 255, 0, 255}, {0, 255, 0, 255}};
  rule.color_scale_thresholds = thresholds;
  rule.color_scale_colors = colors;
  rule.color_scale_count = 3;

  // The same JS object also carried `dataBar` and `iconSet` keys, so the
  // binding pulls their threshold strings into the same store after the
  // colorScale views are already in the record.
  rule.data_bar_engaged = 1;
  rule.data_bar_min.type = 0;
  rule.data_bar_min.value = strings.emplace("5");
  rule.data_bar_max.type = 0;
  rule.data_bar_max.value = strings.emplace("95");
  fm_cfvo_t icon_thresholds[2]{};
  icon_thresholds[0].type = 1;  // Percent
  icon_thresholds[0].value = strings.emplace("40");
  icon_thresholds[1].type = 1;
  icon_thresholds[1].value = strings.emplace("80");
  rule.icon_set_engaged = 1;
  rule.icon_set_thresholds = icon_thresholds;
  rule.icon_set_threshold_count = 2;

  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();

  const auto& blocks = wb.handle->workbook().sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 1U);
  ASSERT_TRUE(blocks[0].rules[0].color_scale.has_value());
  const auto& stored = blocks[0].rules[0].color_scale->thresholds;
  ASSERT_EQ(stored.size(), 3U);
  EXPECT_EQ(stored[0].value, "11");
  EXPECT_EQ(stored[1].value, "22");
  EXPECT_EQ(stored[2].value, "33");

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  const auto& reloaded_blocks = reloaded.handle->workbook().sheet(0).conditional_formats();
  ASSERT_EQ(reloaded_blocks.size(), 1U);
  ASSERT_TRUE(reloaded_blocks[0].rules[0].color_scale.has_value());
  const auto& reloaded_thresholds = reloaded_blocks[0].rules[0].color_scale->thresholds;
  ASSERT_EQ(reloaded_thresholds.size(), 3U);
  EXPECT_EQ(reloaded_thresholds[0].value, "11");
  EXPECT_EQ(reloaded_thresholds[1].value, "22");
  EXPECT_EQ(reloaded_thresholds[2].value, "33");
}

TEST(FormulonCApiCfMutate, RuleEngagingSeveralVisualPayloadsSurfacesLivePointers) {
  // `cf::CFRule` declares the three visual payloads mutually exclusive,
  // but a rule can only be trusted to honour that once every producer
  // does. The getter must keep every pointer it writes into the record
  // valid until it returns, including for a rule that engages two.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  auto& sheet = wb.handle->workbook().sheet(0);
  formulon::cf::ConditionalFormat block{};
  block.sqref.push_back(MakeRange(0, 0, 9, 0));
  formulon::cf::CFRule rule;
  rule.type = formulon::cf::RuleType::ColorScale;
  rule.priority = 1;
  formulon::cf::ColorScaleSpec color_scale;
  for (const char* value : {"11", "22", "33"}) {
    formulon::cf::CfValueObject cfvo;
    cfvo.type = formulon::cf::CfvoType::Number;
    cfvo.value = value;
    color_scale.thresholds.push_back(cfvo);
    color_scale.colors.push_back(formulon::cf::Color{255, 0, 0, 255});
  }
  rule.color_scale = std::move(color_scale);
  formulon::cf::IconSetSpec icon_set;
  for (const char* value : {"40", "80"}) {
    formulon::cf::CfValueObject cfvo;
    cfvo.type = formulon::cf::CfvoType::Percent;
    cfvo.value = value;
    icon_set.thresholds.push_back(cfvo);
  }
  rule.icon_set = std::move(icon_set);
  block.rules.push_back(std::move(rule));
  sheet.mutable_conditional_formats().push_back(std::move(block));

  fm_cf_rule_t out{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out), 0);
  ASSERT_EQ(out.color_scale_count, 3U);
  ASSERT_NE(out.color_scale_thresholds, nullptr);
  ASSERT_NE(out.color_scale_colors, nullptr);
  // Read after the iconSet payload was filled: the colorScale array must
  // not have moved out from under the pointer already in the record.
  EXPECT_STREQ(out.color_scale_thresholds[0].value, "11");
  EXPECT_STREQ(out.color_scale_thresholds[2].value, "33");
  EXPECT_EQ(out.color_scale_colors[0].r, 255U);
  ASSERT_EQ(out.icon_set_threshold_count, 2U);
  ASSERT_NE(out.icon_set_thresholds, nullptr);
  EXPECT_STREQ(out.icon_set_thresholds[1].value, "80");
}
