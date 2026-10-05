// Stable C ABI conditional-format validation tests.

#include <utility>

#include "formulon_c_cf_test_helpers.h"

TEST(FormulonCApiCfMutate, OutOfGridSqrefRejectedAndLeavesTheModelUnchanged) {
  // An out-of-grid rectangle has no A1 spelling, so letting one into the
  // model leaves the save path with a reference it cannot write.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  rule.formula1 = "TRUE";
  rule.sqref_count = 1;
  std::size_t rule_index = 0;

  // One past the last column, on a whole-column range.
  fm_cf_cell_range_t past_last_col{0, formulon::Sheet::kMaxCols, formulon::Sheet::kMaxRows - 1U,
                                   formulon::Sheet::kMaxCols};
  rule.sqref = &past_last_col;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // One past the last row.
  fm_cf_cell_range_t past_last_row{0, 0, formulon::Sheet::kMaxRows, 0};
  rule.sqref = &past_last_row;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // Inverted rectangle.
  fm_cf_cell_range_t inverted{5, 0, 1, 0};
  rule.sqref = &inverted;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // A rejected range leaves no block behind, and the rejection is
  // reported rather than aborting the process at save time.
  EXPECT_TRUE(wb.handle->workbook().sheet(0).conditional_formats().empty());
  BufferGuard rejected_save;
  ASSERT_EQ(fm_workbook_save(wb.handle, &rejected_save.data, &rejected_save.len), 0);

  // The widest legal whole-column range still round-trips.
  fm_cf_cell_range_t whole_col{0, 0, formulon::Sheet::kMaxRows - 1U, 0};
  rule.sqref = &whole_col;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  const auto& blocks = reloaded.handle->workbook().sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 1U);
  ASSERT_EQ(blocks[0].sqref.size(), 1U);
  EXPECT_EQ(blocks[0].sqref[0], MakeRange(0, 0, formulon::Sheet::kMaxRows - 1U, 0));
}

TEST(FormulonCApiCfMutate, EmptySqrefRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_rule_t rule{};
  rule.type = 0;
  rule.sqref = nullptr;
  rule.sqref_count = 0;
  std::size_t rule_index = 0;
  fm_status_t rc = fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiCfMutate, AddRuleRejectsExcessiveSqrefCount) {
  // Hostile caller passes the maximum unsigned 32-bit value as
  // `sqref_count`. The `sqref` pointer is non-null so the early
  // null-check does not short-circuit, but the binding's range-count
  // cap must reject the call before the body attempts a 4 GiB
  // `reserve()`. Sheet state must be unchanged on rejection.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::size_t before = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &before), 0);

  fm_cf_cell_range_t sqref{0, 0, 0, 0};
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  std::string f = "TRUE";
  rule.formula1 = f.c_str();
  rule.sqref = &sqref;
  rule.sqref_count = 0xFFFFFFFFu;
  std::size_t rule_index = 0;
  fm_status_t rc = fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  std::size_t after = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &after), 0);
  EXPECT_EQ(before, after);
}

TEST(FormulonCApiCfMutate, GetAtDisengagesADxfIdTheStylesTableCannotResolve) {
  // A package can carry a `dxfId` past the end of its own `<dxfs>`, and
  // the loader keeps the rule. Surfacing that index would hand the caller
  // a value `fm_styles_get_dxf` rejects and `fm_sheet_cf_add_rule`
  // refuses, so the get / remove / add-back edit path would not close.
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
  rule.dxf_id = 1;
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();

  // Shrink the table under the rule, the way a package whose styles part
  // and worksheet part disagree arrives.
  wb.handle->workbook().mutable_styles().dxfs.resize(1);

  fm_cf_rule_t out{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out), 0) << fm_last_error_message();
  EXPECT_EQ(out.dxf_id_engaged, 0);

  // The record still feeds straight back in, which is what the disengage
  // buys: with the raw index it would be rejected.
  std::size_t reinserted = 0;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, out, &reinserted), 0) << fm_last_error_message();

  // A resolvable index is still reported.
  SeedDxfs(wb.handle, 2);
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.dxf_id_engaged, 1);
  EXPECT_EQ(out.dxf_id, 1U);
}

TEST(FormulonCApiCfMutate, AddRuleRejectsNonGuidId) {
  // The id becomes an `ST_Guid`-typed attribute in the saved package, so
  // a host key like "sales-databar" would make Excel repair the workbook
  // and drop every conditional format in it.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  rule.formula1 = "TRUE";
  rule.sqref = &sqref;
  rule.sqref_count = 1;

  std::size_t rule_index = 99U;
  for (const char* rejected : {"sales-databar", "{}", "FC000000-0000-0000-0000-000000000001",
                               "{FC000000-0000-0000-0000-00000000000}", "{GC000000-0000-0000-0000-000000000001}"}) {
    rule.id = rejected;
    EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << rejected;
  }

  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  // A GUID-shaped id is accepted and comes back verbatim.
  rule.id = "{3B4F1C22-0000-4000-8000-0123456789AB}";
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();
  fm_cf_rule_t out{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.id, "{3B4F1C22-0000-4000-8000-0123456789AB}");

  // An absent id is the "synthesize one for me" request, not a malformed
  // one, and the id it produces is itself GUID-shaped.
  rule.id = "";
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, rule_index, &out), 0);
  ASSERT_NE(out.id, nullptr);
  const std::string synthesized(out.id);
  EXPECT_EQ(synthesized.size(), 38U);
  EXPECT_EQ(synthesized.front(), '{');
  EXPECT_EQ(synthesized.back(), '}');
}

TEST(FormulonCApiCfMutate, AddRuleRejectsEnumOrdinalsPastTheirDomain) {
  // Each of these would otherwise be `static_cast` into its enum and
  // collapse onto the consuming switch's default: an unknown `op` writes
  // `operator="equal"`, an unknown `time_period` writes `today`. The
  // caller sees `kOk` and a rule that means something else.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t base{};
  base.formula1 = "TRUE";
  base.sqref = &sqref;
  base.sqref_count = 1;

  const auto expect_rejected = [&wb](fm_cf_rule_t rule, const char* what) {
    std::size_t rule_index = 99U;
    EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << what;
  };

  fm_cf_rule_t unknown_type = base;
  unknown_type.type = static_cast<std::uint8_t>(formulon::cf::RuleType::UniqueValues) + 1U;
  expect_rejected(unknown_type, "type");

  fm_cf_rule_t unknown_op = base;
  unknown_op.type = 1;  // CellIs
  unknown_op.op_engaged = 1;
  unknown_op.op = static_cast<std::uint8_t>(formulon::cf::CellIsOperator::NotBetween) + 1U;
  expect_rejected(unknown_op, "op");

  fm_cf_rule_t unknown_period = base;
  unknown_period.type = 15;  // TimePeriod
  unknown_period.time_period_engaged = 1;
  unknown_period.time_period = static_cast<std::uint8_t>(formulon::cf::TimePeriod::NextMonth) + 1U;
  expect_rejected(unknown_period, "time_period");

  fm_cfvo_t icon_thresholds[3]{};
  icon_thresholds[1].type = static_cast<std::uint8_t>(formulon::cf::CfvoType::AutoMax) + 1U;
  fm_cf_rule_t unknown_cfvo = base;
  unknown_cfvo.type = 4;  // IconSet
  unknown_cfvo.icon_set_engaged = 1;
  unknown_cfvo.icon_set_thresholds = icon_thresholds;
  unknown_cfvo.icon_set_threshold_count = 3;
  expect_rejected(unknown_cfvo, "icon_set_thresholds[].type");

  fm_cfvo_t legal_thresholds[3]{};
  fm_cf_rule_t unknown_icon_set = base;
  unknown_icon_set.type = 4;  // IconSet
  unknown_icon_set.icon_set_engaged = 1;
  unknown_icon_set.icon_set_name = static_cast<std::uint8_t>(formulon::cf::IconSetName::Five_Quarters) + 1U;
  unknown_icon_set.icon_set_thresholds = legal_thresholds;
  unknown_icon_set.icon_set_threshold_count = 3;
  expect_rejected(unknown_icon_set, "icon_set_name");

  fm_cfvo_t color_thresholds[2]{};
  fm_cf_color_t colors[2]{};
  color_thresholds[0].type = static_cast<std::uint8_t>(formulon::cf::CfvoType::AutoMax) + 1U;
  fm_cf_rule_t unknown_color_cfvo = base;
  unknown_color_cfvo.type = 2;  // ColorScale
  unknown_color_cfvo.color_scale_thresholds = color_thresholds;
  unknown_color_cfvo.color_scale_colors = colors;
  unknown_color_cfvo.color_scale_count = 2;
  expect_rejected(unknown_color_cfvo, "color_scale_thresholds[].type");

  fm_cf_rule_t unknown_floor = base;
  unknown_floor.type = 4;  // IconSet
  unknown_floor.icon_set_engaged = 1;
  unknown_floor.icon_set_thresholds = legal_thresholds;
  unknown_floor.icon_set_threshold_count = 2;
  unknown_floor.icon_set_floor_engaged = 1;
  unknown_floor.icon_set_floor.type = static_cast<std::uint8_t>(formulon::cf::CfvoType::AutoMax) + 1U;
  expect_rejected(unknown_floor, "icon_set_floor.type");

  fm_cf_rule_t unknown_direction = base;
  unknown_direction.type = 3;  // DataBar
  unknown_direction.data_bar_engaged = 1;
  unknown_direction.data_bar_max.type = 4;  // Max
  unknown_direction.data_bar_min.type = 3;  // Min
  unknown_direction.data_bar_min_length_pct = 10;
  unknown_direction.data_bar_max_length_pct = 90;
  unknown_direction.data_bar_direction = static_cast<std::uint8_t>(formulon::cf::DataBarDirection::RightToLeft) + 1U;
  expect_rejected(unknown_direction, "data_bar_direction");

  // Every rejection left the sheet as it was.
  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiCfMutate, OutOfRangeIndexReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_rule_t out{};
  fm_status_t rc = fm_sheet_cf_get_at(wb.handle, 0, 99, &out);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  rc = fm_sheet_cf_remove_at(wb.handle, 0, 99);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiCfMutate, PreLoadedRulesEnumerableViaFlatIndex) {
  WorkbookGuard wb = WorkbookFromMutator([](formulon::Workbook& w) {
    auto& sheet = w.sheet(0);
    formulon::cf::ConditionalFormat block;
    block.sqref = {MakeRange(0, 0, 9, 0)};
    formulon::cf::CFRule r;
    r.type = formulon::cf::RuleType::CellIs;
    r.priority = 1;
    r.op = formulon::cf::CellIsOperator::GreaterThan;
    r.formula1 = "50";
    r.dxf_id = 0;
    block.rules.push_back(std::move(r));
    sheet.mutable_conditional_formats().push_back(std::move(block));
  });
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  fm_cf_rule_t out{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.type, 1U);
  ASSERT_NE(out.formula1, nullptr);
  EXPECT_STREQ(out.formula1, "50");
}

TEST(FormulonCApiCfMutate, DuplicatePredicateRulesGetDistinctStableIndices) {
  // Two rules with the identical predicate (same type/op/formula/sqref)
  // must still be tracked as distinct entries: `fm_sheet_cf_add_rule`
  // always appends a new block rather than deduping against an existing
  // one, so the returned index must reflect append order, not predicate
  // identity.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDxfs(wb.handle, 1);

  fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 1;  // CellIs
  rule.op_engaged = 1;
  rule.op = 5;  // GreaterThan
  rule.formula1 = "50";
  rule.dxf_id_engaged = 1;
  rule.dxf_id = 0;
  rule.sqref = &sqref;
  rule.sqref_count = 1;

  std::size_t first_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &first_index), 0) << fm_last_error_message();

  // Re-populate the rule: `fm_sheet_cf_add_rule` deep-copies string /
  // range payloads, but reuse a fresh `fm_cf_rule_t` to avoid relying on
  // stale scratch state from the first call.
  fm_cf_rule_t rule2{};
  rule2.type = 1;  // CellIs
  rule2.op_engaged = 1;
  rule2.op = 5;  // GreaterThan
  rule2.formula1 = "50";
  rule2.dxf_id_engaged = 1;
  rule2.dxf_id = 0;
  rule2.sqref = &sqref;
  rule2.sqref_count = 1;

  std::size_t second_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule2, &second_index), 0) << fm_last_error_message();

  EXPECT_EQ(first_index, 0U);
  EXPECT_EQ(second_index, 1U);
  EXPECT_NE(first_index, second_index);

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  // The indices returned by add_rule must match what a subsequent
  // flattened readback reports for rule order/position.
  fm_cf_rule_t readback_first{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, first_index, &readback_first), 0);
  EXPECT_EQ(readback_first.priority, 1);  // auto-assigned to the first rule added

  fm_cf_rule_t readback_second{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, second_index, &readback_second), 0);
  EXPECT_EQ(readback_second.priority, 2);  // auto-assigned one past the first
}
