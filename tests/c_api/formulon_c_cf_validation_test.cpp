// Stable C ABI conditional-format validation tests.

#include <utility>

#include "formulon_c_cf_test_helpers.h"

TEST(FormulonCApiCfMutate, OutOfGridSqrefRejectedAndLeavesTheModelUnchanged) {
  // An out-of-grid rectangle has no A1 spelling, so letting one into the
  // model leaves the save path with a reference it cannot write.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;  // Expression
  rule.formula1 = "TRUE";
  rule.range_count = 1;
  const char* rule_id = nullptr;

  // One past the last column, on a whole-column range.
  fm_cell_rect_t past_last_col{0, formulon::Sheet::kMaxCols, formulon::Sheet::kMaxRows - 1U, formulon::Sheet::kMaxCols};
  rule.ranges = &past_last_col;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // One past the last row.
  fm_cell_rect_t past_last_row{0, 0, formulon::Sheet::kMaxRows, 0};
  rule.ranges = &past_last_row;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // Inverted rectangle.
  fm_cell_rect_t inverted{5, 0, 1, 0};
  rule.ranges = &inverted;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // A rejected range leaves no block behind, and the rejection is
  // reported rather than aborting the process at save time.
  EXPECT_TRUE(wb.handle->workbook().sheet(0).conditional_formats().empty());
  BufferGuard rejected_save;
  ASSERT_EQ(fm_workbook_save(wb.handle, &rejected_save.data, &rejected_save.len), 0);

  // The widest legal whole-column range still round-trips.
  fm_cell_rect_t whole_col{0, 0, formulon::Sheet::kMaxRows - 1U, 0};
  rule.ranges = &whole_col;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id), 0) << fm_get_last_error_message();

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
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;
  rule.ranges = nullptr;
  rule.range_count = 0;
  const char* rule_id = nullptr;
  fm_status_t rc = fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiCfMutate, AddRuleRejectsExcessiveSqrefCount) {
  // Hostile caller passes the maximum unsigned 32-bit value as
  // `range_count`. The `ranges` pointer is non-null so the early
  // null-check does not short-circuit, but the binding's range-count
  // cap must reject the call before the body attempts a 4 GiB
  // `reserve()`. Sheet state must be unchanged on rejection.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t before = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &before), 0);

  fm_cell_rect_t sqref{0, 0, 0, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;  // Expression
  std::string f = "TRUE";
  rule.formula1 = f.c_str();
  rule.ranges = &sqref;
  rule.range_count = 0xFFFFFFFFu;
  const char* rule_id = nullptr;
  fm_status_t rc = fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  std::size_t after = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &after), 0);
  EXPECT_EQ(before, after);
}

TEST(FormulonCApiCfMutate, GetAtDisengagesADxfIdTheStylesTableCannotResolve) {
  // A package can carry a `dxfId` past the end of its own `<dxfs>`, and
  // the loader keeps the rule. Surfacing that index would hand the caller
  // a value `fm_workbook_get_dxf_at` rejects and `fm_sheet_add_cf_rule`
  // refuses, so the get / remove / add-back edit path would not close.
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
  rule.dxf_index = 1;
  rule.ranges = &sqref;
  rule.range_count = 1;
  const char* rule_id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id), 0) << fm_get_last_error_message();

  // Shrink the table under the rule, the way a package whose styles part
  // and worksheet part disagree arrives.
  wb.handle->workbook().mutable_styles().dxfs.resize(1);

  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(out.has_dxf_index, 0);

  // The record still feeds straight back in, which is what the disengage
  // buys: with the raw index it would be rejected. It goes in under a fresh
  // id beside the original, whose id it would otherwise repeat.
  out.id = nullptr;
  const char* reinserted = nullptr;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &out, &reinserted), 0) << fm_get_last_error_message();

  // A resolvable index is still reported.
  SeedDxfs(wb.handle, 2);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.has_dxf_index, 1);
  EXPECT_EQ(out.dxf_index, 1U);
}

TEST(FormulonCApiCfMutate, AddRuleRejectsNonGuidId) {
  // The id becomes an `ST_Guid`-typed attribute in the saved package, so
  // a host key like "sales-databar" would make Excel repair the workbook
  // and drop every conditional format in it.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;  // Expression
  rule.formula1 = "TRUE";
  rule.ranges = &sqref;
  rule.range_count = 1;

  const char* rule_id = nullptr;
  for (const char* rejected : {"sales-databar", "{}", "FC000000-0000-0000-0000-000000000001",
                               "{FC000000-0000-0000-0000-00000000000}", "{GC000000-0000-0000-0000-000000000001}"}) {
    rule.id = rejected;
    EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << rejected;
  }

  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  // A GUID-shaped id is accepted and comes back verbatim.
  rule.id = "{3B4F1C22-0000-4000-8000-0123456789AB}";
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id), 0) << fm_get_last_error_message();
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.id, "{3B4F1C22-0000-4000-8000-0123456789AB}");

  // An absent id is the "synthesize one for me" request, not a malformed
  // one, and the id it produces is itself GUID-shaped.
  rule.id = "";
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id), 0) << fm_get_last_error_message();
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, rule_id, &present, &out), 0);
  ASSERT_EQ(present, 1);
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
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t base = FM_STRUCT_INIT(fm_cf_rule_t);
  base.formula1 = "TRUE";
  base.ranges = &sqref;
  base.range_count = 1;

  const auto expect_rejected = [&wb](fm_cf_rule_t rule, const char* what) {
    const char* rule_id = nullptr;
    EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &rule_id),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << what;
  };

  fm_cf_rule_t unknown_type = base;
  unknown_type.type = static_cast<std::uint8_t>(formulon::cf::RuleType::UniqueValues) + 1U;
  expect_rejected(unknown_type, "type");

  fm_cf_rule_t unknown_op = base;
  unknown_op.type = 1;  // CellIs
  unknown_op.has_op = 1;
  unknown_op.op = static_cast<std::uint8_t>(formulon::cf::CellIsOperator::NotBetween) + 1U;
  expect_rejected(unknown_op, "op");

  fm_cf_rule_t unknown_period = base;
  unknown_period.type = 15;  // TimePeriod
  unknown_period.has_time_period = 1;
  unknown_period.time_period = static_cast<std::uint8_t>(formulon::cf::TimePeriod::NextMonth) + 1U;
  expect_rejected(unknown_period, "time_period");

  fm_cfvo_t icon_thresholds[3]{};
  icon_thresholds[1].type = static_cast<std::uint8_t>(formulon::cf::CfvoType::AutoMax) + 1U;
  fm_cf_rule_t unknown_cfvo = base;
  unknown_cfvo.type = 4;  // IconSet
  unknown_cfvo.has_icon_set = 1;
  unknown_cfvo.icon_set_thresholds = icon_thresholds;
  unknown_cfvo.icon_set_threshold_count = 3;
  expect_rejected(unknown_cfvo, "icon_set_thresholds[].type");

  fm_cfvo_t legal_thresholds[3]{};
  fm_cf_rule_t unknown_icon_set = base;
  unknown_icon_set.type = 4;  // IconSet
  unknown_icon_set.has_icon_set = 1;
  unknown_icon_set.icon_set_name = static_cast<std::uint8_t>(formulon::cf::IconSetName::Five_Quarters) + 1U;
  unknown_icon_set.icon_set_thresholds = legal_thresholds;
  unknown_icon_set.icon_set_threshold_count = 3;
  expect_rejected(unknown_icon_set, "icon_set_name");

  fm_cfvo_t color_thresholds[2]{};
  fm_color_t colors[2]{};
  color_thresholds[0].type = static_cast<std::uint8_t>(formulon::cf::CfvoType::AutoMax) + 1U;
  fm_cf_rule_t unknown_color_cfvo = base;
  unknown_color_cfvo.type = 2;  // ColorScale
  unknown_color_cfvo.color_scale_thresholds = color_thresholds;
  unknown_color_cfvo.color_scale_colors = colors;
  unknown_color_cfvo.color_scale_count = 2;
  expect_rejected(unknown_color_cfvo, "color_scale_thresholds[].type");

  fm_cf_rule_t unknown_floor = base;
  unknown_floor.type = 4;  // IconSet
  unknown_floor.has_icon_set = 1;
  unknown_floor.icon_set_thresholds = legal_thresholds;
  unknown_floor.icon_set_threshold_count = 2;
  unknown_floor.has_icon_set_floor = 1;
  unknown_floor.icon_set_floor.type = static_cast<std::uint8_t>(formulon::cf::CfvoType::AutoMax) + 1U;
  expect_rejected(unknown_floor, "icon_set_floor.type");

  fm_cf_rule_t unknown_direction = base;
  unknown_direction.type = 3;  // DataBar
  unknown_direction.has_data_bar = 1;
  unknown_direction.data_bar_max.type = 4;  // Max
  unknown_direction.data_bar_min.type = 3;  // Min
  unknown_direction.data_bar_min_length_pct = 10;
  unknown_direction.data_bar_max_length_pct = 90;
  unknown_direction.data_bar_direction = static_cast<std::uint8_t>(formulon::cf::DataBarDirection::RightToLeft) + 1U;
  expect_rejected(unknown_direction, "data_bar_direction");

  // Every rejection left the sheet as it was.
  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiCfMutate, OutOfRangePositionReadsIndexOutOfRangeAndUnknownIdRemovesNotFound) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  fm_status_t rc = fm_sheet_get_cf_rule_at(wb.handle, 0, 99, &out);
  EXPECT_EQ(rc, FM_STATUS_INDEX_OUT_OF_RANGE);
  rc = fm_sheet_remove_cf_rule(wb.handle, 0, "{00000000-0000-0000-0000-000000000000}");
  EXPECT_EQ(rc, FM_STATUS_NOT_FOUND);
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
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.type, 1U);
  ASSERT_NE(out.formula1, nullptr);
  EXPECT_STREQ(out.formula1, "=50");
}

TEST(FormulonCApiCfMutate, DuplicatePredicateRulesGetDistinctStableIndices) {
  // Two rules with the identical predicate (same type/op/formula/sqref)
  // must still be tracked as distinct entries: `fm_sheet_add_cf_rule`
  // always appends a new block rather than deduping against an existing
  // one, so the returned index must reflect append order, not predicate
  // identity.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  SeedDxfs(wb.handle, 1);

  fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 1;  // CellIs
  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_GREATER_THAN;
  rule.formula1 = "50";
  rule.has_dxf_index = 1;
  rule.dxf_index = 0;
  rule.ranges = &sqref;
  rule.range_count = 1;

  const char* first_ptr = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &first_ptr), 0) << fm_get_last_error_message();
  const std::string first_id = first_ptr;

  // Re-populate the rule: `fm_sheet_add_cf_rule` deep-copies string /
  // range payloads, but reuse a fresh `fm_cf_rule_t` to avoid relying on
  // stale scratch state from the first call.
  fm_cf_rule_t rule2 = FM_STRUCT_INIT(fm_cf_rule_t);
  rule2.type = 1;  // CellIs
  rule2.has_op = 1;
  rule2.op = FM_COMPARISON_OPERATOR_GREATER_THAN;
  rule2.formula1 = "50";
  rule2.has_dxf_index = 1;
  rule2.dxf_index = 0;
  rule2.ranges = &sqref;
  rule2.range_count = 1;

  const char* second_ptr = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule2, &second_ptr), 0) << fm_get_last_error_message();
  const std::string second_id = second_ptr;

  EXPECT_EQ(CfRuleIdAt(wb.handle, 0, 0), first_id);
  EXPECT_EQ(CfRuleIdAt(wb.handle, 0, 1), second_id);
  EXPECT_NE(first_id, second_id);

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  // The ids returned by add must name the rules a subsequent readback reports, in order.
  int32_t present = 0;
  fm_cf_rule_t readback_first = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, first_id.c_str(), &present, &readback_first), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(readback_first.priority, 1);  // auto-assigned to the first rule added

  fm_cf_rule_t readback_second = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, second_id.c_str(), &present, &readback_second), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(readback_second.priority, 2);  // auto-assigned one past the first
}
