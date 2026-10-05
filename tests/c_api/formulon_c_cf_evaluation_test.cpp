// Stable C ABI conditional-format evaluation and readback tests.

#include <cmath>
#include <tuple>
#include <utility>

#include "color_resolve.h"
#include "formulon_c_cf_test_helpers.h"

TEST(FormulonCApiCf, EvaluateRangeOnEmptyCfWorkbookReturnsZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, 9, 9, std::nan(""), &results.handle), 0);
  ASSERT_NE(results.handle, nullptr);
  EXPECT_EQ(fm_cf_results_cell_count(results.handle), 0U);
}

TEST(FormulonCApiCf, EvaluateRangeWithCellIsRule) {
  WorkbookGuard wb = WorkbookFromMutator([](formulon::Workbook& w) {
    auto& sheet = w.sheet(0);
    sheet.set_cell_value(0, 0, formulon::Value::number(10.0));
    sheet.set_cell_value(1, 0, formulon::Value::number(60.0));
    sheet.set_cell_value(2, 0, formulon::Value::number(90.0));

    formulon::cf::ConditionalFormat block{};
    block.sqref.push_back(MakeRange(0, 0, 2, 0));
    formulon::cf::CFRule rule;
    rule.type = formulon::cf::RuleType::CellIs;
    rule.priority = 1;
    rule.dxf_id = 7U;
    rule.op = formulon::cf::CellIsOperator::GreaterThan;
    rule.formula1 = "50";
    block.rules.push_back(std::move(rule));
    sheet.mutable_conditional_formats().push_back(std::move(block));
  });
  ASSERT_NE(wb.handle, nullptr);
  // Recalc so cached cell values are populated for the loaded handle;
  // CF evaluation reads the cached values via `EvalContext`.
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, 2, 0, std::nan(""), &results.handle), 0);
  ASSERT_NE(results.handle, nullptr);
  ASSERT_EQ(fm_cf_results_cell_count(results.handle), 2U);

  // First matched cell: A2 (row=1).
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::size_t match_count = 0;
  ASSERT_EQ(fm_cf_results_cell_at(results.handle, 0, &row, &col, &match_count), 0);
  EXPECT_EQ(row, 1U);
  EXPECT_EQ(col, 0U);
  ASSERT_EQ(match_count, 1U);
  fm_cf_match_t m{};
  ASSERT_EQ(fm_cf_results_match_at(results.handle, 0, 0, &m), 0);
  EXPECT_EQ(m.kind, FM_CF_DIFFERENTIAL_FORMAT);
  EXPECT_EQ(m.priority, 1);
  EXPECT_EQ(m.dxf_id_engaged, 1);
  EXPECT_EQ(m.dxf_id, 7U);

  // Second matched cell: A3 (row=2).
  ASSERT_EQ(fm_cf_results_cell_at(results.handle, 1, &row, &col, &match_count), 0);
  EXPECT_EQ(row, 2U);
  EXPECT_EQ(col, 0U);
  ASSERT_EQ(match_count, 1U);
  ASSERT_EQ(fm_cf_results_match_at(results.handle, 1, 0, &m), 0);
  EXPECT_EQ(m.kind, FM_CF_DIFFERENTIAL_FORMAT);
  EXPECT_EQ(m.dxf_id, 7U);
}

TEST(FormulonCApiCf, EvaluateRangeRejectsRequestPastViewportCeiling) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Two full columns are past the viewport ceiling the engine enforces;
  // the binding must surface the refusal instead of returning a handle.
  CfResultsGuard results;
  EXPECT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, formulon::cf::kCfMaxRows - 1U, 1, std::nan(""),
                                          &results.handle),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSecResourceLimit));
  EXPECT_EQ(results.handle, nullptr);
}

TEST(FormulonCApiCf, EvaluateWholeColumnRangeReturnsBlockMatches) {
  WorkbookGuard wb = WorkbookFromMutator([](formulon::Workbook& w) {
    auto& sheet = w.sheet(0);
    sheet.set_cell_value(0, 0, formulon::Value::number(10.0));
    sheet.set_cell_value(1, 0, formulon::Value::number(60.0));

    formulon::cf::ConditionalFormat block{};
    block.sqref.push_back(MakeRange(0, 0, 1, 0));
    formulon::cf::CFRule rule;
    rule.type = formulon::cf::RuleType::CellIs;
    rule.priority = 1;
    rule.dxf_id = 7U;
    rule.op = formulon::cf::CellIsOperator::GreaterThan;
    rule.formula1 = "50";
    block.rules.push_back(std::move(rule));
    sheet.mutable_conditional_formats().push_back(std::move(block));
  });
  ASSERT_NE(wb.handle, nullptr);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  // A whole column sits exactly at the ceiling and is answered from the
  // block, not from the column's million rows.
  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, formulon::cf::kCfMaxRows - 1U, 0, std::nan(""),
                                          &results.handle),
            0)
      << fm_last_error_message();
  ASSERT_EQ(fm_cf_results_cell_count(results.handle), 1U);

  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::size_t match_count = 0;
  ASSERT_EQ(fm_cf_results_cell_at(results.handle, 0, &row, &col, &match_count), 0);
  EXPECT_EQ(row, 1U);
  EXPECT_EQ(col, 0U);
  ASSERT_EQ(match_count, 1U);
  fm_cf_match_t m{};
  ASSERT_EQ(fm_cf_results_match_at(results.handle, 0, 0, &m), 0);
  EXPECT_EQ(m.dxf_id_engaged, 1);
  EXPECT_EQ(m.dxf_id, 7U);
}

TEST(FormulonCApiCf, EvaluateRangeWithColorScale) {
  WorkbookGuard wb = WorkbookFromMutator([](formulon::Workbook& w) {
    auto& sheet = w.sheet(0);
    sheet.set_cell_value(0, 0, formulon::Value::number(0.0));
    sheet.set_cell_value(1, 0, formulon::Value::number(50.0));
    sheet.set_cell_value(2, 0, formulon::Value::number(100.0));

    formulon::cf::ConditionalFormat block{};
    block.sqref.push_back(MakeRange(0, 0, 2, 0));

    formulon::cf::CFRule rule;
    rule.type = formulon::cf::RuleType::ColorScale;
    rule.priority = 1;

    formulon::cf::ColorScaleSpec spec;
    spec.thresholds.push_back({formulon::cf::CfvoType::Min, "", true});
    spec.thresholds.push_back({formulon::cf::CfvoType::Percentile, "50", true});
    spec.thresholds.push_back({formulon::cf::CfvoType::Max, "", true});
    spec.colors.push_back({255, 0, 0, 255});    // red
    spec.colors.push_back({255, 255, 0, 255});  // yellow
    spec.colors.push_back({0, 255, 0, 255});    // green
    rule.color_scale = std::move(spec);

    block.rules.push_back(std::move(rule));
    sheet.mutable_conditional_formats().push_back(std::move(block));
  });
  ASSERT_NE(wb.handle, nullptr);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, 2, 0, std::nan(""), &results.handle), 0);
  ASSERT_EQ(fm_cf_results_cell_count(results.handle), 3U);

  const std::vector<std::tuple<std::uint8_t, std::uint8_t, std::uint8_t>> expected = {
      {255, 0, 0},    // A1 = 0  → red endpoint
      {255, 255, 0},  // A2 = 50 → yellow midpoint
      {0, 255, 0},    // A3 = 100 → green endpoint
  };

  for (std::size_t i = 0; i < expected.size(); ++i) {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    std::size_t match_count = 0;
    ASSERT_EQ(fm_cf_results_cell_at(results.handle, i, &row, &col, &match_count), 0);
    ASSERT_EQ(match_count, 1U) << "i=" << i;
    fm_cf_match_t m{};
    ASSERT_EQ(fm_cf_results_match_at(results.handle, i, 0, &m), 0) << "i=" << i;
    EXPECT_EQ(m.kind, FM_CF_COLOR_SCALE) << "i=" << i;
    const auto& [er, eg, eb] = expected[i];
    EXPECT_NEAR(static_cast<int>(m.color.r), static_cast<int>(er), 1) << "i=" << i;
    EXPECT_NEAR(static_cast<int>(m.color.g), static_cast<int>(eg), 1) << "i=" << i;
    EXPECT_NEAR(static_cast<int>(m.color.b), static_cast<int>(eb), 1) << "i=" << i;
  }
}

TEST(FormulonCApiCf, DestroyHandlesNullSafely) {
  fm_cf_results_destroy(nullptr);
  // No assertion needed; we just exercise the null path for ASan.
  SUCCEED();
}

TEST(FormulonCApiCf, OutOfRangeIndicesReturnInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, 0, 0, std::nan(""), &results.handle), 0);
  ASSERT_EQ(fm_cf_results_cell_count(results.handle), 0U);

  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::size_t match_count = 0;
  fm_status_t rc = fm_cf_results_cell_at(results.handle, 0, &row, &col, &match_count);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  fm_cf_match_t m{};
  rc = fm_cf_results_match_at(results.handle, 0, 0, &m);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiCf, NullWorkbookSetsBindingError) {
  fm_cf_results_t* out = reinterpret_cast<fm_cf_results_t*>(static_cast<std::uintptr_t>(1));
  fm_status_t rc = fm_workbook_cf_evaluate_range(nullptr, 0, 0, 0, 0, 0, std::nan(""), &out);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(out, nullptr);
  fm_cf_results_destroy(out);
}

TEST(FormulonCApiCf, OutOfRangeSheetIndexReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_results_t* out = reinterpret_cast<fm_cf_results_t*>(static_cast<std::uintptr_t>(1));
  fm_status_t rc = fm_workbook_cf_evaluate_range(wb.handle, 99, 0, 0, 0, 0, std::nan(""), &out);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(out, nullptr);
  fm_cf_results_destroy(out);
}

TEST(FormulonCApiCf, OutOfGridOrReversedRectReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_results_t* out = reinterpret_cast<fm_cf_results_t*>(static_cast<std::uintptr_t>(1));
  // A reversed rectangle (last < first) would wrap the iteration span.
  EXPECT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 5, 5, 0, 0, std::nan(""), &out),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(out, nullptr);
  // A corner past the grid ceiling would materialize billions of cells.
  out = reinterpret_cast<fm_cf_results_t*>(static_cast<std::uintptr_t>(1));
  EXPECT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, formulon::Sheet::kMaxRows, 0, std::nan(""), &out),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(out, nullptr);
  fm_cf_results_destroy(out);
}

TEST(FormulonCApiCf, ExpressionRuleEvaluatesFunctionsAndQualifiedRefs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDxfs(wb.handle, 4);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 11.0), 0);  // Sheet1!A1
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 1, 0, 0, 5.0), 0);   // Sheet2!A1

  fm_cf_cell_range_t sqref{0, 0, 0, 0};
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  rule.formula1 = "AND(A1>10,Sheet2!A1=5)";
  rule.dxf_id_engaged = 1;
  rule.dxf_id = 3;
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();
  EXPECT_EQ(rule_index, 0U);

  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, 0, 0, std::nan(""), &results.handle), 0)
      << fm_last_error_message();
  ASSERT_EQ(fm_cf_results_cell_count(results.handle), 1U);

  fm_cf_match_t match{};
  ASSERT_EQ(fm_cf_results_match_at(results.handle, 0, 0, &match), 0);
  EXPECT_EQ(match.kind, FM_CF_DIFFERENTIAL_FORMAT);
  EXPECT_EQ(match.dxf_id_engaged, 1);
  EXPECT_EQ(match.dxf_id, 3U);
}

TEST(FormulonCApiCf, ExpressionRuleRecursivelyEvaluatesFormulaCells) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedDxfs(wb.handle, 5);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 1, 0, 0, 7.0), 0);                   // Sheet2!A1
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SUM(Sheet2!A1,5)"), 0);  // Sheet1!A1

  fm_cf_cell_range_t sqref{0, 0, 0, 0};
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  rule.formula1 = "A1>10";
  rule.dxf_id_engaged = 1;
  rule.dxf_id = 4;
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  std::size_t rule_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &rule_index), 0) << fm_last_error_message();
  EXPECT_EQ(rule_index, 0U);

  // No explicit recalc: CF evaluation should use a recursive EvalState
  // instead of reading the formula cell's stale blank cached value.
  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, 0, 0, std::nan(""), &results.handle), 0)
      << fm_last_error_message();
  ASSERT_EQ(fm_cf_results_cell_count(results.handle), 1U);

  fm_cf_match_t match{};
  ASSERT_EQ(fm_cf_results_match_at(results.handle, 0, 0, &match), 0);
  EXPECT_EQ(match.kind, FM_CF_DIFFERENTIAL_FORMAT);
  EXPECT_EQ(match.dxf_id_engaged, 1);
  EXPECT_EQ(match.dxf_id, 4U);
}

namespace {

formulon::cf::Color ThemeColor(std::uint32_t theme, double tint) {
  formulon::cf::Color c{};
  c.spec.kind = formulon::ColorSpec::Kind::kTheme;
  c.spec.theme = theme;
  c.spec.tint = tint;
  return c;
}

std::uint32_t ResolvedArgb(const formulon::Workbook& wb, const formulon::cf::Color& c, formulon::ColorContext context) {
  return formulon::resolve_color(wb, c.spec, context).argb;
}

std::uint32_t PackArgb(const fm_cf_color_t& c) {
  return (static_cast<std::uint32_t>(c.a) << 24U) | (static_cast<std::uint32_t>(c.r) << 16U) |
         (static_cast<std::uint32_t>(c.g) << 8U) | static_cast<std::uint32_t>(c.b);
}

WorkbookGuard ThemeColorWorkbook() {
  return WorkbookFromMutator([](formulon::Workbook& w) {
    auto& sheet = w.sheet(0);
    for (std::uint32_t r = 0; r < 3; ++r) {
      sheet.set_cell_value(r, 0, formulon::Value::number(r * 50.0));
      sheet.set_cell_value(r, 1, formulon::Value::number(r * 50.0));
    }
    formulon::cf::ConditionalFormat scale_block{};
    scale_block.sqref.push_back(MakeRange(0, 0, 2, 0));
    formulon::cf::CFRule scale;
    scale.type = formulon::cf::RuleType::ColorScale;
    scale.priority = 1;
    formulon::cf::ColorScaleSpec spec;
    spec.thresholds.push_back({formulon::cf::CfvoType::Min, "", true});
    spec.thresholds.push_back({formulon::cf::CfvoType::Max, "", true});
    spec.colors.push_back(ThemeColor(4, 0.39997558519241921));
    spec.colors.push_back(ThemeColor(5, -0.249977111117893));
    scale.color_scale = std::move(spec);
    scale_block.rules.push_back(std::move(scale));
    sheet.mutable_conditional_formats().push_back(std::move(scale_block));

    formulon::cf::ConditionalFormat bar_block{};
    bar_block.sqref.push_back(MakeRange(0, 1, 2, 1));
    formulon::cf::CFRule bar;
    bar.type = formulon::cf::RuleType::DataBar;
    bar.priority = 2;
    formulon::cf::DataBarSpec bar_spec;
    bar_spec.min = {formulon::cf::CfvoType::Min, "", true};
    bar_spec.max = {formulon::cf::CfvoType::Max, "", true};
    bar_spec.fill = ThemeColor(8, 0.59999389629810485);
    bar_spec.negative_fill = bar_spec.fill;
    bar.data_bar = std::move(bar_spec);
    bar_block.rules.push_back(std::move(bar));
    sheet.mutable_conditional_formats().push_back(std::move(bar_block));
  });
}

}  // namespace

TEST(FormulonCApiCf, ThemeColorScalePayloadIsResolvedNotBlack) {
  WorkbookGuard wb = ThemeColorWorkbook();
  ASSERT_NE(wb.handle, nullptr);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 0, 2, 0, std::nan(""), &results.handle), 0);
  ASSERT_EQ(fm_cf_results_cell_count(results.handle), 3U);
  const auto& block = wb.handle->workbook().sheet(0).conditional_formats()[0];
  const std::uint32_t low = ResolvedArgb(wb.handle->workbook(), block.rules[0].color_scale->colors[0],
                                         formulon::ColorContext::kFillForeground);
  const std::uint32_t high = ResolvedArgb(wb.handle->workbook(), block.rules[0].color_scale->colors[1],
                                          formulon::ColorContext::kFillForeground);
  ASSERT_NE(low & 0xFFFFFFU, 0U);

  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::size_t match_count = 0;
  fm_cf_match_t m{};
  ASSERT_EQ(fm_cf_results_cell_at(results.handle, 0, &row, &col, &match_count), 0);
  ASSERT_EQ(fm_cf_results_match_at(results.handle, 0, 0, &m), 0);
  EXPECT_EQ(PackArgb(m.color), low);
  ASSERT_EQ(fm_cf_results_cell_at(results.handle, 2, &row, &col, &match_count), 0);
  ASSERT_EQ(fm_cf_results_match_at(results.handle, 2, 0, &m), 0);
  EXPECT_EQ(PackArgb(m.color), high);
  EXPECT_NE(PackArgb(m.color) & 0xFFFFFFU, 0U);
}

TEST(FormulonCApiCf, ThemeColorDataBarEvaluatedPayloadIsResolvedNotBlack) {
  WorkbookGuard wb = ThemeColorWorkbook();
  ASSERT_NE(wb.handle, nullptr);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  CfResultsGuard results;
  ASSERT_EQ(fm_workbook_cf_evaluate_range(wb.handle, 0, 0, 1, 2, 1, std::nan(""), &results.handle), 0);
  ASSERT_GT(fm_cf_results_cell_count(results.handle), 0U);
  const auto& bar = *wb.handle->workbook().sheet(0).conditional_formats()[1].rules[0].data_bar;
  const std::uint32_t expected = ResolvedArgb(wb.handle->workbook(), bar.fill, formulon::ColorContext::kFillForeground);
  ASSERT_NE(expected & 0xFFFFFFU, 0U);

  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::size_t match_count = 0;
  fm_cf_match_t m{};
  ASSERT_EQ(fm_cf_results_cell_at(results.handle, 0, &row, &col, &match_count), 0);
  ASSERT_EQ(fm_cf_results_match_at(results.handle, 0, 0, &m), 0);
  EXPECT_EQ(m.kind, FM_CF_DATA_BAR);
  EXPECT_EQ(PackArgb(m.bar_fill), expected);
}

TEST(FormulonCApiCf, ThemeColorRuleReadbackReportsResolvedChannels) {
  WorkbookGuard wb = ThemeColorWorkbook();
  ASSERT_NE(wb.handle, nullptr);
  const auto& blocks = wb.handle->workbook().sheet(0).conditional_formats();
  const std::uint32_t scale_low = ResolvedArgb(wb.handle->workbook(), blocks[0].rules[0].color_scale->colors[0],
                                               formulon::ColorContext::kFillForeground);
  const std::uint32_t bar_fill =
      ResolvedArgb(wb.handle->workbook(), blocks[1].rules[0].data_bar->fill, formulon::ColorContext::kFillForeground);

  fm_cf_rule_t rule{};
  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &rule), 0) << fm_last_error_message();
  ASSERT_GE(rule.color_scale_count, 2U);
  EXPECT_EQ(PackArgb(rule.color_scale_colors[0]), scale_low);
  EXPECT_NE(PackArgb(rule.color_scale_colors[0]) & 0xFFFFFFU, 0U);

  ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 1, &rule), 0) << fm_last_error_message();
  EXPECT_EQ(PackArgb(rule.data_bar_fill), bar_fill);
  EXPECT_NE(PackArgb(rule.data_bar_fill) & 0xFFFFFFU, 0U);
}

TEST(FormulonCApiCf, ThemeColorSurvivesXlsbSaveAndLoad) {
  WorkbookGuard wb = ThemeColorWorkbook();
  ASSERT_NE(wb.handle, nullptr);
  BufferGuard xlsb;
  ASSERT_EQ(fm_workbook_save_as(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &xlsb.data, &xlsb.len), 0)
      << fm_last_error_message();
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(xlsb.data, xlsb.len, &loaded.handle), 0) << fm_last_error_message();

  const auto& blocks = loaded.handle->workbook().sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 2U);
  const auto& colors = blocks[0].rules[0].color_scale->colors;
  ASSERT_EQ(colors.size(), 2U);
  EXPECT_EQ(colors[0].spec.kind, formulon::ColorSpec::Kind::kTheme);
  EXPECT_EQ(colors[0].spec.theme, 4U);
  EXPECT_NEAR(colors[0].spec.tint, 0.39997558519241921, 1e-4);
  EXPECT_EQ(colors[1].spec.theme, 5U);
  EXPECT_NEAR(colors[1].spec.tint, -0.249977111117893, 1e-4);
  const auto& fill = blocks[1].rules[0].data_bar->fill;
  EXPECT_EQ(fill.spec.kind, formulon::ColorSpec::Kind::kTheme);
  EXPECT_EQ(fill.spec.theme, 8U);
}
