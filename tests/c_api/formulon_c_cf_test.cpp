//
// Stable C ABI conditional-format end-to-end tests.
//
// Drives the `fm_workbook_cf_evaluate_range` / `fm_cf_results_*` surface
// declared in `c_api/formulon_c.h` through the same opaque-handle
// pattern the rest of the C ABI suite uses. Evaluation tests seed CF
// blocks through OOXML round-trip; mutation tests below drive the
// public `fm_sheet_cf_*` add/remove/clear surface directly.

#include <cmath>
#include <cstdint>
#include <initializer_list>
#include <string>
#include <utility>
#include <vector>

#include "c_api/borrowed_arena.h"
#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace {

// RAII guard so the workbook handle is released even on test failure.
// Move-only so factory functions can return one by value.
struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
  WorkbookGuard(WorkbookGuard&& other) noexcept : handle(other.handle) { other.handle = nullptr; }
  WorkbookGuard& operator=(WorkbookGuard&& other) noexcept {
    if (this != &other) {
      fm_workbook_destroy(handle);
      handle = other.handle;
      other.handle = nullptr;
    }
    return *this;
  }
};

// RAII guard for `fm_cf_results_t` handles.
struct CfResultsGuard {
  fm_cf_results_t* handle = nullptr;
  ~CfResultsGuard() { fm_cf_results_destroy(handle); }
  CfResultsGuard() = default;
  CfResultsGuard(const CfResultsGuard&) = delete;
  CfResultsGuard& operator=(const CfResultsGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

// Builds a `Workbook`, applies `mutate` to seed cells / CF blocks, then
// serialises it through OOXML and loads the resulting bytes through the
// C ABI. The returned guard owns a populated `fm_workbook_t*` ready for
// CF evaluation.
template <typename MutateFn>
WorkbookGuard WorkbookFromMutator(MutateFn&& mutate) {
  formulon::Workbook wb = formulon::Workbook::create();
  // Registers the `<dxf>` records the CF rules below reference. The
  // writer omits any `dxfId` its styles part cannot resolve, so a rule
  // whose differential format was never registered would come back from
  // the load with no `dxf_id` at all.
  wb.mutable_styles().dxfs.resize(8);
  mutate(wb);
  auto bytes = wb.save();
  EXPECT_TRUE(static_cast<bool>(bytes)) << "Workbook::save: " << (bytes ? "" : bytes.error().message);
  WorkbookGuard guard;
  if (!bytes) {
    return guard;
  }
  const auto& src = bytes.value();
  EXPECT_EQ(fm_workbook_load(src.data(), src.size(), &guard.handle), 0)
      << "fm_workbook_load: " << fm_last_error_message();
  return guard;
}

// Registers `count` `<dxf>` slots on an existing handle. `fm_sheet_cf_add_rule`
// refuses a `dxf_id` past the end of the table, so a mutation test that
// engages a differential format has to seed the table first.
void SeedDxfs(fm_workbook_t* handle, std::size_t count) {
  handle->workbook().mutable_styles().dxfs.resize(count);
}

formulon::cf::CFCellRange MakeRange(std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2) {
  return {{r1, c1}, {r2, c2}};
}

}  // namespace

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

// ---------------------------------------------------------------------------
// CF mutation API (fm_sheet_cf_count / get_at / add_rule / remove_at / clear)
// ---------------------------------------------------------------------------

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

namespace {

// Minimal valid DataBar rule. Extension fields stay zero-initialized so
// each test below engages only what it is about to assert.
fm_cf_rule_t MakeDataBarRule(const fm_cf_cell_range_t* sqref) {
  fm_cf_rule_t rule{};
  rule.type = 3;  // DataBar
  rule.sqref = sqref;
  rule.sqref_count = 1;
  rule.data_bar_engaged = 1;
  rule.data_bar_min.type = 3;  // Min
  rule.data_bar_min.gte = 1;
  rule.data_bar_max.type = 4;  // Max
  rule.data_bar_max.gte = 1;
  rule.data_bar_fill = fm_cf_color_t{99, 142, 198, 255};
  rule.data_bar_show_value = 1;
  rule.data_bar_min_length_pct = 10;
  rule.data_bar_max_length_pct = 90;
  return rule;
}

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
}
