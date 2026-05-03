// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stable C ABI conditional-format end-to-end tests.
//
// Drives the `fm_workbook_cf_evaluate_range` / `fm_cf_results_*` surface
// declared in `c_api/formulon_c.h` through the same opaque-handle
// pattern the rest of the C ABI suite uses. CF blocks are seeded by
// constructing a `formulon::Workbook`, populating
// `Sheet::mutable_conditional_formats()`, serialising via
// `Workbook::save()`, and then loading the bytes through
// `fm_workbook_load` — the C ABI does not expose a CF mutator, so we
// rely on the OOXML round-trip (covered separately in
// `tests/integration/ooxml_cf_test.cpp`) to deliver a workbook handle
// with the rules attached.

#include <cmath>
#include <cstdint>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
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

// Builds a `Workbook`, applies `mutate` to seed cells / CF blocks, then
// serialises it through OOXML and loads the resulting bytes through the
// C ABI. The returned guard owns a populated `fm_workbook_t*` ready for
// CF evaluation.
template <typename MutateFn>
WorkbookGuard WorkbookFromMutator(MutateFn&& mutate) {
  formulon::Workbook wb = formulon::Workbook::create();
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
  fm_cf_results_t* out = nullptr;
  fm_status_t rc = fm_workbook_cf_evaluate_range(nullptr, 0, 0, 0, 0, 0, std::nan(""), &out);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(out, nullptr);
}

TEST(FormulonCApiCf, OutOfRangeSheetIndexReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cf_results_t* out = nullptr;
  fm_status_t rc = fm_workbook_cf_evaluate_range(wb.handle, 99, 0, 0, 0, 0, std::nan(""), &out);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(out, nullptr);
}
