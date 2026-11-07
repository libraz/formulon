// Workbook::ensure_calculated: reads recalculate a dirty workbook in
// Automatic mode and leave it alone in Manual mode.

#include <cstdint>

#include "calc_settings.h"
#include "cell.h"
#include "data_table.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "merge_range.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

Value CachedAt(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(0).cell_at(row, col);
  return cell != nullptr ? cell->cached_value : Value::blank();
}

bool IsDirty(const Workbook& wb) {
  return wb.recalc_engine().dirty().size() != 0U;
}

// A1 = 2, A2 = A1*10, left uncalculated.
Workbook PendingProduct(CalcMode mode) {
  Workbook wb = Workbook::create();
  wb.set_calc_mode(mode);
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(2.0))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1*10")));
  return wb;
}

TEST(WorkbookAutoCalc, AutomaticModeComputesPendingCells) {
  Workbook wb = PendingProduct(CalcMode::kAuto);
  ASSERT_TRUE(IsDirty(wb));
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  const Value a2 = CachedAt(wb, 1U, 0U);
  ASSERT_TRUE(a2.is_number());
  EXPECT_EQ(a2.as_number(), 20.0);
  EXPECT_FALSE(IsDirty(wb));
}

TEST(WorkbookAutoCalc, ManualModeLeavesPendingCellsUntouched) {
  Workbook wb = PendingProduct(CalcMode::kManual);
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  EXPECT_TRUE(CachedAt(wb, 1U, 0U).is_blank());
  EXPECT_TRUE(IsDirty(wb));
  // An explicit recalc still computes them.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(CachedAt(wb, 1U, 0U).as_number(), 20.0);
}

TEST(WorkbookAutoCalc, CleanWorkbookIsNotRecalculated) {
  Workbook wb = PendingProduct(CalcMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  // A sentinel outside the engine's view survives only if no pass runs.
  wb.sheet(0).set_cell_cached_value(1U, 0U, Value::number(-1.0));
  ASSERT_FALSE(IsDirty(wb));
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  EXPECT_EQ(CachedAt(wb, 1U, 0U).as_number(), -1.0);
}

TEST(WorkbookAutoCalc, EditAfterCalculationIsPickedUpByTheNextRead) {
  Workbook wb = PendingProduct(CalcMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(3.0))));
  ASSERT_TRUE(IsDirty(wb));
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  EXPECT_EQ(CachedAt(wb, 1U, 0U).as_number(), 30.0);
}

TEST(WorkbookAutoCalc, VolatileFormulasAreNotRecomputedPerRead) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=RAND()")));
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  const Value first = CachedAt(wb, 0U, 0U);
  ASSERT_TRUE(first.is_number());
  EXPECT_FALSE(IsDirty(wb));
  for (int read = 0; read < 3; ++read) {
    ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
    EXPECT_EQ(CachedAt(wb, 0U, 0U).as_number(), first.as_number()) << "read " << read;
  }
  // An edit elsewhere triggers a pass, and the pass re-draws the volatile.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 5U, 5U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  EXPECT_NE(CachedAt(wb, 0U, 0U).as_number(), first.as_number());
}

// D1 = A1*2 heads a one-variable table D2:D4 over input A1, with C2:C4 = 10, 20, 30.
Workbook PendingDataTable(CalcMode mode) {
  Workbook wb = Workbook::create();
  wb.set_calc_mode(mode);
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=A1*2")));
  for (std::uint32_t k = 0; k < 3U; ++k) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U + k, 2U, Value::number(10.0 * (k + 1U)))));
  }
  DataTableInput input;
  input.present = true;
  input.row = 0U;
  input.col = 0U;
  EXPECT_TRUE(static_cast<bool>(wb.add_data_table(0U, MergeRange{1U, 3U, 3U, 3U}, DataTableInput{}, input)));
  return wb;
}

TEST(WorkbookAutoCalc, AutomaticModeComputesDataTables) {
  Workbook wb = PendingDataTable(CalcMode::kAuto);
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  EXPECT_EQ(CachedAt(wb, 0U, 3U).as_number(), 2.0);
  for (std::uint32_t k = 0; k < 3U; ++k) {
    EXPECT_EQ(CachedAt(wb, 1U + k, 3U).as_number(), 20.0 * (k + 1U)) << "row " << k;
  }
}

TEST(WorkbookAutoCalc, AutoNoTableComputesCellsButSkipsDataTables) {
  Workbook wb = PendingDataTable(CalcMode::kAutoNoTable);
  ASSERT_TRUE(static_cast<bool>(wb.ensure_calculated(eval::default_registry())));
  EXPECT_EQ(CachedAt(wb, 0U, 3U).as_number(), 2.0);
  EXPECT_FALSE(IsDirty(wb));
  for (std::uint32_t k = 0; k < 3U; ++k) {
    const Value result = CachedAt(wb, 1U + k, 3U);
    ASSERT_TRUE(result.is_number()) << "row " << k;
    EXPECT_EQ(result.as_number(), 0.0) << "row " << k;
  }
  // The tables are the host's to compute on demand.
  ASSERT_TRUE(static_cast<bool>(wb.recalc_data_tables(eval::default_registry())));
  EXPECT_EQ(CachedAt(wb, 1U, 3U).as_number(), 20.0);
}

// Reads issued from a progress callback, and what they returned.
struct ReentryProbe {
  Workbook* self = nullptr;
  Workbook* other = nullptr;
  Workbook* manual = nullptr;
  bool called = false;
  FormulonErrorCode self_code = FormulonErrorCode::kOk;
  FormulonErrorCode other_code = FormulonErrorCode::kOk;
  bool manual_ok = false;
};

TEST(WorkbookAutoCalc, ReadFromInsideARecalcFailsAndKeepsTheDirtySet) {
  // A1 = B1+1, B1 = A1/2 runs the iterative solver, whose progress callback reads.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=B1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1/2")));
  IterativeOptions options;
  options.enabled = true;
  options.max_iterations = 2U;
  wb.set_iterative_options(options);

  Workbook other = PendingProduct(CalcMode::kAuto);
  Workbook manual = PendingProduct(CalcMode::kManual);
  ReentryProbe probe;
  probe.self = &wb;
  probe.other = &other;
  probe.manual = &manual;

  wb.set_iterative_progress(
      [](std::uint32_t, double, std::uint32_t, void* data) {
        auto& p = *static_cast<ReentryProbe*>(data);
        if (p.called) {
          return true;
        }
        p.called = true;
        auto self = p.self->ensure_calculated(eval::default_registry());
        if (!self) {
          p.self_code = self.error().code;
        }
        auto other_read = p.other->ensure_calculated(eval::default_registry());
        if (!other_read) {
          p.other_code = other_read.error().code;
        }
        p.manual_ok = static_cast<bool>(p.manual->ensure_calculated(eval::default_registry()));
        return true;
      },
      &probe);
  const auto outer = wb.ensure_calculated(eval::default_registry());
  wb.set_iterative_progress(nullptr, nullptr);

  ASSERT_TRUE(static_cast<bool>(outer));
  ASSERT_TRUE(probe.called);
  EXPECT_EQ(probe.self_code, FormulonErrorCode::kGraphRecalcReentrant);
  EXPECT_EQ(probe.other_code, FormulonErrorCode::kGraphRecalcReentrant);
  EXPECT_TRUE(probe.manual_ok);
  // The failed read left the other workbook pending; a later read computes it.
  EXPECT_TRUE(IsDirty(other));
  EXPECT_TRUE(CachedAt(other, 1U, 0U).is_blank());
  ASSERT_TRUE(static_cast<bool>(other.ensure_calculated(eval::default_registry())));
  EXPECT_EQ(CachedAt(other, 1U, 0U).as_number(), 20.0);
}

}  // namespace
}  // namespace formulon
