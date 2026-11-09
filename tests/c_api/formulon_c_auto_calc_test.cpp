// C ABI reads recalculate a dirty workbook in Automatic mode and leave it
// pending in Manual mode; saves follow `calcOnSave` in Manual mode.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_evaluate_top_left.h"
#include "gtest/gtest.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// A1 = 2, A2 = A1*10, left uncalculated, under calc mode `mode`.
void MakePending(WorkbookGuard* wb, int32_t mode) {
  ASSERT_EQ(fm_workbook_create_default(&wb->handle), 0);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb->handle, mode), 0);
  ASSERT_EQ(fm_sheet_set_number(wb->handle, 0, 0, 0, 2.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb->handle, 0, 1, 0, "=A1*10"), 0);
}

bool HasDirty(const fm_workbook_t* wb) {
  int32_t dirty = -1;
  EXPECT_EQ(fm_workbook_has_dirty_cells(wb, &dirty), 0);
  return dirty == 1;
}

fm_value_t ValueAt(const fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &v), 0);
  return v;
}

// One value-reading entry point, called on a workbook built by `MakePending`.
struct Reader {
  const char* name;
  std::function<fm_status_t(fm_workbook_t*)> read;
};

std::vector<Reader> Readers() {
  return {
      {"get_value",
       [](fm_workbook_t* wb) {
         fm_value_t v{};
         return fm_sheet_get_value(wb, 0, 1, 0, &v);
       }},
      {"get_cells_in_range",
       [](fm_workbook_t* wb) {
         const fm_cell_rect_t rect{0, 0, 1, 0};
         fm_cell_range_t* range = nullptr;
         const fm_status_t rc = fm_sheet_get_cells_in_range(wb, 0, &rect, 0, 0, &range);
         fm_cell_range_destroy(range);
         return rc;
       }},
      {"get_display_text",
       [](fm_workbook_t* wb) {
         const char* text = nullptr;
         int32_t status = 0;
         return fm_sheet_get_display_text(wb, 0, 1, 0, &text, &status);
       }},
      {"evaluate_formula",
       [](fm_workbook_t* wb) {
         fm_value_t v{};
         return evaluate_top_left(wb, 0, 5, 5, "=A2", &v);
       }},
      {"evaluate_formula_array",
       [](fm_workbook_t* wb) {
         fm_value_array_t* array = nullptr;
         const fm_status_t rc = fm_sheet_evaluate_formula_array(wb, 0, 5, 5, "=A2", &array);
         fm_value_array_destroy(array);
         return rc;
       }},
      {"evaluate_cf_formula",
       [](fm_workbook_t* wb) {
         fm_value_t v{};
         return fm_sheet_evaluate_cf_formula(wb, 0, 1, 0, 1, 0, "=A2=20", &v);
       }},
      {"evaluate_cf_range",
       [](fm_workbook_t* wb) {
         const fm_cell_rect_t rect{0, 0, 1, 0};
         fm_cf_results_t* results = nullptr;
         const fm_status_t rc = fm_sheet_evaluate_cf_range(wb, 0, &rect, 0, std::nan(""), &results);
         fm_cf_results_destroy(results);
         return rc;
       }},
      {"get_cell_count",
       [](fm_workbook_t* wb) {
         size_t count = 0;
         return fm_sheet_get_cell_count(wb, 0, &count);
       }},
      {"get_cell_at",
       [](fm_workbook_t* wb) {
         uint32_t row = 0;
         uint32_t col = 0;
         const char* formula = nullptr;
         fm_value_t v{};
         return fm_sheet_get_cell_at(wb, 0, 0, &row, &col, &formula, &v);
       }},
      {"get_spill",
       [](fm_workbook_t* wb) {
         int32_t present = 0;
         fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
         return fm_sheet_get_spill(wb, 0, 1, 0, &present, &info);
       }},
      {"get_pagination",
       [](fm_workbook_t* wb) {
         fm_pagination_t* pagination = nullptr;
         const fm_status_t rc = fm_sheet_paginate(wb, 0, &pagination);
         fm_pagination_destroy(pagination);
         return rc;
       }},
      {"get_precedents",
       [](fm_workbook_t* wb) {
         fm_cell_nodes_t* nodes = nullptr;
         const fm_status_t rc = fm_sheet_get_precedents(wb, 0, 1, 0, 1, &nodes);
         fm_cell_nodes_destroy(nodes);
         return rc;
       }},
      {"get_dependents",
       [](fm_workbook_t* wb) {
         fm_cell_nodes_t* nodes = nullptr;
         const fm_status_t rc = fm_sheet_get_dependents(wb, 0, 0, 0, 1, &nodes);
         fm_cell_nodes_destroy(nodes);
         return rc;
       }},
      {"evaluate_auto_filter",
       [](fm_workbook_t* wb) {
         fm_auto_filter_t filter = FM_STRUCT_INIT(fm_auto_filter_t);
         filter.range = fm_cell_rect_t{0, 0, 1, 0};
         EXPECT_EQ(fm_sheet_set_auto_filter(wb, 0, &filter), 0);
         uint8_t flags[4] = {};
         size_t len = 0;
         uint32_t first_row = 0;
         return fm_sheet_evaluate_auto_filter(wb, 0, flags, sizeof(flags), &len, &first_row);
       }},
  };
}

TEST(FormulonCApiAutoCalc, AutomaticReadsComputePendingCells) {
  for (const int32_t mode :
       {static_cast<int32_t>(FM_CALC_MODE_AUTO), static_cast<int32_t>(FM_CALC_MODE_AUTO_NO_TABLE)}) {
    for (const Reader& reader : Readers()) {
      SCOPED_TRACE(std::string(reader.name) + " mode=" + std::to_string(mode));
      WorkbookGuard wb;
      MakePending(&wb, mode);
      ASSERT_TRUE(HasDirty(wb.handle));
      EXPECT_EQ(reader.read(wb.handle), 0);
      EXPECT_FALSE(HasDirty(wb.handle));
    }
  }
}

TEST(FormulonCApiAutoCalc, ManualReadsLeavePendingCells) {
  for (const Reader& reader : Readers()) {
    SCOPED_TRACE(reader.name);
    WorkbookGuard wb;
    MakePending(&wb, FM_CALC_MODE_MANUAL);
    EXPECT_EQ(reader.read(wb.handle), 0);
    EXPECT_TRUE(HasDirty(wb.handle));
  }
}

TEST(FormulonCApiAutoCalc, ReadAfterWriteSeesTheNewValue) {
  WorkbookGuard wb;
  MakePending(&wb, FM_CALC_MODE_AUTO);
  fm_value_t v = ValueAt(wb.handle, 1, 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(v.u.number, 20.0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 3.0), 0);
  v = ValueAt(wb.handle, 1, 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(v.u.number, 30.0);

  fm_value_t evaluated{};
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 4.0), 0);
  ASSERT_EQ(evaluate_top_left(wb.handle, 0, 5, 5, "=A2+1", &evaluated), 0);
  ASSERT_EQ(evaluated.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(evaluated.u.number, 41.0);
}

TEST(FormulonCApiAutoCalc, ManualReadsKeepTheCacheUntilRecalc) {
  WorkbookGuard wb;
  MakePending(&wb, FM_CALC_MODE_MANUAL);
  EXPECT_EQ(ValueAt(wb.handle, 1, 0).kind, FM_VALUE_KIND_BLANK);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v = ValueAt(wb.handle, 1, 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(v.u.number, 20.0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 3.0), 0);
  v = ValueAt(wb.handle, 1, 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(v.u.number, 20.0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(ValueAt(wb.handle, 1, 0).u.number, 30.0);
}

TEST(FormulonCApiAutoCalc, SaveFollowsCalcOnSaveInManualMode) {
  struct Case {
    int32_t mode;
    int32_t calc_on_save;
    bool computed;
  };
  const Case cases[] = {
      {FM_CALC_MODE_AUTO, 1, true},
      {FM_CALC_MODE_AUTO, 0, true},
      {FM_CALC_MODE_MANUAL, 1, true},
      {FM_CALC_MODE_MANUAL, 0, false},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE("mode=" + std::to_string(c.mode) + " calc_on_save=" + std::to_string(c.calc_on_save));
    WorkbookGuard wb;
    MakePending(&wb, c.mode);
    ASSERT_EQ(fm_workbook_set_calc_on_save(wb.handle, c.calc_on_save), 0);
    uint8_t* bytes = nullptr;
    size_t len = 0;
    ASSERT_EQ(fm_workbook_save(wb.handle, &bytes, &len), 0);
    EXPECT_EQ(HasDirty(wb.handle), !c.computed);

    // The file carries the value the save computed, or the blank cache.
    WorkbookGuard loaded;
    ASSERT_EQ(fm_workbook_load(bytes, len, &loaded.handle), 0);
    fm_buffer_destroy(bytes);
    ASSERT_EQ(fm_workbook_set_calc_mode(loaded.handle, FM_CALC_MODE_MANUAL), 0);
    const fm_value_t v = ValueAt(loaded.handle, 1, 0);
    if (c.computed) {
      ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
      EXPECT_EQ(v.u.number, 20.0);
    } else {
      EXPECT_FALSE(v.kind == FM_VALUE_KIND_NUMBER && v.u.number == 20.0);
    }
  }
}

TEST(FormulonCApiAutoCalc, SaveVariantsComputePendingCells) {
  for (const int32_t which : {0, 1}) {
    SCOPED_TRACE(which == 0 ? "save_as" : "save_with_diagnostics");
    WorkbookGuard wb;
    MakePending(&wb, FM_CALC_MODE_AUTO);
    uint8_t* bytes = nullptr;
    size_t len = 0;
    if (which == 0) {
      ASSERT_EQ(fm_workbook_save_as(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len), 0);
    } else {
      fm_save_diagnostics_info_t diagnostics = FM_STRUCT_INIT(fm_save_diagnostics_info_t);
      ASSERT_EQ(fm_workbook_save_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len, &diagnostics), 0);
      EXPECT_EQ(diagnostics.struct_size, sizeof(fm_save_diagnostics_info_t));
    }
    fm_buffer_destroy(bytes);
    EXPECT_FALSE(HasDirty(wb.handle));
  }
}

TEST(FormulonCApiAutoCalc, LoadedWorkbookKeepsItsCacheInManualMode) {
  WorkbookGuard source;
  MakePending(&source, FM_CALC_MODE_AUTO);
  uint8_t* bytes = nullptr;
  size_t len = 0;
  ASSERT_EQ(fm_workbook_save(source.handle, &bytes, &len), 0);

  WorkbookGuard automatic;
  ASSERT_EQ(fm_workbook_load(bytes, len, &automatic.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(automatic.handle, 0, 0, 0, 5.0), 0);
  EXPECT_EQ(ValueAt(automatic.handle, 1, 0).u.number, 50.0);

  // Manual right after load reads the file's cached values, edits or not.
  WorkbookGuard manual;
  ASSERT_EQ(fm_workbook_load(bytes, len, &manual.handle), 0);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(fm_workbook_set_calc_mode(manual.handle, FM_CALC_MODE_MANUAL), 0);
  ASSERT_EQ(fm_sheet_set_number(manual.handle, 0, 0, 0, 5.0), 0);
  const fm_value_t v = ValueAt(manual.handle, 1, 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(v.u.number, 20.0);
}

// Statuses of the reads issued from inside an iterative progress callback.
struct ReentryProbe {
  fm_workbook_t* wb = nullptr;
  bool called = false;
  fm_status_t value_status = 0;
  fm_status_t evaluate_status = 0;
};

TEST(FormulonCApiAutoCalc, ReadFromAProgressCallbackIsReentrant) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=B1+1"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=A1/2"), 0);
  fm_iterative_settings_patch_t patch = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  patch.enabled_op = FM_PATCH_OP_SET;
  patch.enabled = 1;
  patch.max_iterations_op = FM_PATCH_OP_SET;
  patch.max_iterations = 2;
  patch.max_change_op = FM_PATCH_OP_SET;
  patch.max_change = 0.001;
  ASSERT_EQ(fm_workbook_update_iterative(wb.handle, &patch), 0);

  ReentryProbe probe;
  probe.wb = wb.handle;
  ASSERT_EQ(fm_workbook_set_iterative_progress(
                wb.handle,
                [](uint32_t, double, uint32_t, void* data) -> int32_t {
                  auto& p = *static_cast<ReentryProbe*>(data);
                  if (!p.called) {
                    p.called = true;
                    fm_value_t v{};
                    p.value_status = fm_sheet_get_value(p.wb, 0, 0, 0, &v);
                    p.evaluate_status = evaluate_top_left(p.wb, 0, 3, 3, "=A1", &v);
                  }
                  return 1;
                },
                &probe),
            0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_iterative_progress(wb.handle, nullptr, nullptr), 0);
  ASSERT_TRUE(probe.called);
  EXPECT_EQ(probe.value_status, FM_STATUS_GRAPH_RECALC_REENTRANT);
  EXPECT_EQ(probe.evaluate_status, FM_STATUS_GRAPH_RECALC_REENTRANT);
}

}  // namespace
