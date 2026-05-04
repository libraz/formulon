// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stable C ABI tests for the named-cell-style surface
// (`fm_styles_get_cell_style_count` / `fm_styles_get_cell_style` and
// the parallel `fm_styles_get_cell_style_xf*` accessors). The actual
// reader/writer round-trip lives in the integration suite; these tests
// exercise the boundary's null / range guards and the empty-workbook
// behaviour.

#include <cstdint>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

}  // namespace

TEST(FormulonCApiNamedStyles, FreshWorkbookHasNoNamedStyles) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t cs_count = 99;
  uint32_t cs_xf_count = 99;
  EXPECT_EQ(fm_styles_get_cell_style_count(wb.handle, &cs_count), 0);
  EXPECT_EQ(fm_styles_get_cell_style_xf_count(wb.handle, &cs_xf_count), 0);
  EXPECT_EQ(cs_count, 0U);
  EXPECT_EQ(cs_xf_count, 0U);
}

TEST(FormulonCApiNamedStyles, IndexOutOfRangeReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_style_record_t cs{};
  fm_cell_xf xf{};
  EXPECT_EQ(fm_styles_get_cell_style(wb.handle, 0, &cs),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_styles_get_cell_style_xf(wb.handle, 0, &xf),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiNamedStyles, NullArgsReturnBindingNullPointer) {
  fm_cell_style_record_t cs{};
  fm_cell_xf xf{};
  uint32_t count = 0;
  EXPECT_EQ(fm_styles_get_cell_style_count(nullptr, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_styles_get_cell_style_xf_count(nullptr, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_styles_get_cell_style(nullptr, 0, &cs),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_styles_get_cell_style_xf(nullptr, 0, &xf),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_styles_get_cell_style_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_styles_get_cell_style_xf_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_styles_get_cell_style(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_styles_get_cell_style_xf(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiNamedStyles, RoundTripExposesNamedStylesAfterReload) {
  // The named-style tables can only be populated through the
  // reader path today (no `fm_styles_add_cell_style` setter exists).
  // To verify the C ABI getters surface the tables once they are
  // present, we save / reload a workbook whose styles.xml has been
  // constructed with named styles via the C++ side, then read it back
  // through the public C ABI.
  //
  // The C++ population step lives in the integration suite
  // (`StylesRoundTrip.PreservesNamedCellStyles`). This test only
  // verifies that an empty workbook saved through `fm_workbook_save`
  // and re-opened via `fm_workbook_load` continues to report
  // count == 0 — the most we can assert without a setter on the
  // pure-C surface.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint8_t* saved_data = nullptr;
  size_t saved_len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved_data, &saved_len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved_data, saved_len, &wb2.handle), 0);
  fm_buffer_free(saved_data);

  uint32_t cs_count = 99;
  uint32_t cs_xf_count = 99;
  EXPECT_EQ(fm_styles_get_cell_style_count(wb2.handle, &cs_count), 0);
  EXPECT_EQ(fm_styles_get_cell_style_xf_count(wb2.handle, &cs_xf_count), 0);
  EXPECT_EQ(cs_count, 0U);
  EXPECT_EQ(cs_xf_count, 0U);
}
