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
  // A fresh live model intentionally stays empty. The OOXML writer
  // synthesizes the default cellStyleXfs/Normal pair only in the
  // serialized document, so the pair becomes visible after reload.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t live_cs_count = 99;
  uint32_t live_cs_xf_count = 99;
  ASSERT_EQ(fm_styles_get_cell_style_count(wb.handle, &live_cs_count), 0);
  ASSERT_EQ(fm_styles_get_cell_style_xf_count(wb.handle, &live_cs_xf_count), 0);
  EXPECT_EQ(live_cs_count, 0U);
  EXPECT_EQ(live_cs_xf_count, 0U);

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
  EXPECT_EQ(cs_count, 1U);
  EXPECT_EQ(cs_xf_count, 1U);

  fm_cell_style_record_t normal{};
  ASSERT_EQ(fm_styles_get_cell_style(wb2.handle, 0, &normal), 0);
  ASSERT_NE(normal.name, nullptr);
  EXPECT_STREQ(normal.name, "Normal");
  EXPECT_EQ(normal.xf_id, 0U);
  EXPECT_EQ(normal.builtin_id, 0U);
  EXPECT_EQ(normal.i_level, 0U);
  EXPECT_EQ(normal.hidden, 0);
  EXPECT_EQ(normal.custom_builtin, 0);

  fm_cell_xf normal_xf{};
  ASSERT_EQ(fm_styles_get_cell_style_xf(wb2.handle, 0, &normal_xf), 0);
  EXPECT_EQ(normal_xf.font_index, 0U);
  EXPECT_EQ(normal_xf.fill_index, 0U);
  EXPECT_EQ(normal_xf.border_index, 0U);
  EXPECT_EQ(normal_xf.num_fmt_id, 0U);
  EXPECT_EQ(normal_xf.horizontal_align, 0U);
  EXPECT_EQ(normal_xf.vertical_align, 2U);
  EXPECT_EQ(normal_xf.wrap_text, 0);
}
