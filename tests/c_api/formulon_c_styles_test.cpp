//
// Stable C ABI tests for the styles surface
// (`fm_cell_get_xf_index`, `fm_cell_set_xf_index`,
// `fm_styles_get_cell_xf`, `fm_styles_get_font`,
// `fm_styles_get_num_fmt_string`).
//
// The test driver is C++ for gtest convenience but everything it
// touches across the boundary is the pure-C surface.

#include <cstdint>
#include <cstring>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "io/styles_reader.h"
#include "io/styles_writer.h"
#include "workbook.h"

namespace {

static_assert(sizeof(fm_cell_xf) == 20U, "fm_cell_xf ABI layout changed");
static_assert(offsetof(fm_cell_xf, font_index) == 0U, "fm_cell_xf.font_index offset changed");
static_assert(offsetof(fm_cell_xf, fill_index) == 4U, "fm_cell_xf.fill_index offset changed");
static_assert(offsetof(fm_cell_xf, border_index) == 8U, "fm_cell_xf.border_index offset changed");
static_assert(offsetof(fm_cell_xf, num_fmt_id) == 12U, "fm_cell_xf.num_fmt_id offset changed");
static_assert(offsetof(fm_cell_xf, horizontal_align) == 14U, "fm_cell_xf.horizontal_align offset changed");
static_assert(offsetof(fm_cell_xf, vertical_align) == 15U, "fm_cell_xf.vertical_align offset changed");
static_assert(offsetof(fm_cell_xf, wrap_text) == 16U, "fm_cell_xf.wrap_text offset changed");
static_assert(sizeof(fm_cell_xf_ex) == 28U, "fm_cell_xf_ex ABI layout changed");
static_assert(offsetof(fm_cell_xf_ex, justify_last_line) == 20U, "fm_cell_xf_ex.justify_last_line offset changed");
static_assert(offsetof(fm_cell_xf_ex, xf_id) == 24U, "fm_cell_xf_ex.xf_id offset changed");
static_assert(sizeof(fm_cell_xf_ex2) == 88U, "fm_cell_xf_ex2 ABI layout changed");
static_assert(offsetof(fm_cell_xf_ex2, has_alignment) == 28U, "fm_cell_xf_ex2.has_alignment offset changed");
static_assert(offsetof(fm_cell_xf_ex2, has_text_rotation) == 32U, "fm_cell_xf_ex2.has_text_rotation offset changed");
static_assert(offsetof(fm_cell_xf_ex2, reading_order) == 68U, "fm_cell_xf_ex2.reading_order offset changed");
static_assert(offsetof(fm_cell_xf_ex2, has_horizontal_align) == 72U,
              "fm_cell_xf_ex2.has_horizontal_align offset changed");
static_assert(offsetof(fm_cell_xf_ex2, has_vertical_align) == 76U, "fm_cell_xf_ex2.has_vertical_align offset changed");
static_assert(offsetof(fm_cell_xf_ex2, has_wrap_text) == 80U, "fm_cell_xf_ex2.has_wrap_text offset changed");
static_assert(offsetof(fm_cell_xf_ex2, has_justify_last_line) == 84U,
              "fm_cell_xf_ex2.has_justify_last_line offset changed");

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

}  // namespace

TEST(FormulonCApiStyles, CellXfIndexDefaultsToZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t xf = 99;
  EXPECT_EQ(fm_cell_get_xf_index(wb.handle, 0, 0, 0, &xf), 0);
  EXPECT_EQ(xf, 0U);
}

TEST(FormulonCApiStyles, CellXfIndexRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_cell_set_xf_index(wb.handle, 0, 5, 7, 42), 0);
  uint32_t xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb.handle, 0, 5, 7, &xf), 0);
  EXPECT_EQ(xf, 42U);
}

TEST(FormulonCApiStyles, SetXfIndexRejectsBadSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // sheet 99 does not exist.
  EXPECT_NE(fm_cell_set_xf_index(wb.handle, 99, 0, 0, 1), 0);
}

TEST(FormulonCApiStyles, NullArgumentRejected) {
  uint32_t xf = 0;
  EXPECT_NE(fm_cell_get_xf_index(nullptr, 0, 0, 0, &xf), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_cell_get_xf_index(wb.handle, 0, 0, 0, nullptr), 0);
}

TEST(FormulonCApiStyles, BuiltinNumFmtResolves) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  EXPECT_EQ(fm_styles_get_num_fmt_string(wb.handle, 0, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "General");
  EXPECT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "mm-dd-yy");
}

TEST(FormulonCApiStyles, CustomNumFmtOverridingBuiltinIdWins) {
  // A file may define a custom <numFmt> whose numFmtId collides with a
  // built-in slot; Excel honours the file's definition. Inject such an
  // override and confirm the getter returns the custom string, not the
  // built-in.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  formulon::io::StylesTable& styles = wb.handle->workbook().mutable_styles();
  formulon::io::NumFmtRecord rec;
  rec.id = 14;  // Built-in id whose default is "mm-dd-yy".
  rec.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("yyyy\"年\"m\"月\"d\"日\"");
  styles.num_fmts.push_back(rec);

  const char* s = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "yyyy\"年\"m\"月\"d\"日\"");
}

TEST(FormulonCApiStyles, UnknownNumFmtIdRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  // Reserved built-in slot (id 5) without a custom override surfaces
  // an error rather than an empty string.
  EXPECT_NE(fm_styles_get_num_fmt_string(wb.handle, 5, &s), 0);
}

TEST(FormulonCApiStyles, DifferentialFormatGetterExposesCfDxfTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  formulon::io::StylesTable& styles = wb.handle->workbook().mutable_styles();
  formulon::io::DifferentialFormat dxf;
  dxf.has_font = true;
  dxf.font.bold = true;
  dxf.font.color_argb = 0xFFFF0000U;
  dxf.has_fill = true;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFFFFFF00U;
  dxf.has_border = true;
  dxf.border.left.style = 1;
  dxf.border.left.color_argb = 0xFF000000U;
  dxf.has_num_fmt = true;
  dxf.num_fmt_id = 164;
  dxf.num_fmt_code = "0.00";
  styles.dxfs.push_back(dxf);

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_dxf_record out{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, 0, &out), 0);
  EXPECT_EQ(out.font_engaged, 1);
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.font.color_argb, 0xFFFF0000U);
  EXPECT_EQ(out.fill_engaged, 1);
  EXPECT_EQ(out.fill.pattern, 1U);
  EXPECT_EQ(out.fill.fg_argb, 0xFFFFFF00U);
  EXPECT_EQ(out.border_engaged, 1);
  EXPECT_EQ(out.border.left.style, 1U);
  EXPECT_EQ(out.border.left.color_argb, 0xFF000000U);
  EXPECT_EQ(out.num_fmt_engaged, 1);
  EXPECT_EQ(out.num_fmt_id, 164U);
  ASSERT_NE(out.num_fmt_code, nullptr);
  EXPECT_STREQ(out.num_fmt_code, "0.00");

  EXPECT_NE(fm_styles_get_dxf(wb.handle, 1, &out), 0);
}

TEST(FormulonCApiStyles, AddDxfDedupsAndReadsBack) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_dxf_record dxf{};
  dxf.font_engaged = 1;
  dxf.font.name = "Arial";
  dxf.font.size = 12.0;
  dxf.font.bold = 1;
  dxf.font.color_argb = 0xFFFF0000U;
  dxf.fill_engaged = 1;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFFFFFF00U;
  dxf.num_fmt_engaged = 1;
  dxf.num_fmt_id = 164;
  dxf.num_fmt_code = "0.00";

  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &a), 0);
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &b), 0);
  EXPECT_EQ(a, b);

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_dxf_record out{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, a, &out), 0);
  EXPECT_EQ(out.font_engaged, 1);
  EXPECT_STREQ(out.font.name, "Arial");
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.font.color_argb, 0xFFFF0000U);
  EXPECT_EQ(out.fill_engaged, 1);
  EXPECT_EQ(out.fill.pattern, 1U);
  EXPECT_EQ(out.fill.fg_argb, 0xFFFFFF00U);
  EXPECT_EQ(out.num_fmt_engaged, 1);
  EXPECT_EQ(out.num_fmt_id, 164U);
  ASSERT_NE(out.num_fmt_code, nullptr);
  EXPECT_STREQ(out.num_fmt_code, "0.00");
}

TEST(FormulonCApiStyles, DxfFontRoundTripsVerticalAlignmentThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_dxf_record dxf{};
  dxf.font_engaged = 1;
  dxf.font.name = "Calibri";
  dxf.font.size = 9.0;
  dxf.font.color_argb = 0xFF112233U;
  dxf.font.vert_align = 1;  // superscript
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &index), 0);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<vertAlign val=\"superscript\"/>"), std::string::npos);

  fm_dxf_record loaded{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, index, &loaded), 0);
  EXPECT_EQ(loaded.font.vert_align, 1U);

  uint32_t count_before = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count_before), 0);
  uint32_t reindex = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, loaded, &reindex), 0);
  EXPECT_EQ(reindex, index);
  uint32_t count_after = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count_before);
}

TEST(FormulonCApiStyles, DxfFontDistinguishesAnExplicitBoldOffFromAnAbsentToggle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  formulon::io::DifferentialFormat stored;
  stored.has_font = true;
  stored.font.has_bold = true;  // `<b val="0"/>`: switch bold off
  stored.font.bold = false;
  wb.handle->workbook().mutable_styles().dxfs.push_back(stored);

  // A rule that leaves bold alone must not fold onto the switch-it-off one.
  fm_dxf_record leave_bold_alone{};
  leave_bold_alone.font_engaged = 1;
  leave_bold_alone.font.name = "";
  leave_bold_alone.font.size = 11.0;
  leave_bold_alone.font.color_argb = 0xFF000000U;
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, leave_bold_alone, &index), 0);
  EXPECT_NE(index, 0U);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<b val=\"0\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, GetCellXfRejectsOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  // A fresh workbook has an empty styles table, so any xf_index is
  // out of range. Calling with index 0 still fails because the table
  // has zero entries.
  EXPECT_NE(fm_styles_get_cell_xf(wb.handle, 0, &xf), 0);
}

TEST(FormulonCApiStyles, CellXfGetterDiagnosticsUseInvokedApi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  auto expect_prefix = [](fm_status_t status, const char* prefix) {
    EXPECT_NE(status, 0);
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(prefix, 0), 0U) << message;
  };

  fm_cell_xf cell_xf{};
  fm_cell_xf_ex cell_xf_ex{};
  fm_cell_xf_ex2 cell_xf_ex2{};
  expect_prefix(fm_styles_get_cell_xf(wb.handle, 0, &cell_xf), "fm_styles_get_cell_xf:");
  expect_prefix(fm_styles_get_cell_xf_ex(wb.handle, 0, &cell_xf_ex), "fm_styles_get_cell_xf_ex:");
  expect_prefix(fm_styles_get_cell_xf_ex2(wb.handle, 0, &cell_xf_ex2), "fm_styles_get_cell_xf_ex2:");
  expect_prefix(fm_styles_get_cell_style_xf(wb.handle, 0, &cell_xf), "fm_styles_get_cell_style_xf:");
  expect_prefix(fm_styles_get_cell_style_xf_ex(wb.handle, 0, &cell_xf_ex), "fm_styles_get_cell_style_xf_ex:");
  expect_prefix(fm_styles_get_cell_style_xf_ex2(wb.handle, 0, &cell_xf_ex2), "fm_styles_get_cell_style_xf_ex2:");
}

TEST(FormulonCApiStyles, CellXfAdderDiagnosticsUseInvokedApi) {
  auto expect_prefix = [](fm_status_t status, const char* prefix) {
    EXPECT_NE(status, 0);
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(prefix, 0), 0U) << message;
  };

  uint32_t index = 0;
  fm_cell_xf cell_xf{};
  fm_cell_xf_ex cell_xf_ex{};
  fm_cell_xf_ex2 cell_xf_ex2{};
  expect_prefix(fm_styles_add_cell_xf(nullptr, cell_xf, &index), "fm_styles_add_cell_xf:");
  expect_prefix(fm_styles_add_cell_xf_ex(nullptr, cell_xf_ex, &index), "fm_styles_add_cell_xf_ex:");
  expect_prefix(fm_styles_add_cell_xf_ex2(nullptr, cell_xf_ex2, &index), "fm_styles_add_cell_xf_ex2:");

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  cell_xf.font_index = 1U;
  cell_xf_ex.base.font_index = 1U;
  cell_xf_ex2.base.font_index = 1U;
  expect_prefix(fm_styles_add_cell_xf(wb.handle, cell_xf, &index), "fm_styles_add_cell_xf:");
  expect_prefix(fm_styles_add_cell_xf_ex(wb.handle, cell_xf_ex, &index), "fm_styles_add_cell_xf_ex:");
  expect_prefix(fm_styles_add_cell_xf_ex2(wb.handle, cell_xf_ex2, &index), "fm_styles_add_cell_xf_ex2:");

  cell_xf_ex = fm_cell_xf_ex{};
  cell_xf_ex.xf_id = 1U;
  expect_prefix(fm_styles_add_cell_xf_ex(wb.handle, cell_xf_ex, &index), "fm_styles_add_cell_xf_ex:");
  cell_xf_ex2 = fm_cell_xf_ex2{};
  cell_xf_ex2.xf_id = 1U;
  expect_prefix(fm_styles_add_cell_xf_ex2(wb.handle, cell_xf_ex2, &index), "fm_styles_add_cell_xf_ex2:");

  cell_xf_ex2 = fm_cell_xf_ex2{};
  cell_xf_ex2.has_text_rotation = 1;
  cell_xf_ex2.text_rotation = 181U;
  expect_prefix(fm_styles_add_cell_xf_ex2(wb.handle, cell_xf_ex2, &index), "fm_styles_add_cell_xf_ex2:");
}

TEST(FormulonCApiStyles, CellXfAlignmentEnumsValidateRangesBeforeMutation) {
  const auto expect_invalid = [](fm_status_t status, const char* api) {
    EXPECT_EQ(status, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(api, 0), 0U) << message;
  };

  // The upper boundary of each ordinal is accepted by the legacy entrypoint.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 7;  // distributed
    record.vertical_align = 4;    // distributed
    uint32_t index = 0;
    ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);
  }

  // Every add shape validates before it creates default roots or changes the
  // output index. Keep each case isolated so the size checks cover all paths.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 8;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_xfs.size();
    expect_invalid(fm_styles_add_cell_xf(wb.handle, record, &index), "fm_styles_add_cell_xf:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf_ex record{};
    record.base.vertical_align = 5;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_xfs.size();
    expect_invalid(fm_styles_add_cell_xf_ex(wb.handle, record, &index), "fm_styles_add_cell_xf_ex:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf_ex2 record{};
    record.base.horizontal_align = 8;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_xfs.size();
    expect_invalid(fm_styles_add_cell_xf_ex2(wb.handle, record, &index), "fm_styles_add_cell_xf_ex2:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf_ex record{};
    record.base.horizontal_align = 8;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_style_xfs.size();
    expect_invalid(fm_styles_add_cell_style_xf_ex(wb.handle, record, &index), "fm_styles_add_cell_style_xf_ex:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_style_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf_ex2 record{};
    record.base.vertical_align = 5;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_style_xfs.size();
    expect_invalid(fm_styles_add_cell_style_xf_ex2(wb.handle, record, &index), "fm_styles_add_cell_style_xf_ex2:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_style_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 8;
    uint32_t indices[] = {0xAABBCCDDU};
    const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr, 0U,      nullptr, nullptr, 0U,
                                nullptr, &record, 1U,      indices, nullptr, 0U,      nullptr};
    const std::size_t before = wb.handle->workbook().styles().cell_xfs.size();
    expect_invalid(fm_styles_add_batch(wb.handle, &batch), "fm_styles_add_batch:");
    EXPECT_EQ(indices[0], 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before);
  }
}

TEST(FormulonCApiStyles, SetThenSaveLoadPreservesXfIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Stamp xf_index = 7 on cell A1 and save.
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 3.14), 0);
  EXPECT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, 7), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  uint32_t xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 0, 0, &xf), 0);
  EXPECT_EQ(xf, 7U);
}

// ---- Add-side dedup primitives -------------------------------------

namespace {

fm_font_record MakeArial() {
  fm_font_record r{};
  r.name = "Arial";
  r.size = 12.0;
  r.color_argb = 0xFF112233U;
  r.bold = 1;
  r.italic = 0;
  r.strike = 0;
  r.underline = 0;
  return r;
}

fm_fill_record MakeRedFill() {
  fm_fill_record r{};
  r.pattern = 1;  // solid
  r.fg_argb = 0xFFFF0000U;
  r.bg_argb = 0xFF000000U;
  return r;
}

fm_border_record MakeThinBoxBorder() {
  fm_border_record r{};
  r.left.style = 1;  // thin
  r.left.color_argb = 0xFF000000U;
  r.right = r.left;
  r.top = r.left;
  r.bottom = r.left;
  r.diagonal.style = 0;
  r.diagonal.color_argb = 0;
  r.diagonal_up = 0;
  r.diagonal_down = 0;
  return r;
}

}  // namespace

TEST(FormulonCApiStyles, AddFontDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &a), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddFontDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_font_record arial = MakeArial();
  fm_font_record calibri = MakeArial();
  calibri.name = "Calibri";
  ASSERT_EQ(fm_styles_add_font(wb.handle, arial, &a), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, calibri, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddFontGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  // Round-trip: the freshly added font should read back equal.
  fm_font_record out{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &out), 0);
  EXPECT_STREQ(out.name, "Arial");
  EXPECT_DOUBLE_EQ(out.size, 12.0);
  EXPECT_EQ(out.color_argb, 0xFF112233U);
  EXPECT_EQ(out.bold, 1);
}

TEST(FormulonCApiStyles, GetFontRoundTripIsIdentityForSuperscript) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record superscript = MakeArial();
  superscript.vert_align = 1;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, superscript, &idx), 0);

  fm_font_record loaded{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &loaded), 0);
  EXPECT_STREQ(loaded.name, "Arial");
  EXPECT_EQ(loaded.vert_align, 1U);

  uint32_t count_before = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_before), 0);
  uint32_t reindex = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, loaded, &reindex), 0);
  EXPECT_EQ(reindex, idx);
  uint32_t count_after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count_before);
}

TEST(FormulonCApiStyles, GetEditAddFontPreservesTheUneditedFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record superscript = MakeArial();
  superscript.vert_align = 1;
  superscript.has_family = 1;
  superscript.family = 2;
  superscript.has_charset = 1;
  superscript.charset = 128;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, superscript, &idx), 0);

  fm_font_record edited{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &edited), 0);
  edited.color_argb = 0xFF00FF00U;
  uint32_t recolored = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, edited, &recolored), 0);
  EXPECT_NE(recolored, idx);

  fm_font_record reread{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, recolored, &reread), 0);
  EXPECT_EQ(reread.color_argb, 0xFF00FF00U);
  EXPECT_EQ(reread.vert_align, 1U);
  EXPECT_EQ(reread.has_family, 1);
  EXPECT_EQ(reread.family, 2U);
  EXPECT_EQ(reread.has_charset, 1);
  EXPECT_EQ(reread.charset, 128U);
}

namespace {

// An Excel-authored `<fonts>` section: the theme colour, `<family>` and
// `<charset>` children are what a template carries and what a record built
// from scratch through the C ABI cannot reproduce unless the ABI exposes
// them.
constexpr char kExcelAuthoredStyles[] =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<fonts count=\"3\">"
    "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/>"
    "<charset val=\"128\"/></font>"
    "<font><b/><sz val=\"11\"/><color theme=\"0\" tint=\"-0.25\"/><name val=\"Calibri\"/>"
    "<family val=\"2\"/></font>"
    "<font><vertAlign val=\"superscript\"/><sz val=\"9\"/><color rgb=\"FF112233\"/>"
    "<name val=\"Calibri\"/></font>"
    "</fonts>"
    "<fills count=\"2\">"
    "<fill><patternFill patternType=\"none\"/></fill>"
    "<fill><patternFill patternType=\"solid\"><fgColor theme=\"4\" tint=\"0.5\"/>"
    "<bgColor indexed=\"64\"/></patternFill></fill>"
    "</fills>"
    "<borders count=\"2\">"
    "<border><left/><right/><top/><bottom/><diagonal/></border>"
    "<border><left style=\"thin\"><color theme=\"3\"/></left><right/><top/><bottom/>"
    "<diagonal/></border>"
    "</borders>"
    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>"
    "</styleSheet>";

/// Installs `kExcelAuthoredStyles` as the workbook's styles table.
void LoadExcelAuthoredStyles(fm_workbook_t* handle) {
  const std::vector<std::uint8_t> bytes(kExcelAuthoredStyles, kExcelAuthoredStyles + std::strlen(kExcelAuthoredStyles));
  auto parsed = formulon::io::read_styles(bytes);
  ASSERT_TRUE(parsed.has_value());
  handle->workbook().mutable_styles() = std::move(parsed.value());
}

}  // namespace

TEST(FormulonCApiStyles, AddFontIsIdentityAgainstAFileLoadedTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count), 0);
  ASSERT_EQ(count, 3U);
  for (uint32_t i = 0; i < count; ++i) {
    fm_font_record loaded{};
    ASSERT_EQ(fm_styles_get_font(wb.handle, i, &loaded), 0) << "font " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_font(wb.handle, loaded, &reindex), 0) << "font " << i;
    EXPECT_EQ(reindex, i);
  }
  uint32_t count_after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count);
}

TEST(FormulonCApiStyles, AddFillAndBorderAreIdentityAgainstAFileLoadedTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  uint32_t fills = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &fills), 0);
  ASSERT_EQ(fills, 2U);
  for (uint32_t i = 0; i < fills; ++i) {
    fm_fill_record loaded{};
    ASSERT_EQ(fm_styles_get_fill(wb.handle, i, &loaded), 0) << "fill " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_fill(wb.handle, loaded, &reindex), 0) << "fill " << i;
    EXPECT_EQ(reindex, i);
  }

  uint32_t borders = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &borders), 0);
  ASSERT_EQ(borders, 2U);
  for (uint32_t i = 0; i < borders; ++i) {
    fm_border_record loaded{};
    ASSERT_EQ(fm_styles_get_border(wb.handle, i, &loaded), 0) << "border " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_border(wb.handle, loaded, &reindex), 0) << "border " << i;
    EXPECT_EQ(reindex, i);
  }

  uint32_t fills_after = 0;
  uint32_t borders_after = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &fills_after), 0);
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &borders_after), 0);
  EXPECT_EQ(fills_after, fills);
  EXPECT_EQ(borders_after, borders);
}

TEST(FormulonCApiStyles, AddFontDoesNotAliasAThemeColourOntoAResolvedRgb) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  // Font 0 is `<color theme="1"/>`, which the reader resolves to the same
  // AARRGGBB a caller would pass for plain black. Folding the two together
  // is what turned a template's body text white.
  fm_font_record themed{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, 0U, &themed), 0);
  ASSERT_EQ(themed.color.kind, static_cast<uint8_t>(kFmColorTheme));

  fm_font_record resolved = themed;
  resolved.color = fm_color_spec{};
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, resolved, &index), 0);
  EXPECT_NE(index, 0U);

  // The theme record is untouched and still serialises as a theme colour.
  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<color theme=\"1\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<color theme=\"0\" tint=\"-0.25\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, AddFillAndBorderDistinguishColourSpecifications) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  fm_fill_record themed_fill{};
  ASSERT_EQ(fm_styles_get_fill(wb.handle, 1U, &themed_fill), 0);
  ASSERT_EQ(themed_fill.fg.kind, static_cast<uint8_t>(kFmColorTheme));
  fm_fill_record resolved_fill = themed_fill;
  resolved_fill.fg = fm_color_spec{};
  resolved_fill.fg_argb = themed_fill.fg_argb;
  uint32_t fill_index = 0;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, resolved_fill, &fill_index), 0);
  EXPECT_NE(fill_index, 1U);

  fm_border_record themed_border{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, 1U, &themed_border), 0);
  ASSERT_EQ(themed_border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  fm_border_record resolved_border = themed_border;
  resolved_border.left.color = fm_color_spec{};
  uint32_t border_index = 0;
  ASSERT_EQ(fm_styles_add_border(wb.handle, resolved_border, &border_index), 0);
  EXPECT_NE(border_index, 1U);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<fgColor theme=\"4\" tint=\"0.5\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<bgColor indexed=\"64\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, AddFontDistinguishesAnExplicitOffToggleFromAnAbsentOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record absent = MakeArial();
  absent.bold = 0;
  absent.has_bold = 0;
  fm_font_record explicit_off = absent;
  explicit_off.has_bold = 1;

  uint32_t absent_index = 0;
  uint32_t explicit_index = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, absent, &absent_index), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, explicit_off, &explicit_index), 0);
  EXPECT_NE(absent_index, explicit_index);
}

TEST(FormulonCApiStyles, AddFillDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &a), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddFillDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_fill_record red = MakeRedFill();
  fm_fill_record blue = MakeRedFill();
  blue.fg_argb = 0xFF0000FFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, red, &a), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, blue, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddFillGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  fm_fill_record out{};
  ASSERT_EQ(fm_styles_get_fill(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.pattern, 1U);
  EXPECT_EQ(out.fg_argb, 0xFFFF0000U);
}

TEST(FormulonCApiStyles, AddBorderDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &a), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddBorderDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_border_record thin = MakeThinBoxBorder();
  fm_border_record dashed = MakeThinBoxBorder();
  dashed.left.style = 3;  // dashed
  ASSERT_EQ(fm_styles_add_border(wb.handle, thin, &a), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, dashed, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddBorderGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  fm_border_record out{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.left.style, 1U);
  EXPECT_EQ(out.right.style, 1U);
  EXPECT_EQ(out.diagonal_up, 0);
  EXPECT_EQ(out.diagonal_down, 0);
}

TEST(FormulonCApiStyles, AddNumFmtBuiltinReturnsBuiltinId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "General", &id), 0);
  EXPECT_EQ(id, 0U);
  // Adding a built-in must not create a custom entry.
  uint32_t font_count = 7;  // unrelated, just ensure other tables untouched
  EXPECT_EQ(fm_styles_get_font_count(wb.handle, &font_count), 0);
  EXPECT_EQ(font_count, 0U);
}

TEST(FormulonCApiStyles, AddNumFmtCustomReturnsCustomId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t id = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &id), 0);
  EXPECT_GE(id, 164U);
  // Resolves through the read-side getter.
  const char* s = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, id, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "\"USD\" #,##0");
}

TEST(FormulonCApiStyles, AddNumFmtCustomDedup) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t a = 0;
  uint16_t b = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &a), 0);
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddNumFmtExhaustionReturnsPreconditionWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  auto& styles = wb.handle->workbook().mutable_styles();
  formulon::io::NumFmtRecord max_record;
  max_record.id = 65535U;
  max_record.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("existing");
  styles.num_fmts.push_back(max_record);
  const std::size_t strings_before = styles.num_fmt_strings.size();
  const std::size_t records_before = styles.num_fmts.size();
  uint16_t out = 0xBEEFU;
  EXPECT_EQ(fm_styles_add_num_fmt(wb.handle, "new-format", &out),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(out, 0xBEEFU);
  EXPECT_EQ(styles.num_fmt_strings.size(), strings_before);
  EXPECT_EQ(styles.num_fmts.size(), records_before);
}

TEST(FormulonCApiStyles, AddBatchNumFmtExhaustionIsTransactional) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  auto& styles = wb.handle->workbook().mutable_styles();
  formulon::io::NumFmtRecord max_record;
  max_record.id = 65535U;
  max_record.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("existing");
  styles.num_fmts.push_back(max_record);
  const std::string before_xml = formulon::io::write_styles(styles);
  const char* codes[] = {"first", "second"};
  uint16_t ids[] = {0xAAAAU, 0xBBBBU};
  const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr, 0U,    nullptr, nullptr, 0U,
                              nullptr, nullptr, 0U,      nullptr, codes, 2U,      ids};
  EXPECT_EQ(fm_styles_add_batch(wb.handle, &batch),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(ids[0], 0xAAAAU);
  EXPECT_EQ(ids[1], 0xBBBBU);
  EXPECT_EQ(formulon::io::write_styles(styles), before_xml);
}

TEST(FormulonCApiStyles, AddCellXfDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;  // built-in General
  xf.horizontal_align = 1;
  xf.vertical_align = 2;
  xf.wrap_text = 1;

  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &a), 0);
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddCellXfDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  xf.horizontal_align = 1;
  xf.vertical_align = 2;
  xf.wrap_text = 1;

  uint32_t a = 0;
  uint32_t b = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &a), 0);
  // Flip wrap_text — distinct record.
  xf.wrap_text = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddCellXfGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &before), 0);
  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it.
  xf.wrap_text = 1;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 2U);
  EXPECT_EQ(idx, 1U);
}

TEST(FormulonCApiStyles, AddCellXfOnFreshWorkbookKeepsZeroAsDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  // A fresh workbook's tables start empty, so the first font/fill/border
  // added legitimately becomes index 0 — there is no anonymous placeholder
  // seeded ahead of it.
  EXPECT_EQ(font_idx, 0U);
  EXPECT_EQ(fill_idx, 0U);
  EXPECT_EQ(border_idx, 0U);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it.
  xf.wrap_text = 1;

  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);
  EXPECT_NE(xf_idx, 0U);

  fm_cell_xf default_xf{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, 0, &default_xf), 0);
  EXPECT_EQ(default_xf.font_index, 0U);
  EXPECT_EQ(default_xf.fill_index, 0U);
  EXPECT_EQ(default_xf.border_index, 0U);
  EXPECT_EQ(default_xf.num_fmt_id, 0U);
  EXPECT_EQ(default_xf.wrap_text, 0);
}

TEST(FormulonCApiStyles, AddCellXfRejectsOutOfRangeIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Empty styles table: any non-zero index is out of range.
  fm_cell_xf xf{};
  xf.font_index = 5;  // intentionally OOR
  xf.fill_index = 0;
  xf.border_index = 0;
  xf.num_fmt_id = 0;
  uint32_t idx = 0xFFFFFFFFU;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
}

TEST(FormulonCApiStyles, AddCellXfRejectsUnknownNumFmt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 200;  // not a registered custom and not a documented built-in
  uint32_t idx = 0xFFFFFFFFU;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
}

TEST(FormulonCApiStyles, AddNullArgumentRejected) {
  uint32_t idx = 0;
  EXPECT_NE(fm_styles_add_font(nullptr, MakeArial(), &idx), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_styles_add_font(wb.handle, MakeArial(), nullptr), 0);
  uint16_t id = 0;
  EXPECT_NE(fm_styles_add_num_fmt(nullptr, "General", &id), 0);
  EXPECT_NE(fm_styles_add_num_fmt(wb.handle, "General", nullptr), 0);
  EXPECT_NE(fm_styles_add_dxf(nullptr, fm_dxf_record{}, &idx), 0);
  EXPECT_NE(fm_styles_add_dxf(wb.handle, fm_dxf_record{}, nullptr), 0);
}

TEST(FormulonCApiStyles, FullLifecycleSurvivesSaveLoad) {
  // Build font -> xf -> stamp on cell -> save -> reload -> the font is
  // resolvable via the read-side accessors.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);

  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.5), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  uint32_t reread_xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 0, 0, &reread_xf), 0);
  fm_cell_xf reloaded{};
  EXPECT_EQ(fm_styles_get_cell_xf(wb2.handle, reread_xf, &reloaded), 0);
  fm_font_record reloaded_font{};
  ASSERT_EQ(fm_styles_get_font(wb2.handle, reloaded.font_index, &reloaded_font), 0);
  EXPECT_STREQ(reloaded_font.name, "Arial");
  EXPECT_DOUBLE_EQ(reloaded_font.size, 12.0);
  EXPECT_EQ(reloaded_font.bold, 1);
}

// Regression coverage for the externally-reported symptom fixed by removing
// the spurious `ensure_default_style_roots` seeding from the font/fill/
// border adders: a non-default font + fill + custom numFmt combined into an
// `<xf>` must produce an `<xf>` index that is distinct from (and does not
// dedup to) the all-zero placeholder xf, and that index must survive
// `setCellXfIndex` -> save -> reload with the font, fill, and numFmt content
// intact on the reloaded workbook. Note that a fresh workbook's font/fill
// tables legitimately start at index 0 (that is the behavior the fix
// restored), so this test asserts on round-tripped *content*, not on the
// font/fill indices being non-zero. The WASM (`src/wasm/parts/
// workbook_styles.cpp`) and Node addon (`src/node_addon/parts/styles.cc`)
// bindings marshal the identical field set into `fm_cell_xf` and call this
// same `fm_styles_add_cell_xf` entry point, so this test also exercises
// their shared code path end to end.
TEST(FormulonCApiStyles, AddXfWithNonDefaultFontFillNumFmtRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t font_idx = 0xFFFFFFFFU;
  uint32_t fill_idx = 0xFFFFFFFFU;
  uint16_t num_fmt_id = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "0.00%", &num_fmt_id), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = 0;
  xf.num_fmt_id = num_fmt_id;
  uint32_t xf_idx = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);
  ASSERT_NE(xf_idx, 0U) << "non-default font/fill/numFmt must not dedup to the "
                           "all-zero placeholder xf at index 0";

  // Stability: re-adding the identical record must return the same index.
  uint32_t xf_idx_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx_again), 0);
  EXPECT_EQ(xf_idx, xf_idx_again);

  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 1, 0.4225), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 1, 1, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);

  uint32_t reread_xf = 0;
  ASSERT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 1, 1, &reread_xf), 0);
  ASSERT_EQ(reread_xf, xf_idx);

  fm_cell_xf reloaded{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb2.handle, reread_xf, &reloaded), 0);
  EXPECT_EQ(reloaded.font_index, font_idx);
  EXPECT_EQ(reloaded.fill_index, fill_idx);
  EXPECT_EQ(reloaded.num_fmt_id, num_fmt_id);

  fm_font_record reloaded_font{};
  ASSERT_EQ(fm_styles_get_font(wb2.handle, reloaded.font_index, &reloaded_font), 0);
  EXPECT_STREQ(reloaded_font.name, "Arial");
  EXPECT_DOUBLE_EQ(reloaded_font.size, 12.0);
  EXPECT_EQ(reloaded_font.bold, 1);

  fm_fill_record reloaded_fill{};
  ASSERT_EQ(fm_styles_get_fill(wb2.handle, reloaded.fill_index, &reloaded_fill), 0);
  EXPECT_EQ(reloaded_fill.pattern, 1U);
  EXPECT_EQ(reloaded_fill.fg_argb, 0xFFFF0000U);

  const char* reloaded_fmt = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb2.handle, reloaded.num_fmt_id, &reloaded_fmt), 0);
  ASSERT_NE(reloaded_fmt, nullptr);
  EXPECT_STREQ(reloaded_fmt, "0.00%");
}

TEST(FormulonCApiStyles, JustifyLastLineRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf_ex xf{};
  xf.base.horizontal_align = 7;  // distributed
  xf.base.vertical_align = 2;    // bottom
  xf.justify_last_line = 1;
  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf_ex(wb.handle, xf, &xf_idx), 0);

  uint32_t duplicate_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf_ex(wb.handle, xf, &duplicate_idx), 0);
  EXPECT_EQ(duplicate_idx, xf_idx);

  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "distributed"), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  uint32_t reread_idx = 0;
  ASSERT_EQ(fm_cell_get_xf_index(reloaded.handle, 0, 0, 0, &reread_idx), 0);
  fm_cell_xf_ex reread{};
  ASSERT_EQ(fm_styles_get_cell_xf_ex(reloaded.handle, reread_idx, &reread), 0);
  EXPECT_EQ(reread.base.horizontal_align, 7U);
  EXPECT_EQ(reread.justify_last_line, 1);
}

TEST(FormulonCApiStyles, NamedCellStyleRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf_ex style_xf{};
  style_xf.base.horizontal_align = 2;
  uint32_t xf_id = 0;
  ASSERT_EQ(fm_styles_add_cell_style_xf_ex(wb.handle, style_xf, &xf_id), 0);
  ASSERT_EQ(fm_styles_set_cell_style(wb.handle, "Highlight", xf_id, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);

  fm_cell_xf_ex cell_xf{};
  cell_xf.base.horizontal_align = 2;
  cell_xf.xf_id = xf_id;
  uint32_t cell_xf_id = 0;
  ASSERT_EQ(fm_styles_add_cell_xf_ex(wb.handle, cell_xf, &cell_xf_id), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "styled"), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, cell_xf_id), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  uint32_t style_count = 0;
  ASSERT_EQ(fm_styles_get_cell_style_count(loaded.handle, &style_count), 0);
  ASSERT_EQ(style_count, 1U);
  fm_cell_style_record_t style{};
  ASSERT_EQ(fm_styles_get_cell_style(loaded.handle, 0, &style), 0);
  EXPECT_STREQ(style.name, "Highlight");
  uint32_t reread_cell_xf = 0;
  ASSERT_EQ(fm_cell_get_xf_index(loaded.handle, 0, 0, 0, &reread_cell_xf), 0);
  fm_cell_xf_ex reread{};
  ASSERT_EQ(fm_styles_get_cell_xf_ex(loaded.handle, reread_cell_xf, &reread), 0);
  EXPECT_EQ(reread.xf_id, style.xf_id);
}

TEST(FormulonCApiStyles, NamedStyleXfRejectsDanglingReferences) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // A cell xf may only inherit from a named-style xf that already exists.
  fm_cell_xf_ex cell_xf{};
  cell_xf.xf_id = 3;
  uint32_t cell_xf_index = 0;
  EXPECT_NE(fm_styles_add_cell_xf_ex(wb.handle, cell_xf, &cell_xf_index), 0);

  // The named-style table validates its own font / fill / border / numFmt
  // references exactly like the cell-xf table does.
  fm_cell_xf_ex style_xf{};
  style_xf.base.font_index = 9;
  uint32_t xf_id = 0;
  EXPECT_NE(fm_styles_add_cell_style_xf_ex(wb.handle, style_xf, &xf_id), 0);

  style_xf.base.font_index = 0;
  style_xf.justify_last_line = 1;
  ASSERT_EQ(fm_styles_add_cell_style_xf_ex(wb.handle, style_xf, &xf_id), 0);
  fm_cell_xf_ex reread{};
  ASSERT_EQ(fm_styles_get_cell_style_xf_ex(wb.handle, xf_id, &reread), 0);
  EXPECT_EQ(reread.justify_last_line, 1);

  // 0..47 is the whole OOXML ordinal space; anything else needs the sentinel.
  EXPECT_NE(fm_styles_set_cell_style(wb.handle, "Custom", xf_id, 48), 0);
  EXPECT_EQ(fm_styles_set_cell_style(wb.handle, "Custom", xf_id, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);
  EXPECT_NE(fm_styles_set_cell_style(wb.handle, "Custom", xf_id + 1U, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);
}

TEST(FormulonCApiStyles, AddBatchDeduplicatesAllStyleTables) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const fm_font_record fonts[] = {MakeArial(), MakeArial()};
  const fm_fill_record fills[] = {MakeRedFill(), MakeRedFill()};
  fm_border_record border{};
  border.left.style = 1;
  border.left.color_argb = 0xFF000000U;
  const fm_border_record borders[] = {border, border};
  uint32_t font_indices[2]{};
  uint32_t fill_indices[2]{};
  uint32_t border_indices[2]{};
  const char* const num_fmt_codes[] = {"0.000%", "0.000%"};
  uint16_t num_fmt_ids[2]{};
  fm_cell_xf xf{};
  xf.font_index = 1U;
  xf.fill_index = 1U;
  xf.border_index = 1U;
  const fm_cell_xf xfs[] = {xf, xf};
  uint32_t xf_indices[2]{};
  const fm_styles_batch batch{fonts, 2U, font_indices, fills,         2U, fill_indices, borders, 2U, border_indices,
                              xfs,   2U, xf_indices,   num_fmt_codes, 2U, num_fmt_ids};

  ASSERT_EQ(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_indices[0], font_indices[1]);
  EXPECT_EQ(fill_indices[0], fill_indices[1]);
  EXPECT_EQ(border_indices[0], border_indices[1]);
  EXPECT_EQ(xf_indices[0], xf_indices[1]);
  EXPECT_EQ(num_fmt_ids[0], num_fmt_ids[1]);
}

TEST(FormulonCApiStyles, AddBatchCommitsNumFmtReferencesTransactionally) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const fm_font_record font = MakeArial();
  uint32_t font_index = 0;
  const char* const codes[] = {"0.000%"};
  uint16_t num_fmt_ids[] = {0xBEEF};
  fm_cell_xf xf{};
  xf.font_index = 1U;  // default root plus the staged font
  xf.num_fmt_id = 164U;
  uint32_t xf_indices[] = {0xDEADBEEFU};
  const fm_styles_batch batch{&font,   1U,  &font_index, nullptr,    0U,    nullptr, nullptr,    0U,
                              nullptr, &xf, 1U,          xf_indices, codes, 1U,      num_fmt_ids};

  ASSERT_EQ(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_index, 1U);
  EXPECT_EQ(num_fmt_ids[0], 164U);
  EXPECT_EQ(xf_indices[0], 1U);
  ASSERT_EQ(wb.handle->workbook().styles().num_fmts.size(), 1U);
  EXPECT_EQ(wb.handle->workbook().styles().num_fmts[0].id, 164U);
  ASSERT_EQ(wb.handle->workbook().styles().cell_xfs.size(), 2U);
  EXPECT_EQ(wb.handle->workbook().styles().cell_xfs[1].num_fmt_id, 164U);
}

TEST(FormulonCApiStyles, AddBatchFailureLeavesTableAndOutputsUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const std::string before_xml = formulon::io::write_styles(wb.handle->workbook().styles());
  const std::size_t before_fonts = wb.handle->workbook().styles().fonts.size();
  const std::size_t before_num_fmts = wb.handle->workbook().styles().num_fmts.size();
  const std::size_t before_cell_xfs = wb.handle->workbook().styles().cell_xfs.size();

  const fm_font_record font = MakeArial();
  uint32_t font_index = 0xA1A2A3A4U;
  const char* const codes[] = {"0.000%"};
  uint16_t num_fmt_ids[] = {0xBEEF};
  fm_cell_xf xf{};
  xf.font_index = 1U;    // valid after staging the font
  xf.num_fmt_id = 999U;  // neither built-in nor registered
  uint32_t xf_indices[] = {0xDEADBEEFU};
  const fm_styles_batch batch{&font,   1U,  &font_index, nullptr,    0U,    nullptr, nullptr,    0U,
                              nullptr, &xf, 1U,          xf_indices, codes, 1U,      num_fmt_ids};

  EXPECT_NE(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_index, 0xA1A2A3A4U);
  EXPECT_EQ(num_fmt_ids[0], 0xBEEFU);
  EXPECT_EQ(xf_indices[0], 0xDEADBEEFU);
  EXPECT_EQ(wb.handle->workbook().styles().fonts.size(), before_fonts);
  EXPECT_EQ(wb.handle->workbook().styles().num_fmts.size(), before_num_fmts);
  EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before_cell_xfs);
  EXPECT_EQ(formulon::io::write_styles(wb.handle->workbook().styles()), before_xml);
}

TEST(FormulonCApiStyles, CellXfEx2PreservesOptionalAlignmentAndPresence) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf_ex2 xf{};
  xf.base.vertical_align = 2;
  xf.justify_last_line = 1;
  xf.xf_id = 0;
  xf.has_alignment = 1;
  xf.has_text_rotation = 1;
  xf.text_rotation = 255;
  xf.has_indent = 1;
  xf.indent = 7;
  xf.has_relative_indent = 1;
  xf.relative_indent = -3;
  xf.has_shrink_to_fit = 1;
  xf.shrink_to_fit = 0;
  xf.has_reading_order = 1;
  xf.reading_order = 2;
  xf.has_horizontal_align = 1;
  xf.has_vertical_align = 1;
  xf.has_wrap_text = 1;
  xf.has_justify_last_line = 1;

  uint32_t index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf_ex2(wb.handle, xf, &index), 0);
  uint32_t duplicate = 0;
  ASSERT_EQ(fm_styles_add_cell_xf_ex2(wb.handle, xf, &duplicate), 0);
  EXPECT_EQ(duplicate, index);

  fm_cell_xf_ex2 reread{};
  ASSERT_EQ(fm_styles_get_cell_xf_ex2(wb.handle, index, &reread), 0);
  EXPECT_EQ(reread.justify_last_line, 1);
  EXPECT_EQ(reread.xf_id, 0U);
  EXPECT_EQ(reread.has_alignment, 1);
  EXPECT_EQ(reread.has_text_rotation, 1);
  EXPECT_EQ(reread.text_rotation, 255U);
  EXPECT_EQ(reread.has_indent, 1);
  EXPECT_EQ(reread.indent, 7U);
  EXPECT_EQ(reread.has_relative_indent, 1);
  EXPECT_EQ(reread.relative_indent, -3);
  EXPECT_EQ(reread.has_shrink_to_fit, 1);
  EXPECT_EQ(reread.shrink_to_fit, 0);
  EXPECT_EQ(reread.has_reading_order, 1);
  EXPECT_EQ(reread.reading_order, 2U);
  EXPECT_EQ(reread.has_horizontal_align, 1);
  EXPECT_EQ(reread.has_vertical_align, 1);
  EXPECT_EQ(reread.has_wrap_text, 1);
  EXPECT_EQ(reread.has_justify_last_line, 1);

  // Presence is part of the dedup key: explicit zero / false differs from
  // an omitted attribute even when the effective value is the same.
  fm_cell_xf_ex2 absent = xf;
  absent.has_text_rotation = 0;
  absent.has_indent = 0;
  absent.has_relative_indent = 0;
  absent.has_shrink_to_fit = 0;
  absent.has_reading_order = 0;
  uint32_t absent_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf_ex2(wb.handle, absent, &absent_index), 0);
  EXPECT_NE(absent_index, index);
}

TEST(FormulonCApiStyles, CellXfEx2RejectsInvalidExcelAlignmentRanges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf_ex2 xf{};
  uint32_t index = 0;

  xf.has_text_rotation = 1;
  xf.text_rotation = 181;
  EXPECT_NE(fm_styles_add_cell_xf_ex2(wb.handle, xf, &index), 0);
  xf.text_rotation = 255;
  EXPECT_EQ(fm_styles_add_cell_xf_ex2(wb.handle, xf, &index), 0);

  xf.has_text_rotation = 0;
  xf.has_indent = 1;
  xf.indent = 256;
  EXPECT_NE(fm_styles_add_cell_xf_ex2(wb.handle, xf, &index), 0);
  xf.indent = 255;
  EXPECT_EQ(fm_styles_add_cell_xf_ex2(wb.handle, xf, &index), 0);

  xf.has_indent = 0;
  xf.has_reading_order = 1;
  xf.reading_order = 3;
  EXPECT_NE(fm_styles_add_cell_xf_ex2(wb.handle, xf, &index), 0);
}

TEST(FormulonCApiStyles, CellXfEx2PresenceFlagsDistinguishExplicitDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf_ex2 explicit_defaults{};
  explicit_defaults.base.vertical_align = 2;
  explicit_defaults.has_alignment = 1;
  explicit_defaults.has_horizontal_align = 1;
  explicit_defaults.has_vertical_align = 1;
  explicit_defaults.has_wrap_text = 1;
  explicit_defaults.has_justify_last_line = 1;
  uint32_t explicit_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf_ex2(wb.handle, explicit_defaults, &explicit_index), 0);

  fm_cell_xf_ex2 omitted_defaults = explicit_defaults;
  omitted_defaults.has_horizontal_align = 0;
  omitted_defaults.has_vertical_align = 0;
  omitted_defaults.has_wrap_text = 0;
  omitted_defaults.has_justify_last_line = 0;
  uint32_t omitted_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf_ex2(wb.handle, omitted_defaults, &omitted_index), 0);
  EXPECT_NE(explicit_index, omitted_index);

  fm_cell_xf_ex2 reread{};
  ASSERT_EQ(fm_styles_get_cell_xf_ex2(wb.handle, explicit_index, &reread), 0);
  EXPECT_EQ(reread.has_horizontal_align, 1);
  EXPECT_EQ(reread.has_vertical_align, 1);
  EXPECT_EQ(reread.has_wrap_text, 1);
  EXPECT_EQ(reread.has_justify_last_line, 1);
}

TEST(FormulonCApiStyles, CellStyleXfEx2RoundTripsOptionalAlignment) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf_ex2 style{};
  style.has_text_rotation = 1;
  style.text_rotation = 0;
  style.has_relative_indent = 1;
  style.relative_indent = -2;
  style.has_shrink_to_fit = 1;
  style.shrink_to_fit = false;
  uint32_t style_index = 0;
  ASSERT_EQ(fm_styles_add_cell_style_xf_ex2(wb.handle, style, &style_index), 0);

  fm_cell_xf_ex2 reread{};
  ASSERT_EQ(fm_styles_get_cell_style_xf_ex2(wb.handle, style_index, &reread), 0);
  EXPECT_EQ(reread.has_text_rotation, 1);
  EXPECT_EQ(reread.text_rotation, 0U);
  EXPECT_EQ(reread.has_relative_indent, 1);
  EXPECT_EQ(reread.relative_indent, -2);
  EXPECT_EQ(reread.has_shrink_to_fit, 1);
  EXPECT_EQ(reread.shrink_to_fit, 0);
}
