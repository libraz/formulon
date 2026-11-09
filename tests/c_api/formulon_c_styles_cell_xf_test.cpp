// Stable C ABI cell XF index and mutation tests.

#include "formulon_c_out_index.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, CellXfIndexDefaultsToZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t xf = 99;
  EXPECT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, 0, 0, &xf), 0);
  EXPECT_EQ(xf, 0U);
}

TEST(FormulonCApiStyles, CellXfIndexRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(43);
  EXPECT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 5, 7, 42), 0);
  std::uint32_t xf = 0;
  EXPECT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, 5, 7, &xf), 0);
  EXPECT_EQ(xf, 42U);
}

TEST(FormulonCApiStyles, SetXfIndexRejectsBadSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // sheet 99 does not exist.
  EXPECT_NE(fm_sheet_set_cell_xf_index(wb.handle, 99, 0, 0, 1), 0);
}

TEST(FormulonCApiStyles, NullArgumentRejected) {
  std::uint32_t xf = 0;
  EXPECT_NE(fm_sheet_get_cell_xf_index(nullptr, 0, 0, 0, &xf), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_NE(fm_sheet_get_cell_xf_index(wb.handle, 0, 0, 0, nullptr), 0);
}

TEST(FormulonCApiStyles, GetCellXfRejectsOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_xf_t xf{};
  // A fresh workbook seeds one default xf, so index 0 resolves and index
  // 1 is the first out-of-range one.
  EXPECT_EQ(fm_workbook_get_cell_xf_at(wb.handle, 0, &xf), 0);
  EXPECT_NE(fm_workbook_get_cell_xf_at(wb.handle, 1, &xf), 0);
}

TEST(FormulonCApiStyles, CellXfGetterRecordsDiagnostics) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  auto expect_diagnostic = [](fm_status_t status, const char* field) {
    EXPECT_NE(status, 0);
    const char* message = fm_get_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_STRNE(message, "");
    EXPECT_EQ(std::string(message).find("fm_"), std::string::npos) << message;
    if (field != nullptr) {
      EXPECT_NE(std::string(fm_get_last_error_context()).find(field), std::string::npos) << fm_get_last_error_context();
    }
  };

  // Index 1 is past both seeded single-entry tables, so each getter takes
  // its out-of-range path and must record a diagnostic.
  fm_cell_xf_t cell_xf{};
  expect_diagnostic(fm_workbook_get_cell_xf_at(wb.handle, 1, &cell_xf), "index");
  expect_diagnostic(fm_workbook_get_cell_style_xf_at(wb.handle, 1, &cell_xf), "index");
}

TEST(FormulonCApiStyles, CellXfAdderRecordsDiagnostics) {
  auto expect_diagnostic = [](fm_status_t status, const char* field) {
    EXPECT_NE(status, 0);
    const char* message = fm_get_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_STRNE(message, "");
    EXPECT_EQ(std::string(message).find("fm_"), std::string::npos) << message;
    if (field != nullptr) {
      EXPECT_NE(std::string(fm_get_last_error_context()).find(field), std::string::npos) << fm_get_last_error_context();
    }
  };

  std::uint32_t index = 0;
  fm_cell_xf_t cell_xf{};
  expect_diagnostic(fm_workbook_add_cell_xf(nullptr, &cell_xf, &index), nullptr);

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  cell_xf.font_index = 1U;
  expect_diagnostic(fm_workbook_add_cell_xf(wb.handle, &cell_xf, &index), "font_index");

  cell_xf = fm_cell_xf_t{};
  cell_xf.style_xf_index = 1U;
  expect_diagnostic(fm_workbook_add_cell_xf(wb.handle, &cell_xf, &index), "style_xf_index");

  cell_xf = fm_cell_xf_t{};
  cell_xf.has_text_rotation = 1;
  cell_xf.text_rotation = 181U;
  expect_diagnostic(fm_workbook_add_cell_xf(wb.handle, &cell_xf, &index), "text_rotation");
}

TEST(FormulonCApiStyles, SetThenSaveLoadPreservesXfIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  // Register the xf records the stamp below names. A `<c s="7">` against
  // a shorter `<cellXfs>` resolves to no style, so the reader falls back
  // to the default rather than handing back an index whose record
  // `fm_workbook_get_cell_xf_at` would then refuse to return.
  wb.handle->workbook().mutable_styles().cell_xfs.resize(8);

  // Stamp xf_index = 7 on cell A1 and save.
  EXPECT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 3.14), 0);
  EXPECT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, 7), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  std::uint32_t xf = 0;
  EXPECT_EQ(fm_sheet_get_cell_xf_index(wb2.handle, 0, 0, 0, &xf), 0);
  EXPECT_EQ(xf, 7U);
}

TEST(FormulonCApiStyles, AddCellXfDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t font_index = 0;
  std::size_t fill_index = 0;
  std::size_t border_index = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &font_index), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &fill_index), 0);
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &border_index), 0);

  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(font_index);
  xf.fill_index = static_cast<uint32_t>(fill_index);
  xf.border_index = static_cast<uint32_t>(border_index);
  xf.num_fmt_id = 0;  // built-in General
  xf.horizontal = 1;
  xf.vertical = 2;
  xf.wrap_text = 1;

  std::uint32_t a = 0xFFFFFFFFU;
  std::uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &a), 0);
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddCellXfDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t font_index = 0;
  std::size_t fill_index = 0;
  std::size_t border_index = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &font_index), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &fill_index), 0);
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &border_index), 0);

  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(font_index);
  xf.fill_index = static_cast<uint32_t>(fill_index);
  xf.border_index = static_cast<uint32_t>(border_index);
  xf.num_fmt_id = 0;
  xf.horizontal = 1;
  xf.vertical = 2;
  xf.wrap_text = 1;

  xf.has_wrap_text = 1;

  std::uint32_t a = 0;
  std::uint32_t b = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &a), 0);
  // Flip wrap_text — distinct record.
  xf.wrap_text = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddCellXfGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t font_index = 0;
  std::size_t fill_index = 0;
  std::size_t border_index = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &font_index), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &fill_index), 0);
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &border_index), 0);
  std::size_t before = 0;
  ASSERT_EQ(fm_workbook_get_cell_xf_count(wb.handle, &before), 0);
  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(font_index);
  xf.fill_index = static_cast<uint32_t>(fill_index);
  xf.border_index = static_cast<uint32_t>(border_index);
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it. The
  // presence flag is what makes the value part of the record.
  xf.wrap_text = 1;
  xf.has_wrap_text = 1;
  std::uint32_t idx = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &idx), 0);
  std::size_t after = 0;
  ASSERT_EQ(fm_workbook_get_cell_xf_count(wb.handle, &after), 0);
  // Index 0 is the seeded default xf.
  EXPECT_EQ(before, 1U);
  EXPECT_EQ(after, 2U);
  EXPECT_EQ(idx, 1U);
}

TEST(FormulonCApiStyles, AddCellXfOnFreshWorkbookKeepsZeroAsDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t font_index = 0;
  std::size_t fill_index = 0;
  std::size_t border_index = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &font_index), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &fill_index), 0);
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &border_index), 0);

  // A fresh workbook seeds Excel's reserved slots, so the caller's first
  // record lands after them: one default font, one empty border, and the
  // two fills (`none`, `gray125`) Excel always writes first.
  EXPECT_EQ(font_index, 1U);
  EXPECT_EQ(fill_index, 2U);
  EXPECT_EQ(border_index, 1U);

  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(font_index);
  xf.fill_index = static_cast<uint32_t>(fill_index);
  xf.border_index = static_cast<uint32_t>(border_index);
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it. The
  // presence flag is what makes the value part of the record.
  xf.wrap_text = 1;
  xf.has_wrap_text = 1;

  std::uint32_t xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &xf, &xf_index), 0);
  EXPECT_NE(xf_index, 0U);

  fm_cell_xf_t default_xf{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(wb.handle, 0, &default_xf), 0);
  EXPECT_EQ(default_xf.font_index, 0U);
  EXPECT_EQ(default_xf.fill_index, 0U);
  EXPECT_EQ(default_xf.border_index, 0U);
  EXPECT_EQ(default_xf.num_fmt_id, 0U);
  EXPECT_EQ(default_xf.wrap_text, 0);
}

TEST(FormulonCApiStyles, AddCellXfRejectsOutOfRangeIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // Empty styles table: any non-zero index is out of range.
  fm_cell_xf_t xf{};
  xf.font_index = 5;  // intentionally OOR
  xf.fill_index = 0;
  xf.border_index = 0;
  xf.num_fmt_id = 0;
  std::uint32_t idx = 0xFFFFFFFFU;
  EXPECT_NE(fm_workbook_add_cell_xf(wb.handle, &xf, &idx), 0);
}

TEST(FormulonCApiStyles, AddCellXfRejectsUnknownNumFmt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t font_index = 0;
  std::size_t fill_index = 0;
  std::size_t border_index = 0;
  ASSERT_EQ(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), &font_index), 0);
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, record_arg(MakeRedFill()), &fill_index), 0);
  ASSERT_EQ(fm_workbook_add_border(wb.handle, record_arg(MakeThinBoxBorder()), &border_index), 0);

  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(font_index);
  xf.fill_index = static_cast<uint32_t>(fill_index);
  xf.border_index = static_cast<uint32_t>(border_index);
  xf.num_fmt_id = 200;  // not a registered custom and not a documented built-in
  std::uint32_t idx = 0xFFFFFFFFU;
  EXPECT_NE(fm_workbook_add_cell_xf(wb.handle, &xf, &idx), 0);
}

TEST(FormulonCApiStyles, AddNullArgumentRejected) {
  std::size_t idx = 0;
  std::uint32_t dxf_idx = 0;
  EXPECT_NE(fm_workbook_add_font(nullptr, record_arg(MakeArial()), &idx), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_NE(fm_workbook_add_font(wb.handle, record_arg(MakeArial()), nullptr), 0);
  uint32_t id = 0;
  EXPECT_NE(fm_workbook_add_number_format(nullptr, "General", &id), 0);
  EXPECT_NE(fm_workbook_add_number_format(wb.handle, "General", nullptr), 0);
  EXPECT_NE(fm_workbook_add_dxf(nullptr, record_arg(fm_dxf_t{}), &dxf_idx), 0);
  EXPECT_NE(fm_workbook_add_dxf(wb.handle, record_arg(fm_dxf_t{}), nullptr), 0);
}

TEST(FormulonCApiStyles, NullInputRecordRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_status_t kNull = FM_STATUS_BINDING_NULL_POINTER;
  std::size_t idx = 0;
  std::uint32_t xf_idx = 0;
  uint32_t argb = 0;
  int32_t resolution = 0;
  size_t merges_found = 0;
  EXPECT_EQ(fm_workbook_add_font(wb.handle, nullptr, &idx), kNull);
  EXPECT_EQ(fm_workbook_set_font(wb.handle, 0, nullptr), kNull);
  EXPECT_EQ(fm_workbook_set_default_font(wb.handle, nullptr), kNull);
  EXPECT_EQ(fm_workbook_add_fill(wb.handle, nullptr, &idx), kNull);
  EXPECT_EQ(fm_workbook_add_border(wb.handle, nullptr, &idx), kNull);
  EXPECT_EQ(fm_workbook_add_cell_xf(wb.handle, nullptr, &xf_idx), kNull);
  EXPECT_EQ(fm_workbook_add_cell_style_xf(wb.handle, nullptr, &xf_idx), kNull);
  EXPECT_EQ(fm_workbook_add_dxf(wb.handle, nullptr, &xf_idx), kNull);
  EXPECT_EQ(fm_workbook_resolve_color(wb.handle, nullptr, 0, &argb, &resolution), kNull);
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, nullptr, &idx), kNull);
  EXPECT_EQ(fm_sheet_set_hyperlink(nullptr, 0, 0, 0, nullptr), kNull);
  EXPECT_EQ(fm_sheet_add_merge(wb.handle, 0, nullptr), kNull);
  EXPECT_EQ(fm_sheet_remove_merge(wb.handle, 0, nullptr), kNull);
  EXPECT_EQ(fm_sheet_get_merges_in_range(wb.handle, 0, nullptr, nullptr, 0, &merges_found), kNull);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, nullptr, ignored_index()), kNull);
}

namespace {

// A record that sets every apply flag, quotePrefix and an explicit
// `<protection>` child whose values differ from the schema defaults.
fm_cell_xf_t MakeFlaggedXf() {
  fm_cell_xf_t xf{};
  xf.apply_number_format = 1;
  xf.apply_font = 1;
  xf.apply_fill = 1;
  xf.apply_border = 1;
  xf.apply_alignment = 1;
  xf.apply_protection = 1;
  xf.quote_prefix = 1;
  xf.has_protection = 1;
  xf.locked = 0;
  xf.hidden = 1;
  return xf;
}

void ExpectFlagged(const fm_cell_xf_t& xf) {
  EXPECT_EQ(xf.apply_number_format, 1);
  EXPECT_EQ(xf.apply_font, 1);
  EXPECT_EQ(xf.apply_fill, 1);
  EXPECT_EQ(xf.apply_border, 1);
  EXPECT_EQ(xf.apply_alignment, 1);
  EXPECT_EQ(xf.apply_protection, 1);
  EXPECT_EQ(xf.quote_prefix, 1);
  EXPECT_EQ(xf.has_protection, 1);
  EXPECT_EQ(xf.locked, 0);
  EXPECT_EQ(xf.hidden, 1);
}

}  // namespace

TEST(FormulonCApiStyles, XfApplyFlagsRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, record_arg(MakeFlaggedXf()), &xf_index), 0);
  ASSERT_NE(xf_index, 0U);
  std::uint32_t style_xf_id = 0;
  ASSERT_EQ(fm_workbook_add_cell_style_xf(wb.handle, record_arg(MakeFlaggedXf()), &style_xf_id), 0);

  fm_cell_xf_t got{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(wb.handle, xf_index, &got), 0);
  ExpectFlagged(got);

  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, xf_index), 0);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  EXPECT_NE(styles_xml.find("applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" "
                            "applyAlignment=\"1\" applyProtection=\"1\" quotePrefix=\"1\""),
            std::string::npos)
      << styles_xml;
  EXPECT_NE(styles_xml.find("<protection locked=\"0\" hidden=\"1\"/>"), std::string::npos) << styles_xml;

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  std::uint32_t reloaded_index = 0;
  ASSERT_EQ(fm_sheet_get_cell_xf_index(reloaded.handle, 0, 0, 0, &reloaded_index), 0);
  fm_cell_xf_t after{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(reloaded.handle, reloaded_index, &after), 0);
  ExpectFlagged(after);
  fm_cell_xf_t style_after{};
  ASSERT_EQ(fm_workbook_get_cell_style_xf_at(reloaded.handle, style_xf_id, &style_after), 0);
  ExpectFlagged(style_after);
}

TEST(FormulonCApiStyles, XfProtectionDistinguishesDedup) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_xf_t locked{};
  locked.has_protection = 1;
  locked.locked = 1;
  fm_cell_xf_t unlocked = locked;
  unlocked.locked = 0;
  std::uint32_t a = 0;
  std::uint32_t b = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &locked, &a), 0);
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &unlocked, &b), 0);
  EXPECT_NE(a, b);

  // Without the child, `locked` / `hidden` are ignored: the record
  // deduplicates with the default XF and reads back the schema defaults.
  fm_cell_xf_t absent{};
  absent.locked = 0;
  absent.hidden = 1;
  std::uint32_t c = 99;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &absent, &c), 0);
  EXPECT_EQ(c, 0U);
  fm_cell_xf_t got{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(wb.handle, c, &got), 0);
  EXPECT_EQ(got.has_protection, 0);
  EXPECT_EQ(got.locked, 1);
  EXPECT_EQ(got.hidden, 0);
}

TEST(FormulonCApiStyles, ResetXfIndexRestoresColumnInheritance) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 0, 0, 1), 0);

  fm_effective_style_info_t style = FM_STRUCT_INIT(fm_effective_style_info_t);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, 2), 0);
  ASSERT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 0, 0, &style), 0);
  EXPECT_EQ(style.source, 0);
  EXPECT_EQ(style.xf_index, 2U);

  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, 0, 0), 0);
  ASSERT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 0, 0, &style), 0);
  EXPECT_EQ(style.source, 2);
  EXPECT_EQ(style.xf_index, 1U);

  // A direct Normal assignment is not the same state as a cleared one.
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, 0), 0);
  ASSERT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 0, 0, &style), 0);
  EXPECT_EQ(style.source, 0);
  EXPECT_EQ(style.xf_index, 0U);

  // The cleared cell follows a later column restyle.
  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, 0, 0), 0);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 0, 0, 2), 0);
  ASSERT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 0, 0, &style), 0);
  EXPECT_EQ(style.source, 2);
  EXPECT_EQ(style.xf_index, 2U);
}

TEST(FormulonCApiStyles, ResetXfIndexRestoresRowAndDefaultInheritance) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3);
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 4, 4, 1), 0);
  fm_effective_style_info_t style = FM_STRUCT_INIT(fm_effective_style_info_t);

  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 4, 3, 2), 0);
  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, 4, 3), 0);
  ASSERT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 4, 3, &style), 0);
  EXPECT_EQ(style.source, 1);
  EXPECT_EQ(style.xf_index, 1U);

  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 7, 3, 2), 0);
  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, 7, 3), 0);
  ASSERT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 7, 3, &style), 0);
  EXPECT_EQ(style.source, 3);
  EXPECT_EQ(style.xf_index, 0U);
}

TEST(FormulonCApiStyles, ResetXfIndexKeepsValueAndFormula) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 5.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=A1*2"), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, 2), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 1, 2), 0);
  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, 0, 0), 0);
  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, 0, 1), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  const formulon::Sheet& sheet = wb.handle->workbook().sheet(0);
  ASSERT_NE(sheet.cell_at(0, 0), nullptr);
  EXPECT_EQ(sheet.cell_at(0, 0)->cached_value.as_number(), 5.0);
  ASSERT_NE(sheet.cell_at(0, 1), nullptr);
  EXPECT_EQ(sheet.cell_at(0, 1)->formula_text, "=A1*2");
  EXPECT_EQ(sheet.cell_at(0, 1)->cached_value.as_number(), 10.0);
  // A value cell owns its style, so a cleared one is a cell-sourced Normal.
  fm_effective_style_info_t style = FM_STRUCT_INIT(fm_effective_style_info_t);
  ASSERT_EQ(fm_sheet_get_effective_style(wb.handle, 0, 0, 0, &style), 0);
  EXPECT_EQ(style.source, 0);
  EXPECT_EQ(style.xf_index, 0U);
}

TEST(FormulonCApiStyles, ResetXfIndexOnAbsentCellIsNoOp) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, 9, 9), 0);
  EXPECT_EQ(wb.handle->workbook().sheet(0).cell_at(9, 9), nullptr);
}

TEST(FormulonCApiStyles, ResetXfIndexRejectsBadArguments) {
  EXPECT_EQ(fm_sheet_reset_cell_xf_index(nullptr, 0, 0, 0),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 99, 0, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, formulon::Sheet::kMaxRows, 0), invalid);
}

TEST(FormulonCApiStyles, ResetCellKeepsInheritanceAcrossSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(3);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 0, 0, 1), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0, 0, 2), 0);
  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb.handle, 0, 0, 0), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  fm_effective_style_info_t style = FM_STRUCT_INIT(fm_effective_style_info_t);
  ASSERT_EQ(fm_sheet_get_effective_style(reloaded.handle, 0, 0, 0, &style), 0);
  EXPECT_EQ(style.source, 2);
  EXPECT_EQ(style.xf_index, 1U);
}

TEST(FormulonCApiStyles, AddCellXfDoesNotReuseARecordWithUnmodelledAttributes) {
  // A loaded record that differs from the request only in an attribute the
  // model does not represent must not be handed back as a match.
  for (const bool pivot_button : {true, false}) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    auto& xfs = wb.handle->workbook().mutable_styles().cell_xfs;
    formulon::CellXf loaded = xfs[0];
    loaded.num_fmt_id = 2U;
    loaded.apply_number_format = true;
    loaded.extra_attrs = pivot_button ? " pivotButton=\"1\"" : "";
    xfs.push_back(loaded);
    fm_cell_xf_t request{};
    ASSERT_EQ(fm_workbook_get_cell_xf_at(wb.handle, 1U, &request), 0);
    std::uint32_t index = 99U;
    ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &request, &index), 0) << fm_get_last_error_message();
    if (pivot_button) {
      EXPECT_EQ(index, 2U);
      EXPECT_TRUE(xfs[index].extra_attrs.empty());
    } else {
      EXPECT_EQ(index, 1U);
    }
  }
}
