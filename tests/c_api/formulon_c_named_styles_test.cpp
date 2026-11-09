//
// Stable C ABI tests for the named-cell-style surface
// (`fm_workbook_get_cell_style_count` / `fm_workbook_get_cell_style_at` and
// the parallel `fm_workbook_get_cell_style_xf_at*` accessors). The actual
// reader/writer round-trip lives in the integration suite; these tests
// exercise the boundary's null / range guards and the empty-workbook
// behaviour.

#include <cstdint>
#include <cstring>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

static_assert(sizeof(fm_cell_style_t) == (sizeof(void*) == 4U ? 24U : 32U), "fm_cell_style_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

/// `builtin_id` for a custom style, which carries no `builtinId`.
constexpr uint32_t kCustomStyle = UINT32_MAX;

fm_cell_style_t StyleRecord(const char* name, uint32_t xf_id, uint32_t builtin_id) {
  fm_cell_style_t record{};
  record.name = name;
  record.style_xf_index = xf_id;
  record.has_builtin_id = builtin_id != kCustomStyle ? 1 : 0;
  record.builtin_id = builtin_id != kCustomStyle ? builtin_id : 0U;
  return record;
}

/// Returns the index of the named style, or `UINT32_MAX` when absent.
uint32_t FindStyle(fm_workbook_t* wb, const char* name, fm_cell_style_t* out) {
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_get_cell_style_count(wb, &count), 0);
  for (uint32_t i = 0; i < count; ++i) {
    EXPECT_EQ(fm_workbook_get_cell_style_at(wb, i, out), 0);
    if (std::strcmp(out->name, name) == 0) {
      return i;
    }
  }
  return UINT32_MAX;
}

/// Adds a named-style xf distinct from Normal's (a bold font and a date format).
uint32_t AddBoldDateStyleXf(fm_workbook_t* wb, std::uint32_t* out_font) {
  fm_font_t font{};
  font.name = "Calibri";
  font.size_pt = 11.0;
  font.bold = 1;
  font.has_bold = 1;
  EXPECT_EQ(fm_workbook_add_font(wb, &font, out_font), 0);
  fm_cell_xf_t xf{};
  xf.font_index = static_cast<uint32_t>(*out_font);
  xf.num_fmt_id = 14;
  std::uint32_t xf_id = 0;
  EXPECT_EQ(fm_workbook_add_cell_style_xf(wb, &xf, &xf_id), 0);
  return static_cast<uint32_t>(xf_id);
}

}  // namespace

TEST(FormulonCApiNamedStyles, FreshWorkbookHasSeededNormalStyle) {
  // A new workbook seeds Excel's minimum style table, which includes the
  // `Normal` named style and its `<cellStyleXfs>` entry. Before that seed
  // the pair existed only in the serialized document, so a caller could not
  // reference it without a save/load cycle.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t cs_count = 99;
  std::size_t cs_xf_count = 99;
  EXPECT_EQ(fm_workbook_get_cell_style_count(wb.handle, &cs_count), 0);
  EXPECT_EQ(fm_workbook_get_cell_style_xf_count(wb.handle, &cs_xf_count), 0);
  EXPECT_EQ(cs_count, 1U);
  EXPECT_EQ(cs_xf_count, 1U);

  fm_cell_style_t normal{};
  ASSERT_EQ(fm_workbook_get_cell_style_at(wb.handle, 0, &normal), 0);
  ASSERT_NE(normal.name, nullptr);
  EXPECT_STREQ(normal.name, "Normal");
  EXPECT_EQ(normal.has_builtin_id, 1);
  EXPECT_EQ(normal.builtin_id, 0U);
}

TEST(FormulonCApiNamedStyles, IndexOutOfRangeReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_style_t cs{};
  fm_cell_xf_t xf{};
  // Index 0 is the seeded `Normal` style; index 1 is past both tables.
  EXPECT_EQ(fm_workbook_get_cell_style_at(wb.handle, 1, &cs),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  EXPECT_EQ(fm_workbook_get_cell_style_xf_at(wb.handle, 1, &xf),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
}

TEST(FormulonCApiNamedStyles, NullArgsReturnBindingNullPointer) {
  fm_cell_style_t cs{};
  fm_cell_xf_t xf{};
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_get_cell_style_count(nullptr, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_cell_style_xf_count(nullptr, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_cell_style_at(nullptr, 0, &cs),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_cell_style_xf_at(nullptr, 0, &xf),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_get_cell_style_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_cell_style_xf_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_cell_style_at(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_cell_style_xf_at(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiNamedStyles, RoundTripPreservesSeededNamedStyles) {
  // The seeded `Normal` pair must survive a save/load cycle unchanged -
  // the writer must not duplicate it while synthesizing its own default.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  std::size_t live_cs_count = 99;
  std::size_t live_cs_xf_count = 99;
  ASSERT_EQ(fm_workbook_get_cell_style_count(wb.handle, &live_cs_count), 0);
  ASSERT_EQ(fm_workbook_get_cell_style_xf_count(wb.handle, &live_cs_xf_count), 0);
  EXPECT_EQ(live_cs_count, 1U);
  EXPECT_EQ(live_cs_xf_count, 1U);

  uint8_t* saved_data = nullptr;
  size_t saved_len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved_data, &saved_len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved_data, saved_len, &wb2.handle), 0);
  fm_buffer_destroy(saved_data);

  std::size_t cs_count = 99;
  std::size_t cs_xf_count = 99;
  EXPECT_EQ(fm_workbook_get_cell_style_count(wb2.handle, &cs_count), 0);
  EXPECT_EQ(fm_workbook_get_cell_style_xf_count(wb2.handle, &cs_xf_count), 0);
  EXPECT_EQ(cs_count, 1U);
  EXPECT_EQ(cs_xf_count, 1U);

  fm_cell_style_t normal{};
  ASSERT_EQ(fm_workbook_get_cell_style_at(wb2.handle, 0, &normal), 0);
  ASSERT_NE(normal.name, nullptr);
  EXPECT_STREQ(normal.name, "Normal");
  EXPECT_EQ(normal.style_xf_index, 0U);
  EXPECT_EQ(normal.builtin_id, 0U);
  EXPECT_EQ(normal.outline_level, 0U);
  EXPECT_EQ(normal.hidden, 0);
  EXPECT_EQ(normal.custom_builtin, 0);

  fm_cell_xf_t normal_xf{};
  ASSERT_EQ(fm_workbook_get_cell_style_xf_at(wb2.handle, 0, &normal_xf), 0);
  EXPECT_EQ(normal_xf.font_index, 0U);
  EXPECT_EQ(normal_xf.fill_index, 0U);
  EXPECT_EQ(normal_xf.border_index, 0U);
  EXPECT_EQ(normal_xf.num_fmt_id, 0U);
  EXPECT_EQ(normal_xf.horizontal, 0);
  EXPECT_EQ(normal_xf.vertical, 2);
  EXPECT_EQ(normal_xf.wrap_text, 0);
}

TEST(FormulonCApiNamedStyles, BuiltinId53RoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t font = 0;
  const uint32_t xf_id = AddBoldDateStyleXf(wb.handle, &font);
  const fm_cell_style_t record = StyleRecord("Explanatory Text", xf_id, 53U);
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &record), 0);

  uint8_t* saved_data = nullptr;
  size_t saved_len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved_data, &saved_len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved_data, saved_len, &loaded.handle), 0);
  fm_buffer_destroy(saved_data);

  fm_cell_style_t reread{};
  ASSERT_NE(FindStyle(loaded.handle, "Explanatory Text", &reread), UINT32_MAX);
  EXPECT_EQ(reread.builtin_id, 53U);
}

TEST(FormulonCApiNamedStyles, SetCellStyleCarriesEveryFieldAndReplacesByName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t font = 0;
  const uint32_t xf_id = AddBoldDateStyleXf(wb.handle, &font);
  fm_cell_style_t record = StyleRecord("RowLevel_4", xf_id, 1U);
  record.outline_level = 3;
  record.hidden = 1;
  record.custom_builtin = 1;
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &record), 0);

  uint8_t* saved_data = nullptr;
  size_t saved_len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved_data, &saved_len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved_data, saved_len, &loaded.handle), 0);
  fm_buffer_destroy(saved_data);
  fm_cell_style_t reread{};
  ASSERT_NE(FindStyle(loaded.handle, "RowLevel_4", &reread), UINT32_MAX);
  EXPECT_EQ(reread.style_xf_index, xf_id);
  EXPECT_EQ(reread.builtin_id, 1U);
  EXPECT_EQ(reread.outline_level, 3U);
  EXPECT_EQ(reread.hidden, 1);
  EXPECT_EQ(reread.custom_builtin, 1);

  // Setting the same name again overwrites every field instead of appending.
  std::size_t before = 0;
  ASSERT_EQ(fm_workbook_get_cell_style_count(wb.handle, &before), 0);
  const fm_cell_style_t replacement = StyleRecord("RowLevel_4", 0U, kCustomStyle);
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &replacement), 0);
  std::size_t after = 0;
  ASSERT_EQ(fm_workbook_get_cell_style_count(wb.handle, &after), 0);
  EXPECT_EQ(after, before);
  ASSERT_NE(FindStyle(wb.handle, "RowLevel_4", &reread), UINT32_MAX);
  EXPECT_EQ(reread.style_xf_index, 0U);
  EXPECT_EQ(reread.has_builtin_id, 0) << "a custom style reads has_builtin_id 0";
  EXPECT_EQ(reread.builtin_id, 0U);
  EXPECT_EQ(reread.outline_level, 0U);
  EXPECT_EQ(reread.hidden, 0);
  EXPECT_EQ(reread.custom_builtin, 0);
}

TEST(FormulonCApiNamedStyles, SetCellStyleRejectsInvalidRecords) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_style_t record = StyleRecord("Styled", 0U, kCustomStyle);
  EXPECT_EQ(fm_workbook_set_cell_style(nullptr, &record), kNullPointer);
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, nullptr), kNullPointer);
  record.name = nullptr;
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &record), kNullPointer);
  record.name = "";
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &record), kInvalidArgument);

  record = StyleRecord("Styled", 1U, kCustomStyle);
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &record), kInvalidArgument);
  record = StyleRecord("Styled", 0U, 54U);
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &record), kInvalidArgument);

  // An outline level belongs to RowLevel_n / ColLevel_n only, and stops at 6.
  record = StyleRecord("ColLevel_7", 0U, 2U);
  record.outline_level = 6;
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &record), 0);
  record.outline_level = 7;
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &record), kInvalidArgument);
  record = StyleRecord("Heading 1", 0U, 16U);
  record.outline_level = 1;
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &record), kInvalidArgument);
  record.outline_level = 0;
  EXPECT_EQ(fm_workbook_set_cell_style(wb.handle, &record), 0);
}

TEST(FormulonCApiNamedStyles, RemoveCellStyleRepointsXfs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t bold = 0;
  const uint32_t xf_id = AddBoldDateStyleXf(wb.handle, &bold);
  const fm_cell_style_t record = StyleRecord("Dated", xf_id, kCustomStyle);
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &record), 0);

  // The font is the cell's own (apply_font); the number format is inherited.
  fm_cell_xf_t cell_xf{};
  cell_xf.style_xf_index = xf_id;
  cell_xf.font_index = static_cast<uint32_t>(bold);
  cell_xf.num_fmt_id = 14;
  cell_xf.apply_font = 1;
  std::uint32_t cell_xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &cell_xf, &cell_xf_index), 0);

  std::size_t xf_count = 0;
  std::size_t style_xf_count = 0;
  std::size_t style_count = 0;
  ASSERT_EQ(fm_workbook_get_cell_xf_count(wb.handle, &xf_count), 0);
  ASSERT_EQ(fm_workbook_get_cell_style_xf_count(wb.handle, &style_xf_count), 0);
  ASSERT_EQ(fm_workbook_get_cell_style_count(wb.handle, &style_count), 0);

  ASSERT_EQ(fm_workbook_remove_cell_style(wb.handle, "Dated"), 0);

  fm_cell_xf_t reread{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(wb.handle, cell_xf_index, &reread), 0);
  EXPECT_EQ(reread.style_xf_index, 0U);
  EXPECT_EQ(reread.font_index, bold);
  EXPECT_EQ(reread.num_fmt_id, 0U);
  EXPECT_EQ(reread.apply_font, 1);
  EXPECT_EQ(reread.apply_number_format, 0);

  // Neither xf table is compacted; only the cellStyle record is gone.
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_cell_xf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, xf_count);
  ASSERT_EQ(fm_workbook_get_cell_style_xf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, style_xf_count);
  ASSERT_EQ(fm_workbook_get_cell_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, style_count - 1U);
  fm_cell_style_t found{};
  EXPECT_EQ(FindStyle(wb.handle, "Dated", &found), UINT32_MAX);
}

TEST(FormulonCApiNamedStyles, RemoveCellStyleKeepsXfsOfASharedStyleXf) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t bold = 0;
  const uint32_t xf_id = AddBoldDateStyleXf(wb.handle, &bold);
  const fm_cell_style_t first = StyleRecord("First", xf_id, kCustomStyle);
  const fm_cell_style_t second = StyleRecord("Second", xf_id, kCustomStyle);
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &first), 0);
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &second), 0);
  fm_cell_xf_t cell_xf{};
  cell_xf.style_xf_index = xf_id;
  cell_xf.font_index = static_cast<uint32_t>(bold);
  cell_xf.num_fmt_id = 14;
  std::uint32_t cell_xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &cell_xf, &cell_xf_index), 0);

  ASSERT_EQ(fm_workbook_remove_cell_style(wb.handle, "First"), 0);
  fm_cell_xf_t reread{};
  ASSERT_EQ(fm_workbook_get_cell_xf_at(wb.handle, cell_xf_index, &reread), 0);
  EXPECT_EQ(reread.style_xf_index, xf_id);
  EXPECT_EQ(reread.num_fmt_id, 14U);
}

TEST(FormulonCApiNamedStyles, RemoveCellStyleRejectsNormalAndUnknownNames) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_remove_cell_style(wb.handle, "Normal"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_remove_cell_style(wb.handle, "Missing"), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_workbook_remove_cell_style(nullptr, "Normal"), kNullPointer);
  EXPECT_EQ(fm_workbook_remove_cell_style(wb.handle, nullptr), kNullPointer);
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_cell_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiNamedStyles, CellStyleXfIndexPastTheEndNamesTheParameter) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_cell_style_xf_count(wb.handle, &count), 0);
  fm_cell_xf_t xf{};
  EXPECT_EQ(fm_workbook_get_cell_style_xf_at(wb.handle, static_cast<uint32_t>(count), &xf),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
  EXPECT_NE(std::strstr(fm_get_last_error_context(), "cell_style_xf_index="), nullptr) << fm_get_last_error_context();
}

TEST(FormulonCApiNamedStyles, NamesCompareWithUnicodeCaseFolding) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t font = 0;
  const uint32_t xf_id = AddBoldDateStyleXf(wb.handle, &font);
  const fm_cell_style_t record = StyleRecord("\xC3\x84pfel", xf_id, kCustomStyle);  // "Äpfel"
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &record), 0);
  std::size_t count_before = 0;
  ASSERT_EQ(fm_workbook_get_cell_style_count(wb.handle, &count_before), 0);

  int32_t present = 0;
  fm_cell_style_t read{};
  ASSERT_EQ(fm_workbook_get_cell_style(wb.handle, "\xC3\xA4pfel", &present, &read), 0);  // "äpfel"
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(read.name, "\xC3\x84pfel");

  // Setting under the other case replaces the style instead of adding one.
  const fm_cell_style_t again = StyleRecord("\xC3\xA4pfel", xf_id, kCustomStyle);
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &again), 0);
  std::size_t count_after = 0;
  ASSERT_EQ(fm_workbook_get_cell_style_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count_before);

  ASSERT_EQ(fm_workbook_remove_cell_style(wb.handle, "\xC3\x84PFEL"), 0);
  ASSERT_EQ(fm_workbook_get_cell_style(wb.handle, "\xC3\xA4pfel", &present, &read), 0);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiCellStyleOutlineLevel, RoundTripsOutlineLevel) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_style_t record = StyleRecord("ColLevel_5", 0U, 2U);
  record.outline_level = 4;
  ASSERT_EQ(fm_workbook_set_cell_style(wb.handle, &record), 0);
  int32_t present = 0;
  fm_cell_style_t read{};
  ASSERT_EQ(fm_workbook_get_cell_style(wb.handle, "ColLevel_5", &present, &read), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(read.outline_level, 4U);
  EXPECT_EQ(read.builtin_id, 2U);
}
