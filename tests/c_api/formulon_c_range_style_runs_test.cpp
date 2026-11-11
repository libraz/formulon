// Range styles over blank cells held as style runs: every range writer accepts
// a large partial rect without a stored cell per blank address, one area guard
// refuses past 2^27 cells, an explicit default (xf 0) blank survives only where
// it overrides a row or column format, and runs follow copies, structural
// edits, enumeration and the undo capture.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "style_region.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace {

constexpr std::uint32_t kLastCol = formulon::Sheet::kMaxCols - 1U;
constexpr std::uint32_t kLastRow = formulon::Sheet::kMaxRows - 1U;
constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kIndexOutOfRange = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange);
constexpr fm_status_t kSecurityResourceLimit =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kSecurityResourceLimit);
constexpr std::int64_t kNoFormat = -1;
constexpr std::size_t kTwoMiB = 2U * 1024U * 1024U;

void MakeWorkbook(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(10U);
}

formulon::Sheet& SheetOf(WorkbookGuard& wb, std::size_t index = 0U) {
  return wb.handle->workbook().sheet(index);
}

// The coordinate's own xf (stored cell or style run), or `kNoFormat`.
std::int64_t OwnXf(WorkbookGuard& wb, std::uint32_t row, std::uint32_t col, std::size_t sheet = 0U) {
  const std::optional<formulon::StyleSlot> slot = SheetOf(wb, sheet).cell_style_at(row, col);
  return slot.has_value() ? static_cast<std::int64_t>(slot->xf) : kNoFormat;
}

// `fm_sheet_get_cell_xf_index`'s view: (present, xf).
std::pair<std::int32_t, std::uint32_t> CXf(WorkbookGuard& wb, std::uint32_t row, std::uint32_t col) {
  std::int32_t present = -1;
  std::uint32_t xf = 99U;
  EXPECT_EQ(fm_sheet_get_cell_xf_index(wb.handle, 0, row, col, &present, &xf), 0);
  return {present, xf};
}

std::size_t CellCount(WorkbookGuard& wb, std::size_t sheet = 0U) {
  std::size_t count = 0U;
  EXPECT_EQ(fm_sheet_get_cell_count(wb.handle, sheet, &count), 0);
  return count;
}

std::size_t MemoryBytes(WorkbookGuard& wb) {
  std::size_t bytes = 0U;
  EXPECT_EQ(fm_workbook_get_memory_usage(wb.handle, &bytes), 0);
  return bytes;
}

fm_status_t SetRange(WorkbookGuard& wb, fm_cell_rect_t rect, std::uint32_t xf) {
  return fm_sheet_set_range_xf_index(wb.handle, 0, &rect, xf);
}

fm_status_t RemapRange(WorkbookGuard& wb, fm_cell_rect_t rect, const std::vector<fm_xf_remap_t>& pairs) {
  return fm_sheet_remap_range_xfs(wb.handle, 0, &rect, pairs.data(), pairs.size());
}

fm_status_t SetLink(WorkbookGuard& wb, fm_cell_rect_t rect, const char* display = nullptr) {
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  link.rect = rect;
  link.target = "https://example.com";
  link.display = display;
  return fm_sheet_set_hyperlink(wb.handle, 0, &link);
}

std::vector<std::uint8_t> Snapshot(WorkbookGuard& wb, fm_cell_rect_t rect) {
  BufferGuard capture;
  EXPECT_EQ(fm_sheet_snapshot_range_xfs(wb.handle, 0, &rect, &capture.data, &capture.len), 0);
  return std::vector<std::uint8_t>(capture.data, capture.data + capture.len);
}

// Little-endian capture bytes built by hand.
struct Blob {
  std::vector<std::uint8_t> bytes;
  Blob& u8(std::uint8_t v) {
    bytes.push_back(v);
    return *this;
  }
  Blob& u32(std::uint32_t v) {
    for (int shift = 0; shift < 32; shift += 8) {
      bytes.push_back(static_cast<std::uint8_t>((v >> shift) & 0xFFU));
    }
    return *this;
  }
  // Magic, version and the partial rect A1:E5 (no axis captures).
  static Blob Header(std::uint8_t version) {
    Blob blob;
    blob.u8('F').u8('M').u8('R').u8('S').u8(version);
    blob.u32(0U).u32(0U).u32(4U).u32(4U).u8(0U);
    return blob;
  }
  Blob& cell(std::uint32_t col, std::uint32_t xf, std::uint8_t explicit_xf) {
    return u8(0U).u32(col).u32(xf).u8(explicit_xf);
  }
  Blob& run(std::uint32_t first, std::uint32_t last, std::uint32_t xf) { return u8(1U).u32(first).u32(last).u32(xf); }
};

fm_status_t Restore(WorkbookGuard& wb, const Blob& blob) {
  return fm_sheet_restore_range_xfs(wb.handle, 0, blob.bytes.data(), blob.bytes.size());
}

}  // namespace

TEST(FormulonCApiRangeStyleRuns, EveryRangeWriterAcceptsA1ToZ100000) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  const fm_cell_rect_t rect{0U, 0U, 99999U, 25U};

  ASSERT_EQ(SetLink(wb, rect, "Shown"), 0);
  EXPECT_EQ(CellCount(wb), 2600000U);
  EXPECT_NE(OwnXf(wb, 99999U, 25U), kNoFormat);
  EXPECT_EQ(SheetOf(wb).cell_at(99999U, 25U), nullptr);

  ASSERT_EQ(fm_sheet_remove_hyperlink(wb.handle, 0, &rect), 0);
  EXPECT_EQ(CellCount(wb), 1U);  // A1's text
  EXPECT_EQ(OwnXf(wb, 50000U, 13U), kNoFormat);
  EXPECT_EQ(OwnXf(wb, 0U, 0U), 0);

  ASSERT_EQ(SetLink(wb, rect), 0);
  ASSERT_EQ(fm_sheet_clear_hyperlinks(wb.handle, 0), 0);
  EXPECT_EQ(CellCount(wb), 1U);

  ASSERT_EQ(SetRange(wb, rect, 3U), 0);
  EXPECT_EQ(CellCount(wb), 2600000U);
  EXPECT_EQ(OwnXf(wb, 0U, 0U), 3);
  EXPECT_EQ(OwnXf(wb, 50000U, 13U), 3);

  ASSERT_EQ(RemapRange(wb, rect, {{3U, 4U}}), 0);
  EXPECT_EQ(OwnXf(wb, 0U, 0U), 4);
  EXPECT_EQ(OwnXf(wb, 99999U, 25U), 4);
  EXPECT_EQ(CellCount(wb), 2600000U);
}

TEST(FormulonCApiRangeStyleRuns, HyperlinkOverTenThousandRowsGrowsMemoryByLessThanTwoMiB) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  const std::size_t before = MemoryBytes(wb);
  ASSERT_EQ(SetLink(wb, {0U, 0U, 9999U, 25U}, "Shown"), 0);
  EXPECT_EQ(CellCount(wb), 260000U);
  EXPECT_LT(MemoryBytes(wb) - before, kTwoMiB);
}

TEST(FormulonCApiRangeStyleRuns, AreaGuardBoundaries) {
  {
    // A1:LCB16384 is exactly 2^27 cells.
    WorkbookGuard wb;
    MakeWorkbook(wb);
    ASSERT_EQ(SetRange(wb, {0U, 0U, 16383U, 8191U}, 2U), 0);
    EXPECT_EQ(CellCount(wb), std::size_t{1} << 27U);
    EXPECT_EQ(OwnXf(wb, 0U, 0U), 2);
    EXPECT_EQ(OwnXf(wb, 0U, 8191U), 2);
    EXPECT_EQ(OwnXf(wb, 16383U, 0U), 2);
    EXPECT_EQ(OwnXf(wb, 16383U, 8191U), 2);
    EXPECT_EQ(OwnXf(wb, 0U, 8192U), kNoFormat);
  }
  {
    WorkbookGuard wb;
    MakeWorkbook(wb);
    ASSERT_EQ(RemapRange(wb, {0U, 1U, 8191U, kLastCol}, {{0U, 2U}}), 0);
    EXPECT_EQ(CellCount(wb), std::size_t{8192U} * 16383U);
    EXPECT_EQ(OwnXf(wb, 0U, 1U), 2);
    EXPECT_EQ(OwnXf(wb, 8191U, kLastCol), 2);
    EXPECT_EQ(OwnXf(wb, 0U, 0U), kNoFormat);
  }
  {
    // B1:XFD8193 is one row past; nothing changes, and a bad xf is reported first.
    WorkbookGuard wb;
    MakeWorkbook(wb);
    const fm_cell_rect_t over{0U, 1U, 8192U, kLastCol};
    EXPECT_EQ(SetRange(wb, over, 2U), kSecurityResourceLimit);
    EXPECT_EQ(RemapRange(wb, over, {{0U, 2U}}), kSecurityResourceLimit);
    EXPECT_EQ(SetLink(wb, over, "Shown"), kSecurityResourceLimit);
    EXPECT_EQ(SetRange(wb, over, 10U), kIndexOutOfRange);
    EXPECT_EQ(CellCount(wb), 0U);
    EXPECT_TRUE(SheetOf(wb).layout().columns.empty());
    // Whole columns are axis formats and never refused.
    EXPECT_EQ(SetRange(wb, {0U, 1U, kLastRow, kLastCol}, 2U), 0);
    EXPECT_EQ(CellCount(wb), 0U);
  }
}

struct MapCase {
  const char* name;
  std::vector<fm_xf_remap_t> pairs;
  // Expected own xf of A1:D2, row-major.
  std::int64_t expected[8];
};

class FormulonCApiRangeStyleRunsMap : public ::testing::TestWithParam<MapCase> {};

// Row 1: A1 a number with xf 1 (stored owner), B1 run 2, C1 gap, D1 run 1.
// Row 2: A2 text with xf 2 (stored owner), B2 gap, C2 run 1, D2 run 2.
TEST_P(FormulonCApiRangeStyleRunsMap, MapsRunOwnersStoredOwnersAndGapsOnce) {
  const MapCase& c = GetParam();
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0U, 0U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0U, 1U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0U, 3U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 1U, 0U, "x"), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 1U, 0U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 1U, 2U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 1U, 3U, 2U), 0);

  ASSERT_EQ(RemapRange(wb, {0U, 0U, 1U, 3U}, c.pairs), 0);
  for (std::uint32_t i = 0; i < 8U; ++i) {
    EXPECT_EQ(OwnXf(wb, i / 4U, i % 4U), c.expected[i]) << c.name << " " << i;
  }
  EXPECT_NE(SheetOf(wb).cell_at(0U, 0U), nullptr);
  EXPECT_EQ(SheetOf(wb).cell_at(0U, 1U), nullptr);
}

INSTANTIATE_TEST_SUITE_P(Maps, FormulonCApiRangeStyleRunsMap,
                         ::testing::Values(
                             // A gap maps xf 0 to 0, which nothing overrides, so it stays empty.
                             MapCase{"swap", {{1U, 2U}, {2U, 1U}}, {2, 1, kNoFormat, 2, 1, kNoFormat, 2, 1}},
                             // Each coordinate is mapped once: 0 -> 1 never continues to 2.
                             MapCase{"chain", {{0U, 1U}, {1U, 2U}}, {2, 2, 1, 2, 2, 1, 2, 2}}),
                         [](const ::testing::TestParamInfo<MapCase>& info) { return std::string(info.param.name); });

TEST(FormulonCApiRangeStyleRuns, AxisRemapDecidesExplicitDefaultsAgainstTheRewrittenAxis) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  // B5 is an explicit default under column B's xf 3; D6 holds xf 2 under an unstyled column.
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 1U, 1U, 3U), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 4U, 1U, 0U), 0);
  ASSERT_EQ(CXf(wb, 4U, 1U), (std::pair<std::int32_t, std::uint32_t>{1, 0U}));
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 5U, 3U, 2U), 0);

  // Column B's 3 becomes 0: B5's override no longer overrides anything.
  ASSERT_EQ(RemapRange(wb, {0U, 1U, kLastRow, 1U}, {{3U, 0U}}), 0);
  EXPECT_EQ(CXf(wb, 4U, 1U), (std::pair<std::int32_t, std::uint32_t>{0, 0U}));

  // Column D's 0 becomes 6 while D6's 2 becomes 0: that 0 now overrides the column.
  ASSERT_EQ(RemapRange(wb, {0U, 3U, kLastRow, 3U}, {{0U, 6U}, {2U, 0U}}), 0);
  EXPECT_EQ(CXf(wb, 5U, 3U), (std::pair<std::int32_t, std::uint32_t>{1, 0U}));
  EXPECT_EQ(CXf(wb, 6U, 3U), (std::pair<std::int32_t, std::uint32_t>{0, 0U}));
  EXPECT_EQ(CellCount(wb), 1U);
}

TEST(FormulonCApiRangeStyleRuns, SheetCopyCarriesRuns) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(SetRange(wb, {1U, 1U, 3U, 4U}, 5U), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2U, 2U, 9.0), 0);
  std::size_t copy = 0U;
  ASSERT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 1, "Copy", &copy), 0);
  EXPECT_EQ(CellCount(wb, copy), CellCount(wb, 0));
  EXPECT_EQ(OwnXf(wb, 1U, 1U, copy), 5);
  EXPECT_EQ(OwnXf(wb, 3U, 4U, copy), 5);
  EXPECT_EQ(OwnXf(wb, 2U, 2U, copy), 5);
  EXPECT_EQ(SheetOf(wb, copy).cell_at(3U, 4U), nullptr);
}

TEST(FormulonCApiRangeStyleRuns, DataTableInputOnARunReturnsToItsRun) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  // B1 (the column input) is a blank with xf 3; B2 = B1 * B1 over A3:A5 = {1, 2, 3}.
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0U, 1U, 3U), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1U, 1U, "=B1*B1"), 0);
  for (std::uint32_t i = 0; i < 3U; ++i) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2U + i, 0U, static_cast<double>(i + 1U)), 0);
  }
  fm_data_table_t spec = FM_STRUCT_INIT(fm_data_table_t);
  spec.rect = fm_cell_rect_t{2U, 1U, 4U, 1U};
  spec.has_column_input = 1;
  spec.column_input_row = 0U;
  spec.column_input_col = 1U;
  ASSERT_EQ(fm_sheet_add_data_table(wb.handle, 0, &spec, nullptr), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  for (std::uint32_t i = 0; i < 3U; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 2U + i, 1U, &v), 0);
    ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
    EXPECT_DOUBLE_EQ(v.u.number, static_cast<double>((i + 1U) * (i + 1U)));
  }
  EXPECT_EQ(SheetOf(wb).cell_at(0U, 1U), nullptr);
  EXPECT_EQ(OwnXf(wb, 0U, 1U), 3);
}

namespace {

// "Sales" A1:C4 on sheet 0: header, two data rows, totals row.
void AddSalesTable(formulon::Workbook& book) {
  formulon::TableMetadata table;
  table.sheet_index = 0;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:C4";
  table.totals_row = true;
  for (const char* name : {"Region", "Amount", "Qty"}) {
    formulon::TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  for (std::uint32_t col = 0; col < 3U; ++col) {
    ASSERT_TRUE(book.set_cell_text(0, 0U, col, table.columns[col].name).has_value());
    ASSERT_TRUE(book.set_cell_value(0, 1U, col, formulon::Value::number(1.0)).has_value());
    ASSERT_TRUE(book.set_cell_value(0, 2U, col, formulon::Value::number(2.0)).has_value());
  }
  ASSERT_TRUE(book.add_table(std::move(table)).has_value());
}

}  // namespace

TEST(FormulonCApiRangeStyleRuns, TableHeaderRemovalDropsAFormattedBlankHeaderCell) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  formulon::Workbook& book = wb.handle->workbook();
  AddSalesTable(book);
  // B1 left as a blank carrying a format, as a loaded file can hold it.
  SheetOf(wb).set_cell_value(0U, 1U, formulon::Value::blank());
  SheetOf(wb).set_cell_xf_index(0U, 1U, 4U);
  ASSERT_EQ(OwnXf(wb, 0U, 1U), 4);
  ASSERT_TRUE(book.set_table_header_row(0, false).has_value());
  EXPECT_EQ(OwnXf(wb, 0U, 0U), kNoFormat);
  EXPECT_EQ(OwnXf(wb, 0U, 1U), kNoFormat);
  EXPECT_EQ(OwnXf(wb, 0U, 2U), kNoFormat);
}

TEST(FormulonCApiRangeStyleRuns, TotalsSyncTreatsAFormattedBlankTotalsCellAsPresent) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  formulon::Workbook& book = wb.handle->workbook();
  AddSalesTable(book);
  book.mutable_tables()[0].columns[1].totals_function = "sum";
  SheetOf(wb).set_cell_xf_index(3U, 1U, 4U);
  book.normalize_table_totals_cells();
  EXPECT_TRUE(book.tables()[0].columns[1].totals_function.empty());
  EXPECT_EQ(OwnXf(wb, 3U, 1U), 4);
  EXPECT_EQ(SheetOf(wb).cell_at(3U, 1U), nullptr);
}

TEST(FormulonCApiRangeStyleRuns, InsertAndDeleteCellsShiftRunsWithTheBand) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(SetRange(wb, {1U, 1U, 1U, 2U}, 3U), 0);  // B2:C2
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2U, 1U, "t"), 0);
  const std::size_t count = CellCount(wb);

  const fm_cell_rect_t band{0U, 1U, 0U, 2U};  // B1:C1, no neighbour above
  ASSERT_EQ(fm_sheet_insert_cells(wb.handle, 0, &band, FM_CELL_SHIFT_DOWN), 0);
  EXPECT_EQ(OwnXf(wb, 1U, 1U), kNoFormat);
  EXPECT_EQ(OwnXf(wb, 2U, 1U), 3);
  EXPECT_EQ(OwnXf(wb, 2U, 2U), 3);
  EXPECT_NE(SheetOf(wb).cell_at(3U, 1U), nullptr);
  EXPECT_EQ(CellCount(wb), count);

  ASSERT_EQ(fm_sheet_delete_cells(wb.handle, 0, &band, FM_CELL_SHIFT_UP), 0);
  EXPECT_EQ(OwnXf(wb, 1U, 1U), 3);
  EXPECT_EQ(OwnXf(wb, 1U, 2U), 3);
  EXPECT_NE(SheetOf(wb).cell_at(2U, 1U), nullptr);
  EXPECT_EQ(CellCount(wb), count);

  // Inserted below B2:C2, the new cells take its format as runs.
  const fm_cell_rect_t below{2U, 1U, 2U, 2U};
  ASSERT_EQ(fm_sheet_insert_cells(wb.handle, 0, &below, FM_CELL_SHIFT_DOWN), 0);
  EXPECT_EQ(OwnXf(wb, 2U, 1U), 3);
  EXPECT_EQ(OwnXf(wb, 2U, 2U), 3);
  EXPECT_EQ(SheetOf(wb).cell_at(2U, 1U), nullptr);
  EXPECT_NE(SheetOf(wb).cell_at(3U, 1U), nullptr);
  EXPECT_EQ(CellCount(wb), count + 2U);
}

TEST(FormulonCApiRangeStyleRuns, InsertingTenThousandRowsUnderAFormattedBandStaysUnderTwoMiB) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(SetRange(wb, {0U, 0U, 0U, 25U}, 3U), 0);
  const std::size_t before = MemoryBytes(wb);
  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0, 1U, 10000U), 0);
  EXPECT_EQ(CellCount(wb), 26U * 10001U);
  EXPECT_EQ(OwnXf(wb, 10000U, 25U), 3);
  EXPECT_EQ(SheetOf(wb).cell_at(10000U, 25U), nullptr);
  EXPECT_LT(MemoryBytes(wb) - before, kTwoMiB);
}

struct EdgeCase {
  const char* name;
  void (*setup)(fm_workbook_t* wb);
  void (*act)(fm_workbook_t* wb);
  std::int64_t own_xf;     // at A1 after `act`
  bool stored;             // A1 is a stored cell after `act`
  std::size_t cell_count;  // after `act`
};

class FormulonCApiRangeStyleRunsEdge : public ::testing::TestWithParam<EdgeCase> {};

TEST_P(FormulonCApiRangeStyleRunsEdge, FollowsTheBlankSlotRule) {
  const EdgeCase& c = GetParam();
  WorkbookGuard wb;
  MakeWorkbook(wb);
  c.setup(wb.handle);
  c.act(wb.handle);
  EXPECT_EQ(OwnXf(wb, 0U, 0U), c.own_xf) << c.name;
  EXPECT_EQ(SheetOf(wb).cell_at(0U, 0U) != nullptr, c.stored) << c.name;
  EXPECT_EQ(CellCount(wb), c.cell_count) << c.name;
}

namespace {

const fm_cell_rect_t kA1{0U, 0U, 0U, 0U};

void Nothing(fm_workbook_t*) {}
void Number(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 0U, 0U, 1.0), 0);
}
void NumberXf3(fm_workbook_t* wb) {
  Number(wb);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb, 0, 0U, 0U, 3U), 0);
}
void NumberXf0(fm_workbook_t* wb) {
  Number(wb);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb, 0, 0U, 0U, 0U), 0);
}
void RunXf3(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb, 0, 0U, 0U, 3U), 0);
}
void ColumnXf4(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb, 0, 0U, 0U, 4U), 0);
}
void RowXf4(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb, 0, 0U, 0U, 4U), 0);
}
void Blank(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_blank(wb, 0, 0U, 0U), 0);
}
void RangeXf3(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_range_xf_index(wb, 0, &kA1, 3U), 0);
}
void RangeXf0(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_range_xf_index(wb, 0, &kA1, 0U), 0);
}
void CellXf0(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb, 0, 0U, 0U, 0U), 0);
}
void ResetXf(fm_workbook_t* wb) {
  ASSERT_EQ(fm_sheet_reset_cell_xf_index(wb, 0, 0U, 0U), 0);
}

}  // namespace

INSTANTIATE_TEST_SUITE_P(
    Edges, FormulonCApiRangeStyleRunsEdge,
    ::testing::Values(EdgeCase{"content_styled_in_place", Number, RangeXf3, 3, true, 1U},
                      EdgeCase{"blank_becomes_run", Nothing, RangeXf3, 3, false, 1U},
                      EdgeCase{"range_xf0_without_axis_leaves_nothing", Nothing, RangeXf0, kNoFormat, false, 0U},
                      EdgeCase{"range_xf0_under_column_is_kept", ColumnXf4, RangeXf0, 0, false, 1U},
                      EdgeCase{"cell_xf0_without_axis_leaves_nothing", Nothing, CellXf0, kNoFormat, false, 0U},
                      EdgeCase{"cell_xf0_under_row_is_kept", RowXf4, CellXf0, 0, false, 1U},
                      EdgeCase{"blank_on_run_changes_nothing", RunXf3, Blank, 3, false, 1U},
                      EdgeCase{"blank_on_formatted_content_leaves_run", NumberXf3, Blank, 3, false, 1U},
                      EdgeCase{"blank_on_explicit_default_content_removes", NumberXf0, Blank, kNoFormat, false, 0U},
                      EdgeCase{"blank_on_unformatted_content_removes", Number, Blank, kNoFormat, false, 0U},
                      EdgeCase{"reset_on_run_removes", RunXf3, ResetXf, kNoFormat, false, 0U},
                      EdgeCase{"reset_on_content_keeps_cell", NumberXf3, ResetXf, 0, true, 1U}),
    [](const ::testing::TestParamInfo<EdgeCase>& info) { return std::string(info.param.name); });

TEST(FormulonCApiRangeStyleRuns, EnumerationSeesRunsAndFollowsARangeWrite) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0U, 4U, "=SEQUENCE(3)"), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 1U, 4U, 3U), 0);  // a run under the spill's E2
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  const auto addresses = [&wb] {
    std::vector<std::pair<std::uint32_t, std::uint32_t>> out;
    const std::size_t count = CellCount(wb);
    for (std::size_t i = 0; i < count; ++i) {
      std::uint32_t row = 0U;
      std::uint32_t col = 0U;
      const char* formula = nullptr;
      fm_value_t value{};
      EXPECT_EQ(fm_sheet_get_cell_at(wb.handle, 0, i, &row, &col, &formula, &value), 0);
      out.emplace_back(row, col);
    }
    std::uint32_t row = 0U;
    std::uint32_t col = 0U;
    fm_value_t value{};
    EXPECT_NE(fm_sheet_get_cell_at(wb.handle, 0, count, &row, &col, nullptr, &value), 0);
    return out;
  };
  using Cells = std::vector<std::pair<std::uint32_t, std::uint32_t>>;
  EXPECT_EQ(addresses(), (Cells{{0U, 0U}, {0U, 4U}, {1U, 4U}, {2U, 4U}}));

  ASSERT_EQ(SetRange(wb, {1U, 1U, 1U, 2U}, 3U), 0);
  EXPECT_EQ(addresses(), (Cells{{0U, 0U}, {0U, 4U}, {1U, 1U}, {1U, 2U}, {1U, 4U}, {2U, 4U}}));
  std::uint32_t row = 0U;
  std::uint32_t col = 0U;
  const char* formula = "unset";
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 2U, &row, &col, &formula, &value), 0);
  EXPECT_EQ(formula, nullptr);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_BLANK);
}

TEST(FormulonCApiRangeStyleRuns, CellXfIndicesCountAndCopyWithASmallCapacity) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0U, 0U, "a"), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0U, 0U, 2U), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1U, 0U, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 1U, 0U, 0U), 0);  // own xf 0: not listed
  ASSERT_EQ(SetRange(wb, {1U, 1U, 1U, 3U}, 3U), 0);
  const fm_cell_rect_t all{0U, 0U, 9U, 9U};

  std::size_t total = 0U;
  ASSERT_EQ(fm_sheet_get_cell_xf_indices(wb.handle, 0, &all, nullptr, 0U, &total), 0);
  EXPECT_EQ(total, 4U);
  fm_xf_cell_t out[2] = {};
  ASSERT_EQ(fm_sheet_get_cell_xf_indices(wb.handle, 0, &all, out, 2U, &total), 0);
  EXPECT_EQ(total, 4U);
  EXPECT_EQ(out[0].row, 0U);
  EXPECT_EQ(out[0].col, 0U);
  EXPECT_EQ(out[0].xf_index, 2U);
  EXPECT_EQ(out[1].row, 1U);
  EXPECT_EQ(out[1].col, 1U);
  EXPECT_EQ(out[1].xf_index, 3U);
}

TEST(FormulonCApiRangeStyleRuns, SnapshotBytesAreCanonical) {
  const fm_cell_rect_t rect{0U, 0U, 4U, 4U};
  WorkbookGuard singles;
  MakeWorkbook(singles);
  for (std::uint32_t col = 1U; col <= 3U; ++col) {
    ASSERT_EQ(fm_sheet_set_cell_xf_index(singles.handle, 0, 1U, col, 3U), 0);
  }
  WorkbookGuard span;
  MakeWorkbook(span);
  ASSERT_EQ(SetRange(span, {1U, 1U, 1U, 3U}, 3U), 0);
  const std::vector<std::uint8_t> bytes = Snapshot(singles, rect);
  EXPECT_EQ(bytes, Snapshot(span, rect));
  ASSERT_GT(bytes.size(), 4U);
  EXPECT_EQ(bytes[4], 2U);
  // One row group holding one run B2:D2.
  const Blob expected = Blob::Header(2U).u32(1U).u32(1U).u32(1U).run(1U, 3U, 3U);
  EXPECT_EQ(bytes, expected.bytes);
}

TEST(FormulonCApiRangeStyleRuns, SnapshotRestoresRunsAndCells) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2U, 2U, 7.0), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 2U, 2U, 5U), 0);
  ASSERT_EQ(SetRange(wb, {1U, 0U, 1U, 4U}, 3U), 0);
  const fm_cell_rect_t rect{0U, 0U, 4U, 4U};
  const std::vector<std::uint8_t> capture = Snapshot(wb, rect);

  ASSERT_EQ(SetRange(wb, rect, 6U), 0);
  ASSERT_EQ(fm_sheet_restore_range_xfs(wb.handle, 0, capture.data(), capture.size()), 0);
  EXPECT_EQ(Snapshot(wb, rect), capture);
  EXPECT_EQ(OwnXf(wb, 0U, 0U), kNoFormat);
  EXPECT_EQ(OwnXf(wb, 1U, 4U), 3);
  EXPECT_EQ(OwnXf(wb, 2U, 2U), 5);
  EXPECT_EQ(CellCount(wb), 6U);
}

TEST(FormulonCApiRangeStyleRuns, RestoreAcceptsTheEarlierPerCellVersion) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  // B2 xf 3, C3 an explicit default with no axis format, D4 one under row 4's xf 4.
  ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 3U, 3U, 4U), 0);
  const Blob v1 =
      Blob::Header(1U).u32(3U).u32(1U).u32(1U).u32(3U).u8(1U).u32(2U).u32(2U).u32(0U).u8(1U).u32(3U).u32(3U).u32(0U).u8(
          1U);
  ASSERT_EQ(Restore(wb, v1), 0);
  EXPECT_EQ(OwnXf(wb, 1U, 1U), 3);
  EXPECT_EQ(OwnXf(wb, 2U, 2U), kNoFormat);
  EXPECT_EQ(OwnXf(wb, 3U, 3U), 0);
  EXPECT_EQ(SheetOf(wb).cell_at(1U, 1U), nullptr);
}

TEST(FormulonCApiRangeStyleRuns, RestoreMergesAdjacentEqualRuns) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  const Blob split = Blob::Header(2U).u32(1U).u32(0U).u32(2U).run(0U, 1U, 3U).run(2U, 4U, 3U);
  ASSERT_EQ(Restore(wb, split), 0);
  const Blob merged = Blob::Header(2U).u32(1U).u32(0U).u32(1U).run(0U, 4U, 3U);
  EXPECT_EQ(Snapshot(wb, {0U, 0U, 4U, 4U}), merged.bytes);
}

struct RefusalCase {
  const char* name;
  Blob blob;
};

class FormulonCApiRangeStyleRunsRefusal : public ::testing::TestWithParam<RefusalCase> {};

TEST_P(FormulonCApiRangeStyleRunsRefusal, ChangesNothing) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(SetRange(wb, {1U, 1U, 1U, 2U}, 3U), 0);
  const fm_cell_rect_t rect{0U, 0U, 4U, 4U};
  const std::vector<std::uint8_t> before = Snapshot(wb, rect);
  EXPECT_EQ(Restore(wb, GetParam().blob), kInvalidArgument) << GetParam().name;
  EXPECT_EQ(Snapshot(wb, rect), before) << GetParam().name;
}

INSTANTIATE_TEST_SUITE_P(
    Refusals, FormulonCApiRangeStyleRunsRefusal,
    ::testing::Values(
        RefusalCase{"row_outside_rect", Blob::Header(2U).u32(1U).u32(5U).u32(1U).run(0U, 0U, 3U)},
        RefusalCase{"rows_not_ascending",
                    Blob::Header(2U).u32(2U).u32(2U).u32(1U).run(0U, 0U, 3U).u32(1U).u32(1U).run(0U, 0U, 3U)},
        RefusalCase{"rows_repeated",
                    Blob::Header(2U).u32(2U).u32(1U).u32(1U).run(0U, 0U, 3U).u32(1U).u32(1U).run(2U, 2U, 3U)},
        RefusalCase{"entries_overlapping", Blob::Header(2U).u32(1U).u32(1U).u32(2U).run(0U, 2U, 3U).run(2U, 3U, 3U)},
        RefusalCase{"entries_not_ascending", Blob::Header(2U).u32(1U).u32(1U).u32(2U).cell(3U, 3U, 1U).run(0U, 1U, 3U)},
        RefusalCase{"column_outside_rect", Blob::Header(2U).u32(1U).u32(1U).u32(1U).run(3U, 5U, 3U)},
        RefusalCase{"first_after_last", Blob::Header(2U).u32(1U).u32(1U).u32(1U).run(3U, 2U, 3U)},
        RefusalCase{"xf_past_table", Blob::Header(2U).u32(1U).u32(1U).u32(1U).run(0U, 0U, 10U)},
        RefusalCase{"unknown_tag", Blob::Header(2U).u32(1U).u32(1U).u32(1U).u8(2U).u32(0U).u32(0U).u32(3U)},
        RefusalCase{"explicit_flag_out_of_range", Blob::Header(2U).u32(1U).u32(1U).u32(1U).cell(0U, 3U, 2U)},
        RefusalCase{"unknown_version", Blob::Header(3U).u32(0U)},
        RefusalCase{"trailing_bytes", Blob::Header(2U).u32(0U).u8(0U)}),
    [](const ::testing::TestParamInfo<RefusalCase>& info) { return std::string(info.param.name); });
