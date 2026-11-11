// The XLSB sheet reader's load budget: new cells, new row nodes, the style
// runs blank cells add and every buffered row header are charged, with the
// .xlsx reader's error shape; a blank `BrtCellBlank` loads as a style run.

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/sheet_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/sheet_writer.h"
#include "io/xlsb/sst_writer.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint64_t kRowHeaderBytes = sizeof(SheetDecodeState::DecodedRowHeader);

Workbook StyledWorkbook() {
  Workbook wb = Workbook::create();
  wb.mutable_styles().cell_xfs.resize(4U);
  return wb;
}

std::vector<std::uint8_t> EmitBody(const Sheet& sheet) {
  SstBuilder sst;
  auto body = emit_sheet(sheet, sst, {sheet.name()}, SheetRangeTable{}, NameTable{}, hyperlink_relationship_ids(sheet));
  EXPECT_TRUE(static_cast<bool>(body)) << body.error().message;
  return body.value();
}

Expected<SheetDecodeState, Error> Decode(const std::vector<std::uint8_t>& body, Workbook& wb,
                                         std::uint64_t max_cell_bytes) {
  std::deque<std::string> text_storage;
  std::uint32_t undecoded = 0;
  return DecodeSheetBin(body, 0, wb, {}, {}, {}, text_storage, {wb.sheet(0).name()}, {}, {}, {}, &undecoded,
                        /*rich_errors=*/nullptr, /*deferred=*/nullptr, max_cell_bytes);
}

// One row of `count` numbers costs a row header, a row node and `count` cells.
TEST(XlsbSheetBudget, ValueCellsChargeTheRowNodeAndEachCell) {
  Workbook source = StyledWorkbook();
  for (std::uint32_t col = 0; col < 11U; ++col) {
    source.sheet(0).set_cell_value(0U, col, Value::number(1.0));
  }
  const std::vector<std::uint8_t> body = EmitBody(source.sheet(0));
  const std::uint64_t exact = kRowHeaderBytes + kSheetLoadRowBytes + 11U * kSheetLoadCellBytes;

  Workbook fits = StyledWorkbook();
  auto ok = Decode(body, fits, exact);
  ASSERT_TRUE(static_cast<bool>(ok)) << ok.error().message;
  EXPECT_EQ(fits.sheet(0).cell_count(), 11U);

  Workbook over = StyledWorkbook();
  auto refused = Decode(body, over, exact - 1U);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kSecurityResourceLimit);
  EXPECT_EQ(refused.error().context.rfind("context=xlsb_reader sheet_index=0 row=0 col=10 ", 0), 0U)
      << refused.error().context;
}

// Blank formatted cells are style runs: three non-adjacent ones cost a row
// node and three runs, and load as runs rather than stored cells.
TEST(XlsbSheetBudget, BlankCellsLoadAsRunsAndChargeEachRun) {
  Workbook source = StyledWorkbook();
  for (const std::uint32_t col : {0U, 2U, 4U}) {
    source.sheet(0).set_cell_xf_index(0U, col, 2U);
  }
  const std::vector<std::uint8_t> body = EmitBody(source.sheet(0));
  const std::uint64_t exact = kRowHeaderBytes + kSheetLoadRowBytes + 3U * kSheetLoadRunBytes;

  Workbook fits = StyledWorkbook();
  auto ok = Decode(body, fits, exact);
  ASSERT_TRUE(static_cast<bool>(ok)) << ok.error().message;
  const Sheet& sheet = fits.sheet(0);
  EXPECT_EQ(sheet.cell_count(), 3U);
  EXPECT_EQ(sheet.cell_at(0U, 2U), nullptr);
  ASSERT_TRUE(sheet.cell_style_at(0U, 2U).has_value());
  EXPECT_EQ(sheet.cell_style_at(0U, 2U)->xf, 2U);
  EXPECT_FALSE(sheet.cell_style_at(0U, 1U).has_value());

  Workbook over = StyledWorkbook();
  auto refused = Decode(body, over, exact - 1U);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kSecurityResourceLimit);
  EXPECT_EQ(refused.error().context.rfind("context=xlsb_reader sheet_index=0 row=0 col=4 ", 0), 0U)
      << refused.error().context;
}

// Row headers are buffered until the stream ends, so each one is charged even
// when its row holds no cell.
TEST(XlsbSheetBudget, EveryBufferedRowHeaderIsCharged) {
  Workbook source = StyledWorkbook();
  for (std::uint32_t row = 0; row < 20U; ++row) {
    RowLayout entry;
    entry.row = row;
    entry.height = 30.0;
    entry.has_height = true;
    entry.custom_height = true;
    source.sheet(0).mutable_layout().row_overrides.push_back(entry);
  }
  const std::vector<std::uint8_t> body = EmitBody(source.sheet(0));

  Workbook fits = StyledWorkbook();
  ASSERT_TRUE(static_cast<bool>(Decode(body, fits, 20U * kRowHeaderBytes)));
  Workbook over = StyledWorkbook();
  auto refused = Decode(body, over, 20U * kRowHeaderBytes - 1U);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kSecurityResourceLimit);
}

// An xf past the cellXfs table names no style and loads as the default, as
// the .xlsx reader does for `s=`; with no axis format that blank then goes.
TEST(XlsbSheetBudget, BlankCellXfPastTheTableLoadsAsTheDefault) {
  Workbook source = StyledWorkbook();
  source.mutable_styles().cell_xfs.resize(40U);
  source.sheet(0).set_cell_xf_index(0U, 0U, 30U);
  source.sheet(0).set_cell_value(0U, 1U, Value::number(1.0));
  source.sheet(0).set_cell_xf_index(0U, 1U, 31U);
  const std::vector<std::uint8_t> body = EmitBody(source.sheet(0));

  Workbook loaded = StyledWorkbook();
  auto ok = Decode(body, loaded, kMaxSheetLoadCellBytes);
  ASSERT_TRUE(static_cast<bool>(ok)) << ok.error().message;
  EXPECT_FALSE(loaded.sheet(0).cell_style_at(0U, 0U).has_value());
  ASSERT_NE(loaded.sheet(0).cell_at(0U, 1U), nullptr);
  EXPECT_EQ(loaded.sheet(0).cell_at(0U, 1U)->xf_index, 0U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
