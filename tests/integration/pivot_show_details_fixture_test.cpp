//
// Workbooks Mac Excel 365 (ja-JP) saved after running Show Details on pivot
// value cells: each holds the pivot and the detail sheets Excel made.
//
// The Excel-made detail sheets are captured and then removed, which returns
// the workbook to its state before Show Details (their tables go with them).
// Running Show Details on the same cells, in the same order, must then
// recreate each sheet under Excel's name and position, with its cells, its
// table, and Excel's active sheet.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "sheet.h"
#include "table.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFixture(const char* file) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_options/" + file;
  std::vector<std::uint8_t> out;
  FILE* handle = std::fopen(path.c_str(), "rb");
  if (handle == nullptr) {
    return out;
  }
  std::fseek(handle, 0, SEEK_END);
  const long size = std::ftell(handle);
  std::fseek(handle, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), handle) != out.size()) {
      out.clear();
    }
  }
  std::fclose(handle);
  return out;
}

/// Text of each cell of the sheet's used block, with its number format id.
struct CellText {
  std::string text;
  std::uint32_t num_fmt_id = 0;
  bool operator==(const CellText& other) const { return text == other.text && num_fmt_id == other.num_fmt_id; }
};

/// One detail sheet as observed: where it sits, what it holds, its table.
struct DetailSheet {
  std::string name;
  std::size_t index = 0;
  std::vector<std::vector<CellText>> cells;
  std::optional<TableMetadata> table;
};

std::vector<std::vector<CellText>> SheetCells(const Workbook& wb, std::size_t sheet) {
  std::vector<std::vector<CellText>> out;
  const Sheet& s = wb.sheet(sheet);
  for (std::uint32_t r = 0; s.cell_at(r, 0U) != nullptr; ++r) {
    std::vector<CellText> row;
    for (std::uint32_t c = 0; s.cell_at(r, c) != nullptr; ++c) {
      const Cell& cell = *s.cell_at(r, c);
      CellText text;
      if (cell.cached_value.is_text()) {
        text.text = std::string(cell.cached_value.as_text());
      } else if (cell.cached_value.is_number()) {
        text.text = std::to_string(cell.cached_value.as_number());
      }
      if (cell.xf_index < wb.styles().cell_xfs.size()) {
        text.num_fmt_id = wb.styles().cell_xfs[cell.xf_index].num_fmt_id;
      }
      row.push_back(std::move(text));
    }
    out.push_back(std::move(row));
  }
  return out;
}

DetailSheet Observe(const Workbook& wb, std::size_t sheet) {
  DetailSheet out;
  out.name = wb.sheet(sheet).name();
  out.index = sheet;
  out.cells = SheetCells(wb, sheet);
  for (const TableMetadata& table : wb.tables()) {
    if (table.sheet_index == sheet) {
      out.table = table;
    }
  }
  return out;
}

void ExpectSameDetail(const DetailSheet& made, const DetailSheet& excel) {
  EXPECT_EQ(made.name, excel.name);
  EXPECT_EQ(made.index, excel.index);
  EXPECT_EQ(made.cells, excel.cells) << excel.name;
  ASSERT_TRUE(made.table.has_value());
  ASSERT_TRUE(excel.table.has_value());
  EXPECT_EQ(made.table->name, excel.table->name);
  EXPECT_EQ(made.table->display_name, excel.table->display_name);
  EXPECT_EQ(made.table->ref, excel.table->ref);
  EXPECT_EQ(made.table->header_row, excel.table->header_row);
  EXPECT_EQ(made.table->totals_row, excel.table->totals_row);
  EXPECT_EQ(made.table->totals_row_shown, excel.table->totals_row_shown);
  ASSERT_EQ(made.table->columns.size(), excel.table->columns.size());
  for (std::size_t c = 0; c < made.table->columns.size(); ++c) {
    EXPECT_EQ(made.table->columns[c].name, excel.table->columns[c].name);
  }
  EXPECT_EQ(made.table->style_info.present, excel.table->style_info.present);
  EXPECT_EQ(made.table->style_info.name, excel.table->style_info.name);
  EXPECT_EQ(made.table->style_info.show_first_column, excel.table->style_info.show_first_column);
  EXPECT_EQ(made.table->style_info.show_last_column, excel.table->style_info.show_last_column);
  EXPECT_EQ(made.table->style_info.show_row_stripes, excel.table->style_info.show_row_stripes);
  EXPECT_EQ(made.table->style_info.show_column_stripes, excel.table->style_info.show_column_stripes);
}

struct DetailCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
};

/// Loads `file`, removes its first `excel_sheets` sheets (the detail sheets,
/// in the order Show Details made them), runs Show Details on `cells` in
/// order, and compares each new sheet with the one Excel made.
void Replay(const char* file, std::size_t excel_sheets, const std::vector<DetailCell>& cells) {
  const std::vector<std::uint8_t> bytes = ReadFixture(file);
  ASSERT_FALSE(bytes.empty()) << file;
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  auto loaded = std::make_unique<io::OoxmlReadResult>(std::move(read.value()));
  Workbook& wb = loaded->workbook;
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_EQ(cells.size(), excel_sheets);

  std::vector<DetailSheet> excel;
  for (std::size_t i = 0; i < excel_sheets; ++i) {
    excel.push_back(Observe(wb, i));
  }
  const std::uint32_t excel_active = wb.active_sheet();
  for (std::size_t i = excel_sheets; i-- > 0U;) {
    ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(static_cast<std::uint32_t>(i))));
  }
  ASSERT_TRUE(wb.tables().empty());

  std::size_t pivot_sheet = 0;
  while (pivot_sheet < wb.sheet_count() && wb.sheet(pivot_sheet).pivot_tables().empty()) {
    ++pivot_sheet;
  }
  ASSERT_LT(pivot_sheet, wb.sheet_count());

  for (std::size_t i = 0; i < cells.size(); ++i) {
    auto shown = wb.show_pivot_details(pivot_sheet, 0U, cells[i].row, cells[i].col);
    ASSERT_TRUE(static_cast<bool>(shown)) << shown.error().message;
    EXPECT_EQ(shown.value(), pivot_sheet);
    ++pivot_sheet;
  }
  for (std::size_t i = 0; i < excel_sheets; ++i) {
    ExpectSameDetail(Observe(wb, excel[i].index), excel[i]);
  }
  EXPECT_EQ(wb.active_sheet(), excel_active);
}

TEST(PivotShowDetailsFixture, LeafCell) {
  Replay("show_details_leaf.xlsx", 1U, {{3U, 1U}});  // B4
}

TEST(PivotShowDetailsFixture, SubtotalThenGrandTotal) {
  Replay("show_details_subtotal_and_grand.xlsx", 2U, {{7U, 2U}, {11U, 2U}});  // C8, C12
}

TEST(PivotShowDetailsFixture, TwoLeafCells) {
  Replay("show_details_twice.xlsx", 2U, {{3U, 1U}, {4U, 1U}});  // B4, B5
}

}  // namespace
}  // namespace formulon
