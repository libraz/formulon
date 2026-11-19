//
// Tracks two places where the engine's conditional-format sqref output is
// known to differ from Excel's, against Excel-saved packages
// (`tests/fixtures/excel/cf_mixed_part_split/`,
// `tests/fixtures/excel/cf_far_edge_keep_whole/`).
//
// Mixed part: Excel splits one `<conditionalFormatting>` holding a merging
// rule (expression / cellIs) and another kind into two blocks on save; the
// engine keeps one block. The `_one_part` inputs are the authored books, as
// Excel never writes such a part.
//
// Far edge: a delete whose cells lie on a rectangle's far line, short of the
// whole line, keeps some rules' rectangle whole in Excel; the engine splits
// every kind. Where Excel splits, it writes the full-width bands before the
// shortened band, which the engine places in row order. The `_before`
// packages are Excel's save of the authored books, `_after` its save after
// the edit.
//
// Each differing case pins the engine's current text and fails once the
// engine produces Excel's; rows where the two already agree must keep
// agreeing.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "miniz.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::string ZipEntry(const std::vector<std::uint8_t>& zip, const std::string& name) {
  mz_zip_archive archive{};
  if (mz_zip_reader_init_mem(&archive, zip.data(), zip.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "not a zip package";
    return {};
  }
  std::size_t size = 0;
  void* data = mz_zip_reader_extract_file_to_heap(&archive, name.c_str(), &size, 0);
  std::string out;
  if (data != nullptr) {
    out.assign(static_cast<const char*>(data), size);
    mz_free(data);
  } else {
    ADD_FAILURE() << "missing part " << name;
  }
  mz_zip_reader_end(&archive);
  return out;
}

// Every block's sqref in file order, legacy blocks first, then x14 blocks,
// joined by " || ".
std::string BlockText(const std::string& xml) {
  std::string out;
  const auto append = [&out](std::string_view piece) {
    out += out.empty() ? std::string(piece) : " || " + std::string(piece);
  };
  for (std::size_t at = xml.find("<conditionalFormatting "); at != std::string::npos;
       at = xml.find("<conditionalFormatting ", at + 1U)) {
    const std::size_t end = xml.find('>', at);
    const std::size_t attr = xml.find("sqref=\"", at);
    if (attr == std::string::npos || attr > end) {
      continue;
    }
    const std::size_t first = attr + 7U;
    append(std::string_view(xml).substr(first, xml.find('"', first) - first));
  }
  for (std::size_t at = xml.find("<xm:sqref>"); at != std::string::npos; at = xml.find("<xm:sqref>", at + 1U)) {
    const std::size_t first = at + 10U;
    append(std::string_view(xml).substr(first, xml.find('<', first) - first));
  }
  return out;
}

std::string SheetPart(std::size_t index) {
  return "xl/worksheets/sheet" + std::to_string(index + 1U) + ".xml";
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto bytes = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes));
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

struct MixedCase {
  const char* stem;
  const char* excel;   // Excel's saved blocks
  const char* engine;  // the engine's current save
};

// Excel's order: the merging rules form one block (joined, sorted), the other
// rules a second block; the engine writes the authored part unchanged.
constexpr MixedCase kMixed[] = {
    {"cellis_colorscale3", "B5:C6 H2 E3 || H2 E3 B5:C6", "H2 B5:C6 E3"},
    {"expression_iconset", "H2 B5:C6 E3 || H2 E3 B5:C6", "H2 B5:C6 E3"},
    {"expression_top10", "B2:B3 C4:D4 || C4:D4 B2 B3", "C4:D4 B2 B3"},
    {"expression_aboveaverage", "B2 C2 D2 || B2:D2", "B2 C2 D2"},
};

TEST(CfSqrefExcelGapFixture, AMixedPartStaysOneBlockWhereExcelWritesTwo) {
  for (const MixedCase& c : kMixed) {
    SCOPED_TRACE(c.stem);
    const std::string dir = "cf_mixed_part_split/";
    const std::string excel = BlockText(ZipEntry(ReadFixture(dir + c.stem + "_excel_saved.xlsx"), SheetPart(0)));
    ASSERT_EQ(excel, c.excel);
    const std::string engine =
        BlockText(ZipEntry(Save(Load(ReadFixture(dir + c.stem + "_one_part.xlsx"))), SheetPart(0)));
    EXPECT_EQ(engine, c.engine);
    EXPECT_NE(engine, excel) << "the engine now matches Excel: replace this case with an equality pin";
  }
}

struct FarEdgeBook {
  const char* stem;
  const char* edit;  // the deleted cells on the content sheets
  CellShift shift;
};

struct FarEdgeRow {
  const char* sheet;
  const char* excel;   // Excel's saved blocks after the edit
  const char* engine;  // the engine's current save after the same edit
};

MergeRange Rect(std::string_view a1) {
  const auto cell = [](std::string_view s) {
    std::uint32_t col = 0;
    std::size_t i = 0;
    for (; i < s.size() && s[i] >= 'A' && s[i] <= 'Z'; ++i) {
      col = col * 26U + static_cast<std::uint32_t>(s[i] - 'A' + 1);
    }
    std::uint32_t row = 0;
    for (; i < s.size(); ++i) {
      row = row * 10U + static_cast<std::uint32_t>(s[i] - '0');
    }
    return std::pair<std::uint32_t, std::uint32_t>{row - 1U, col - 1U};
  };
  const std::size_t colon = a1.find(':');
  const auto first = cell(a1.substr(0, colon));
  const auto last = colon == std::string_view::npos ? first : cell(a1.substr(colon + 1U));
  return MergeRange{first.first, first.second, last.first, last.second};
}

void CheckFarEdge(const FarEdgeBook& book, const std::vector<FarEdgeRow>& rows) {
  const std::string dir = "cf_far_edge_keep_whole/";
  const std::vector<std::uint8_t> after = ReadFixture(dir + book.stem + "_after.xlsx");
  Workbook wb = Load(ReadFixture(dir + book.stem + "_before.xlsx"));
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).name().size() == 3U && wb.sheet(i).name()[0] != 'K' && wb.sheet(i).name() != "Src") {
      ASSERT_TRUE(static_cast<bool>(wb.delete_cells(i, Rect(book.edit), book.shift))) << wb.sheet(i).name();
    }
  }
  const std::vector<std::uint8_t> saved = Save(wb);
  for (const FarEdgeRow& row : rows) {
    SCOPED_TRACE(row.sheet);
    std::size_t index = wb.sheet_count();
    for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
      if (wb.sheet(i).name() == row.sheet) {
        index = i;
      }
    }
    ASSERT_LT(index, wb.sheet_count());
    const std::string excel = BlockText(ZipEntry(after, SheetPart(index)));
    ASSERT_EQ(excel, row.excel);
    const std::string engine = BlockText(ZipEntry(saved, SheetPart(index)));
    EXPECT_EQ(engine, row.engine);
    if (std::string_view(row.engine) != row.excel) {
      EXPECT_NE(engine, excel) << "the engine now matches Excel: make this row an equality pin";
    }
  }
}

// F6:H9 with G9 deleted upward.
TEST(CfSqrefExcelGapFixture, BottomLineCellDeletedUpward) {
  CheckFarEdge({"bottom_line_cell_up", "G9", CellShift::kUp}, {
                                                                  {"A01", "F6:H9", "F6:H8 F9 H9"},
                                                                  {"A02", "F6:H9", "F6:H8 F9 H9"},
                                                                  {"A03", "F6:H9", "F6:H8 F9 H9"},
                                                                  {"A04", "F6:H9", "F6:H8 F9 H9"},
                                                                  {"A05", "F6:H8 F9 H9", "F6:H8 F9 H9"},
                                                                  {"A06", "F6:H9", "F6:H8 F9 H9"},
                                                                  {"A07", "F6:H8 F9 H9", "F6:H8 F9 H9"},
                                                                  {"A08", "F6:H9", "F6:H8 F9 H9"},
                                                                  {"A09", "F6:H8 F9 H9", "F6:H8 F9 H9"},
                                                                  {"A10", "F6:H8 F9 H9", "F6:H8 F9 H9"},
                                                                  {"A11", "F6:H8 F9 H9", "F6:H8 F9 H9"},
                                                                  {"A12", "F6:H8 F9 H9", "F6:H8 F9 H9"},
                                                                  {"A13", "F6:H9", "F6:H8 F9 H9"},
                                                                  {"A14", "F6:H9", "F6:H8 F9 H9"},
                                                                  {"A15", "F6:H9", "F6:H8 F9 H9"},
                                                              });
}

// B10:E11 with E11 deleted leftward.
TEST(CfSqrefExcelGapFixture, RightLineCornerDeletedLeftward) {
  CheckFarEdge({"right_line_corner_left", "E11", CellShift::kLeft}, {
                                                                        {"B01", "B10:E11", "B10:E10 B11:D11"},
                                                                        {"B02", "B10:E11", "B10:E10 B11:D11"},
                                                                        {"B03", "B10:E11", "B10:E10 B11:D11"},
                                                                        {"B04", "B10:E11", "B10:E10 B11:D11"},
                                                                        {"B05", "B10:E10 B11:D11", "B10:E10 B11:D11"},
                                                                        {"B06", "B10:E11", "B10:E10 B11:D11"},
                                                                        {"B07", "B10:E10 B11:D11", "B10:E10 B11:D11"},
                                                                        {"B08", "B10:E11", "B10:E10 B11:D11"},
                                                                        {"B09", "B10:E10 B11:D11", "B10:E10 B11:D11"},
                                                                        {"B10", "B10:E10 B11:D11", "B10:E10 B11:D11"},
                                                                        {"B11", "B10:E10 B11:D11", "B10:E10 B11:D11"},
                                                                        {"B12", "B10:E10 B11:D11", "B10:E10 B11:D11"},
                                                                        {"B13", "B10:E11", "B10:E10 B11:D11"},
                                                                        {"B14", "B10:E11", "B10:E10 B11:D11"},
                                                                        {"B15", "B10:E11", "B10:E10 B11:D11"},
                                                                    });
}

// D4:F8 with F6 deleted leftward; where the rule splits, Excel writes the
// full-width bands first and the shortened band last.
TEST(CfSqrefExcelGapFixture, RightLineMiddleCellDeletedLeftward) {
  CheckFarEdge({"right_line_middle_left", "F6", CellShift::kLeft},
               {
                   {"E01", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E02", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E03", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E04", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E05", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E06", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E07", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E08", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E09", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E10", "D4:F5 D7:F8 D6:E6", "D4:F5 D6:E6 D7:F8"},
                   {"E11", "D4:F5 D7:F8 D6:E6", "D4:F5 D6:E6 D7:F8"},
                   {"E12", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E13", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E14", "D4:F8", "D4:F5 D6:E6 D7:F8"},
                   {"E15", "D4:F5 D7:F8 D6:E6", "D4:F5 D6:E6 D7:F8"},
                   {"E16", "D4:F5 D7:F8 D6:E6", "D4:F5 D6:E6 D7:F8"},
                   {"E17", "D4:F8", "D4:F5 D6:E6 D7:F8"},
               });
}

}  // namespace
}  // namespace formulon
