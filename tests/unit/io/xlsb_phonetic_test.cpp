//
// Phonetic-guide fidelity for the MS-XLSB path, against a real Mac Excel
// 365-produced pair (`tests/fixtures/excel/xlsb_phonetic.{xlsb,xlsx}`).
//
// The fixture was authored by writing the `<rPh>` blocks through the
// OOXML writer, opening the result in Excel and letting Excel save both
// containers, so the `.xlsb` carries Excel's own encoding of the guide
// rather than one derived from this reader's assumptions. Excel's
// re-saved `.xlsx` sibling is the cross-check: the two containers must
// decode to the same runs.
//
// The fixture's `Sheet1` column A:
//   A1 = 東京都   with 「東京」→トウキョウ and 「都」→ト   (two runs)
//   A2 = 山田太郎 with 「山田」→ヤマダ  and 「太郎」→タロウ (two runs)
//   A3 = 大阪     with 「大阪」→オオサカ                    (whole-string)
//   A4 = plain    with no guide
//
// A3 is the case the binary format encodes differently from OOXML: Excel
// elides the run array entirely for a whole-string reading, so an empty
// array with a non-empty kana string is one run rather than none.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "phonetic.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string XlsbPath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_phonetic.xlsb";
}
std::string XlsxTwinPath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_phonetic.xlsx";
}

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
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

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

/// The runs attached to `(row, 0)` of the first sheet, flattened to a
/// comparable form so a mismatch prints the whole annotation.
std::string DescribeRuns(const Workbook& wb, std::uint32_t row) {
  const Cell* cell = wb.sheet(0).cell_at(row, 0);
  if (cell == nullptr) {
    return "<absent>";
  }
  std::string out;
  for (const PhoneticRun& run : cell->phonetic_runs) {
    out.append("[").append(std::to_string(run.sb)).append(",").append(std::to_string(run.eb)).append(")=");
    out.append(run.text).append(" ");
  }
  return out;
}

std::string CellText(const Workbook& wb, std::uint32_t row) {
  const Cell* cell = wb.sheet(0).cell_at(row, 0);
  if (cell == nullptr || !cell->cached_value.is_text()) {
    return {};
  }
  return std::string(cell->cached_value.as_text());
}

Workbook LoadXlsb() {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(XlsbPath());
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  auto result_or = io::xlsb::read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(result_or)) << "read_xlsb failed: " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

TEST(XlsbPhonetic, ReadsExcelsEncodingOfEveryRun) {
  Workbook wb = LoadXlsb();
  ASSERT_GE(wb.sheet_count(), 1U);

  EXPECT_EQ(CellText(wb, 0), "東京都");
  EXPECT_EQ(DescribeRuns(wb, 0), "[0,2)=トウキョウ [2,3)=ト ");
  EXPECT_EQ(CellText(wb, 1), "山田太郎");
  EXPECT_EQ(DescribeRuns(wb, 1), "[0,2)=ヤマダ [2,4)=タロウ ");
  // Excel wrote no run array for this one; the reader has to recover the
  // implied whole-string span rather than dropping the reading.
  EXPECT_EQ(CellText(wb, 2), "大阪");
  EXPECT_EQ(DescribeRuns(wb, 2), "[0,2)=オオサカ ");
  EXPECT_EQ(CellText(wb, 3), "plain");
  EXPECT_EQ(DescribeRuns(wb, 3), "");
}

TEST(XlsbPhonetic, AgreesWithTheOoxmlTwinOfTheSameWorkbook) {
  const std::vector<std::uint8_t> xlsx_bytes = ReadFileBytes(XlsxTwinPath());
  ASSERT_FALSE(xlsx_bytes.empty());
  auto xlsx_or = io::read_ooxml(SpanOf(xlsx_bytes));
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << (xlsx_or ? "" : xlsx_or.error().message);
  Workbook& from_xlsx = xlsx_or.value().workbook;
  Workbook from_xlsb = LoadXlsb();

  for (std::uint32_t row = 0; row < 4U; ++row) {
    EXPECT_EQ(CellText(from_xlsb, row), CellText(from_xlsx, row)) << "row=" << row;
    EXPECT_EQ(DescribeRuns(from_xlsb, row), DescribeRuns(from_xlsx, row)) << "row=" << row;
  }
}

TEST(XlsbPhonetic, SurvivesAWriteReadCycleThroughTheBinaryContainer) {
  Workbook wb = LoadXlsb();
  ASSERT_GE(wb.sheet_count(), 1U);

  auto written_or = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written_or)) << (written_or ? "" : written_or.error().message);
  auto reread_or = io::xlsb::read_xlsb(SpanOf(written_or.value()));
  ASSERT_TRUE(static_cast<bool>(reread_or)) << (reread_or ? "" : reread_or.error().message);
  Workbook& back = reread_or.value().workbook;

  for (std::uint32_t row = 0; row < 4U; ++row) {
    EXPECT_EQ(CellText(back, row), CellText(wb, row)) << "row=" << row;
    EXPECT_EQ(DescribeRuns(back, row), DescribeRuns(wb, row)) << "row=" << row;
  }
}

TEST(XlsbPhonetic, KeepsTwoReadingsOfTheSameSurfaceTextApart) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0, 0, 0, "東京都")));
  wb.sheet(0).set_cell_phonetic_runs(0, 0, {{0U, 2U, "トウキョウ"}, {2U, 3U, "ト"}});
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0, 1, 0, "東京都")));
  wb.sheet(0).set_cell_phonetic_runs(1, 0, {{0U, 3U, "ヒガシキョウト"}});
  // Same text again, this time unannotated: the shared-string table has to
  // hold three distinct entries for one surface string.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0, 2, 0, "東京都")));

  auto written_or = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written_or)) << (written_or ? "" : written_or.error().message);
  auto reread_or = io::xlsb::read_xlsb(SpanOf(written_or.value()));
  ASSERT_TRUE(static_cast<bool>(reread_or)) << (reread_or ? "" : reread_or.error().message);
  Workbook& back = reread_or.value().workbook;

  EXPECT_EQ(DescribeRuns(back, 0), "[0,2)=トウキョウ [2,3)=ト ");
  EXPECT_EQ(DescribeRuns(back, 1), "[0,3)=ヒガシキョウト ");
  EXPECT_EQ(DescribeRuns(back, 2), "");
}

}  // namespace
}  // namespace formulon
