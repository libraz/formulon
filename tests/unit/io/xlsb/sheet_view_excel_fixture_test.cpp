// XLSB worksheet-view records (BrtBeginWsView / BrtPane / BrtSel) against
// Excel-saved fixtures: each `.xlsb` has an `.xlsx` twin Excel saved from the
// same workbook, so the OOXML reader supplies the expected model.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/sqref_text.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

namespace formulon {
namespace {

using io::xlsb::XlsbRecordType;

std::string FixturePath(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
}

struct Twin {
  Workbook xlsb;
  Workbook xlsx;
};

Twin LoadTwin(const std::string& stem) {
  const std::vector<std::uint8_t> xlsb_bytes = test::read_file_bytes(FixturePath(stem + ".xlsb"));
  const std::vector<std::uint8_t> xlsx_bytes = test::read_file_bytes(FixturePath(stem + ".xlsx"));
  auto xlsb_or = io::xlsb::read_xlsb(test::span_of(xlsb_bytes));
  auto xlsx_or = io::read_ooxml(test::span_of(xlsx_bytes));
  EXPECT_TRUE(static_cast<bool>(xlsb_or)) << (xlsb_or ? "" : xlsb_or.error().message);
  EXPECT_TRUE(static_cast<bool>(xlsx_or)) << (xlsx_or ? "" : xlsx_or.error().message);
  return Twin{std::move(xlsb_or.value().workbook), std::move(xlsx_or.value().workbook)};
}

std::string NormalisedSqref(const SheetSelection& selection) {
  std::vector<MergeRange> ranges;
  io::parse_selection_sqref_text(selection.sqref, ranges);
  return io::format_sqref_text(ranges);
}

void ExpectSameViews(const Twin& twin) {
  ASSERT_EQ(twin.xlsb.sheet_count(), twin.xlsx.sheet_count());
  EXPECT_EQ(twin.xlsb.active_sheet(), twin.xlsx.active_sheet());
  for (std::size_t i = 0; i < twin.xlsx.sheet_count(); ++i) {
    SCOPED_TRACE("sheet " + std::to_string(i));
    const SheetView& got = twin.xlsb.sheet(i).view();
    const SheetView& want = twin.xlsx.sheet(i).view();
    EXPECT_EQ(got.freeze_rows, want.freeze_rows);
    EXPECT_EQ(got.freeze_cols, want.freeze_cols);
    EXPECT_EQ(got.pane_top_left_cell, want.pane_top_left_cell);
    EXPECT_EQ(got.active_pane, want.active_pane);
    EXPECT_EQ(got.top_left_cell, want.top_left_cell);
    EXPECT_EQ(got.tab_selected, want.tab_selected);
    ASSERT_EQ(got.selections.size(), want.selections.size());
    for (std::size_t k = 0; k < want.selections.size(); ++k) {
      EXPECT_EQ(SheetView::pane_key(got.selections[k].pane), SheetView::pane_key(want.selections[k].pane));
      EXPECT_EQ(got.selections[k].active_cell, want.selections[k].active_cell);
      EXPECT_EQ(NormalisedSqref(got.selections[k]), NormalisedSqref(want.selections[k]));
    }
  }
}

// The (type, payload) records from BrtBeginWsView through BrtEndWsView.
std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>> ViewRecords(const std::vector<std::uint8_t>& package,
                                                                             std::size_t sheet) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  auto part_or = zip.read_entry("xl/worksheets/sheet" + std::to_string(sheet + 1U) + ".bin");
  EXPECT_TRUE(static_cast<bool>(part_or));
  std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>> out;
  io::ByteSpan records = test::span_of(part_or.value());
  bool in_view = false;
  while (records.size != 0U) {
    auto rec_or = io::xlsb::read_record(records);
    EXPECT_TRUE(static_cast<bool>(rec_or));
    if (!rec_or) {
      break;
    }
    const auto& rec = rec_or.value();
    in_view = in_view || rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtBeginWsView);
    if (in_view) {
      out.emplace_back(rec.type, std::vector<std::uint8_t>(rec.payload.data, rec.payload.data + rec.payload.size));
    }
    if (rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtEndWsView)) {
      break;
    }
  }
  return out;
}

void ExpectWriterMatchesExcel(const std::string& stem, const std::vector<std::size_t>& sheets) {
  const Twin twin = LoadTwin(stem);
  auto written_or = io::xlsb::write_xlsb(twin.xlsx);
  ASSERT_TRUE(static_cast<bool>(written_or)) << written_or.error().message;
  const std::vector<std::uint8_t> excel = test::read_file_bytes(FixturePath(stem + ".xlsb"));
  for (std::size_t sheet : sheets) {
    SCOPED_TRACE("sheet " + std::to_string(sheet));
    EXPECT_EQ(ViewRecords(written_or.value(), sheet), ViewRecords(excel, sheet));
  }
}

TEST(XlsbSheetViewExcelFixture, ReaderMatchesTheXlsxTwinForScrollSelectionAndSplit) {
  ExpectSameViews(LoadTwin("sheet_view_selection"));
}

TEST(XlsbSheetViewExcelFixture, ReaderMatchesTheXlsxTwinForFrozenStates) {
  ExpectSameViews(LoadTwin("sheet_view_frozen"));
}

TEST(XlsbSheetViewExcelFixture, ShowPhoneticFlagsMatchTheXlsxTwin) {
  const Twin twin = LoadTwin("sheet_view_selection");
  for (std::uint32_t col = 0; col < 4U; ++col) {
    const Cell* got = twin.xlsb.sheet(0).cell_at(0, col);
    const Cell* want = twin.xlsx.sheet(0).cell_at(0, col);
    ASSERT_NE(got, nullptr);
    ASSERT_NE(want, nullptr);
    EXPECT_EQ(got->show_phonetic, want->show_phonetic) << "col " << col;
  }
  EXPECT_TRUE(twin.xlsx.sheet(0).cell_at(0, 0)->show_phonetic);
  EXPECT_FALSE(twin.xlsx.sheet(0).cell_at(0, 1)->show_phonetic);
}

// Left out of the byte comparison: the selection fixture's sheet 2, an
// unfrozen split view, and the frozen fixture's sheet 1, a `frozenSplit`
// pane; the model carries neither state.
TEST(XlsbSheetViewExcelFixture, WriterEmitsExcelsViewRecordsForTheSelectionFixture) {
  ExpectWriterMatchesExcel("sheet_view_selection", {0U, 1U, 3U});
}

TEST(XlsbSheetViewExcelFixture, WriterEmitsExcelsViewRecordsForFrozenStates) {
  ExpectWriterMatchesExcel("sheet_view_frozen", {0U, 2U, 3U, 4U});
}

TEST(XlsbSheetViewExcelFixture, WriterEmitsOneSelectionPerKnownPane) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  SheetView& view = wb.sheet(0).mutable_view();
  view.freeze_rows = 2U;
  view.freeze_cols = 1U;
  view.selections = {{"bottomRight", "C5", "C5"}, {"bottomRight", "D9", "D9"}, {"middle", "E1", "E1"}};
  auto written_or = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written_or)) << written_or.error().message;
  std::vector<std::uint32_t> panes;
  for (const auto& [type, payload] : ViewRecords(written_or.value(), 0U)) {
    if (type == static_cast<std::uint16_t>(XlsbRecordType::BrtSel)) {
      panes.push_back(payload[0]);
    }
  }
  EXPECT_EQ(panes, (std::vector<std::uint32_t>{3U, 0U}));
}

}  // namespace
}  // namespace formulon
