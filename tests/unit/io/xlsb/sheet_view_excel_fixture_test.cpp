// XLSB worksheet-view records (BrtBeginWsView / BrtPane / BrtSel) against
// Excel-saved fixtures: each `.xlsb` has an `.xlsx` twin Excel saved from the
// same workbook, so the OOXML reader supplies the expected model.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/sqref_text.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "support/ooxml_package_fixture.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
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
    EXPECT_EQ(got.frozen_split, want.frozen_split);
    EXPECT_EQ(got.split_x, want.split_x);
    EXPECT_EQ(got.split_y, want.split_y);
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

// `package` with the part `name` replaced by `bytes`.
std::vector<std::uint8_t> ReplacePart(const std::vector<std::uint8_t>& package, const std::string& name,
                                      const std::vector<std::uint8_t>& bytes) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  mz_zip_archive writer{};
  EXPECT_TRUE(mz_zip_writer_init_heap(&writer, 0, 4096) != MZ_FALSE);
  for (std::size_t i = 0; i < zip.entry_count(); ++i) {
    const std::string entry(zip.entry_name(i));
    auto part_or = zip.read_entry(entry);
    EXPECT_TRUE(static_cast<bool>(part_or));
    const std::vector<std::uint8_t>& part = entry == name ? bytes : part_or.value();
    EXPECT_TRUE(mz_zip_writer_add_mem(&writer, entry.c_str(), part.data(), part.size(),
                                      static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)) != MZ_FALSE);
  }
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_TRUE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size) != MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<std::uint8_t*>(archive),
                                static_cast<std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  mz_zip_writer_end(&writer);
  return out;
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

void ExpectWriterMatchesExcel(const std::string& stem, const std::vector<std::size_t>& sheets,
                              const Workbook* model = nullptr) {
  const Twin twin = LoadTwin(stem);
  auto written_or = io::xlsb::write_xlsb(model != nullptr ? *model : twin.xlsx);
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

TEST(XlsbSheetViewExcelFixture, ReaderMatchesTheXlsxTwinForOneAxisSplitsAndFreezes) {
  ExpectSameViews(LoadTwin("sheet_view_split_edge"));
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

TEST(XlsbSheetViewExcelFixture, WriterEmitsExcelsViewRecordsForTheSelectionFixture) {
  ExpectWriterMatchesExcel("sheet_view_selection", {0U, 1U, 2U, 3U});
}

TEST(XlsbSheetViewExcelFixture, WriterEmitsExcelsViewRecordsForFrozenStates) {
  ExpectWriterMatchesExcel("sheet_view_frozen", {0U, 1U, 2U, 3U, 4U});
}

// The `<sheetViews>` element of each worksheet part.
std::string SheetViewsXml(const std::vector<std::uint8_t>& package, std::size_t sheet) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  auto part_or = zip.read_entry("xl/worksheets/sheet" + std::to_string(sheet + 1U) + ".xml");
  EXPECT_TRUE(static_cast<bool>(part_or));
  const std::string xml(part_or.value().begin(), part_or.value().end());
  const std::size_t begin = xml.find("<sheetViews>");
  const std::size_t end = xml.find("</sheetViews>");
  if (begin == std::string::npos || end == std::string::npos) {
    return std::string();
  }
  return xml.substr(begin, end + std::string("</sheetViews>").size() - begin);
}

void ExpectXlsxWriterMatchesExcel(const std::string& stem, const Workbook* model = nullptr) {
  const Twin twin = LoadTwin(stem);
  auto written_or = io::write_ooxml(model != nullptr ? *model : twin.xlsx);
  ASSERT_TRUE(static_cast<bool>(written_or)) << written_or.error().message;
  const std::vector<std::uint8_t> excel = test::read_file_bytes(FixturePath(stem + ".xlsx"));
  for (std::size_t sheet = 0; sheet < twin.xlsx.sheet_count(); ++sheet) {
    SCOPED_TRACE("sheet " + std::to_string(sheet));
    EXPECT_EQ(SheetViewsXml(written_or.value(), sheet), SheetViewsXml(excel, sheet));
  }
}

TEST(XlsbSheetViewExcelFixture, XlsxWriterEmitsExcelsSheetViewsForTheSelectionFixture) {
  ExpectXlsxWriterMatchesExcel("sheet_view_selection");
}

TEST(XlsbSheetViewExcelFixture, XlsxWriterEmitsExcelsSheetViewsForFrozenStates) {
  ExpectXlsxWriterMatchesExcel("sheet_view_frozen");
}

// Excel drops the flag of a cell that is a formula when the file loads; a
// constant keeps it.
TEST(XlsbSheetViewExcelFixture, ReadersDropThePhoneticFlagOfAFormulaCell) {
  const std::vector<std::uint8_t> xlsx = test::BuildOoxmlPackage(
      test::OoxmlContentTypes(test::kXlsxWorkbookContentType),
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\">"
      "<c r=\"A1\" ph=\"1\"><v>1</v></c><c r=\"B1\" ph=\"1\"><f>1+1</f><v>2</v></c>"
      "<c r=\"C1\" ph=\"1\"><f t=\"dataTable\" ref=\"C1\" dt2D=\"0\" dtr=\"0\" r1=\"A1\"/><v>3</v></c>"
      "</row></sheetData></worksheet>");
  auto from_xlsx = io::read_ooxml(test::span_of(xlsx));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  EXPECT_TRUE(from_xlsx.value().workbook.sheet(0).cell_at(0, 0)->show_phonetic);
  EXPECT_FALSE(from_xlsx.value().workbook.sheet(0).cell_at(0, 1)->show_phonetic);
  const Cell* data_table = from_xlsx.value().workbook.sheet(0).cell_at(0, 2);
  EXPECT_TRUE(data_table == nullptr || !data_table->show_phonetic);

  // The writer never sets fPhShow on a formula record, so set it by hand.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=1+1")));
  wb.sheet(0).set_cell_show_phonetic(0U, 0U, true);
  auto written_or = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written_or)) << written_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(test::span_of(written_or.value()))));
  std::vector<std::uint8_t> part = zip.read_entry("xl/worksheets/sheet1.bin").value();
  io::ByteSpan records = test::span_of(part);
  bool patched = false;
  while (records.size != 0U) {
    auto rec_or = io::xlsb::read_record(records);
    ASSERT_TRUE(static_cast<bool>(rec_or));
    const std::uint16_t type = rec_or.value().type;
    if (type >= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString) &&
        type <= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaError)) {
      // Cell header: column (4 bytes), iStyleRef (3), then fPhShow.
      part[static_cast<std::size_t>(rec_or.value().payload.data - part.data()) + 7U] |= 0x01U;
      patched = true;
    }
  }
  ASSERT_TRUE(patched);
  auto from_xlsb =
      io::xlsb::read_xlsb(test::span_of(ReplacePart(written_or.value(), "xl/worksheets/sheet1.bin", part)));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  EXPECT_TRUE(from_xlsb.value().workbook.sheet(0).cell_at(0, 0)->show_phonetic);
  EXPECT_FALSE(from_xlsb.value().workbook.sheet(0).cell_at(0, 1)->show_phonetic);
}

TEST(XlsbSheetViewExcelFixture, WritersEmitExcelsViewsForOneAxisSplitsAndFreezes) {
  ExpectWriterMatchesExcel("sheet_view_split_edge", {0U, 1U, 2U, 3U});
  ExpectXlsxWriterMatchesExcel("sheet_view_split_edge");
}

// Excel moves an active cell outside every range to the first range's
// top-left (sheet 0 was authored as H1 in A1:B2), and refuses a file with a
// selection in a pane the view lacks (sheet 3 freezes columns only).
TEST(XlsbSheetViewExcelFixture, WritersNormaliseTheActiveCellAndDropMissingPanes) {
  Twin twin = LoadTwin("sheet_view_split_edge");
  twin.xlsx.sheet(0).mutable_view().selections = {{"topLeft", "H1", "A1:B2"}};
  std::vector<SheetSelection>& frozen = twin.xlsx.sheet(3).mutable_view().selections;
  frozen.insert(frozen.begin(), {{"bottomLeft", "A9", "A9"}, {"bottomRight", "D9", "D9"}});
  ExpectWriterMatchesExcel("sheet_view_split_edge", {0U, 3U}, &twin.xlsx);
  ExpectXlsxWriterMatchesExcel("sheet_view_split_edge", &twin.xlsx);
}

// One record per pane the view has, in Excel's pane order, the first stored
// selection of a pane winning and a pane with none defaulted.
// Authored with no selections, a frozen view active in its top-right or
// bottom-left pane, and splits active in their top-left and top-right panes,
// all scrolled to B2. Excel re-derives a frozen view's active pane and selects
// A1, the scroll origin, or the lower-right pane's origin in the empty panes.
TEST(XlsbSheetViewExcelFixture, WritersDefaultActivePanesAndSelectionsTheExcelWay) {
  ExpectSameViews(LoadTwin("sheet_view_active_pane"));
  Twin twin = LoadTwin("sheet_view_active_pane");
  const char* const authored[] = {"topRight", "bottomLeft", "topLeft", "topRight"};
  for (std::size_t sheet = 0; sheet < 4U; ++sheet) {
    SheetView& view = twin.xlsx.sheet(sheet).mutable_view();
    view.selections.clear();
    view.active_pane = authored[sheet];
  }
  ExpectWriterMatchesExcel("sheet_view_active_pane", {0U, 1U, 2U, 3U}, &twin.xlsx);
  ExpectXlsxWriterMatchesExcel("sheet_view_active_pane", &twin.xlsx);
}

TEST(XlsbSheetViewExcelFixture, WriterEmitsOneSelectionPerExistingPane) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  SheetView& view = wb.sheet(0).mutable_view();
  view.freeze_rows = 2U;
  view.freeze_cols = 1U;
  view.selections = {{"bottomRight", "C5", "C5"}, {"bottomRight", "D9", "D9"}, {"middle", "E1", "E1"}};
  auto written_or = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written_or)) << written_or.error().message;
  std::vector<std::uint32_t> panes;
  std::uint32_t bottom_right_row = 0;
  for (const auto& [type, payload] : ViewRecords(written_or.value(), 0U)) {
    if (type == static_cast<std::uint16_t>(XlsbRecordType::BrtSel)) {
      panes.push_back(payload[0]);
      if (payload[0] == 0U) {
        bottom_right_row = payload[4];
      }
    }
  }
  EXPECT_EQ(panes, (std::vector<std::uint32_t>{3U, 1U, 2U, 0U}));
  EXPECT_EQ(bottom_right_row, 4U);
}

}  // namespace
}  // namespace formulon
