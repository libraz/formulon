// A chartsheet's view (tab selection and zoom) and its print settings are read from its part. The view is written
// back into it: only the first `<sheetView>` / BrtBeginCsView changes, every other byte of the part is kept.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/chartsheet_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "sheet_layout.h"
#include "styles.h"
#include "workbook.h"

namespace formulon {
namespace {

using io::xlsb::ChartsheetView;

std::vector<std::uint8_t> ReadFixture(const std::string& relative) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + relative;
  std::ifstream in(path, std::ios::binary);
  EXPECT_TRUE(in.good()) << path;
  return std::vector<std::uint8_t>((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::uint8_t> Part(const std::vector<std::uint8_t>& package, const char* path) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry_or = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(entry_or)) << path;
  return entry_or ? std::move(entry_or.value()) : std::vector<std::uint8_t>{};
}

std::string Text(const std::vector<std::uint8_t>& bytes) {
  return std::string(reinterpret_cast<const char*>(bytes.data()), bytes.size());
}

Workbook LoadXlsx(const std::vector<std::uint8_t>& bytes) {
  auto read_or = io::read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  return std::move(read_or.value().workbook);
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto read_or = io::xlsb::read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  return std::move(read_or.value().workbook);
}

std::vector<std::uint8_t> SaveXlsx(const Workbook& wb) {
  auto bytes_or = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;
  return bytes_or ? std::move(bytes_or.value()) : std::vector<std::uint8_t>{};
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto bytes_or = io::xlsb::write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;
  return bytes_or ? std::move(bytes_or.value()) : std::vector<std::uint8_t>{};
}

void SetZoom(Workbook& wb, std::size_t sheet, std::uint32_t zoom) {
  wb.sheet(sheet).mutable_view().zoom_scale = zoom;
  wb.sheet(sheet).mutable_view().zoom_to_fit = false;
}

// Excel-saved: Sheet2, then chartsheets グラフ3 / グラフ2 / グラフ1 (parts sheet1..3.xml), each fit to the window
// at 125%, then Sheet1. The zoom80 twin is the same book after Excel set グラフ3's zoom to 80%.
constexpr const char* kFit = "chartsheet_zoom/fit_ja_JP.xlsx";
constexpr const char* kZoom80 = "chartsheet_zoom/zoom80_ja_JP.xlsx";
constexpr const char* kChartsheetParts[] = {"xl/chartsheets/sheet1.xml", "xl/chartsheets/sheet2.xml",
                                            "xl/chartsheets/sheet3.xml"};

TEST(ChartsheetViewRoundtrip, XlsxReadsTheStoredZoom) {
  const Workbook wb = LoadXlsx(ReadFixture(kFit));
  for (std::size_t i = 1; i <= 3; ++i) {
    EXPECT_EQ(wb.sheet(i).view().zoom_scale, 125U) << i;
    EXPECT_TRUE(wb.sheet(i).view().zoom_to_fit) << i;
  }
  EXPECT_FALSE(wb.sheet(0).view().zoom_to_fit);
}

TEST(ChartsheetViewRoundtrip, XlsxZoomSetWritesWhatExcelWrites) {
  const std::vector<std::uint8_t> source = ReadFixture(kFit);
  Workbook wb = LoadXlsx(source);
  SetZoom(wb, 1, 80U);
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);

  // Excel's part also selects the tab it was zoomed on; this save leaves the selection where it was.
  std::string expected = Text(Part(ReadFixture(kZoom80), kChartsheetParts[0]));
  const std::string selected = " tabSelected=\"1\"";
  ASSERT_NE(expected.find(selected), std::string::npos);
  expected.erase(expected.find(selected), selected.size());
  EXPECT_EQ(Text(Part(saved, kChartsheetParts[0])), expected);
  EXPECT_EQ(Part(saved, kChartsheetParts[1]), Part(source, kChartsheetParts[1]));
  EXPECT_EQ(Part(saved, kChartsheetParts[2]), Part(source, kChartsheetParts[2]));

  const Workbook reread = LoadXlsx(saved);
  EXPECT_EQ(reread.sheet(1).view().zoom_scale, 80U);
  EXPECT_FALSE(reread.sheet(1).view().zoom_to_fit);
  EXPECT_TRUE(reread.sheet(2).view().zoom_to_fit);
}

TEST(ChartsheetViewRoundtrip, XlsxReadsTheStoredTabSelection) {
  EXPECT_FALSE(LoadXlsx(ReadFixture(kFit)).sheet(1).view().tab_selected);
  const Workbook zoomed = LoadXlsx(ReadFixture(kZoom80));
  EXPECT_TRUE(zoomed.sheet(1).view().tab_selected);
  EXPECT_FALSE(zoomed.sheet(2).view().tab_selected);
}

TEST(ChartsheetViewRoundtrip, XlsxActivatingAndZoomingWritesWhatExcelWrites) {
  Workbook wb = LoadXlsx(ReadFixture(kFit));
  wb.set_active_sheet(1);
  SetZoom(wb, 1, 80U);
  EXPECT_EQ(Part(SaveXlsx(wb), kChartsheetParts[0]), Part(ReadFixture(kZoom80), kChartsheetParts[0]));
}

TEST(ChartsheetViewRoundtrip, XlsxActivatingAnotherSheetDeselectsTheChartsheet) {
  const std::vector<std::uint8_t> source = ReadFixture(kZoom80);
  Workbook wb = LoadXlsx(source);
  wb.set_active_sheet(0);
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  std::string expected = Text(Part(source, kChartsheetParts[0]));
  const std::string selected = " tabSelected=\"1\"";
  ASSERT_NE(expected.find(selected), std::string::npos);
  expected.erase(expected.find(selected), selected.size());
  EXPECT_EQ(Text(Part(saved, kChartsheetParts[0])), expected);
  EXPECT_FALSE(LoadXlsx(saved).sheet(1).view().tab_selected);
}

TEST(ChartsheetViewRoundtrip, XlsbActivationRewritesOnlyTheViewFlags) {
  const std::vector<std::uint8_t> source = ReadFixture("chartsheet_twins/printer_settings.xlsb");
  const std::vector<std::uint8_t> stored = Part(source, "xl/chartsheets/sheet1.bin");
  Workbook wb = LoadXlsb(source);
  const bool was_selected = wb.sheet(1).view().tab_selected;
  wb.set_active_sheet(was_selected ? 0U : 1U);
  const std::vector<std::uint8_t> package = SaveXlsb(wb);
  const std::vector<std::uint8_t> saved = Part(package, "xl/chartsheets/sheet1.bin");
  const std::optional<ChartsheetView> view = io::xlsb::chartsheet_view_from_records(SpanOf(saved));
  ASSERT_TRUE(view.has_value());
  EXPECT_EQ(view->tab_selected, !was_selected);
  ASSERT_EQ(saved.size(), stored.size());
  std::size_t changed = 0;
  for (std::size_t i = 0; i < saved.size(); ++i) {
    changed += saved[i] != stored[i] ? 1U : 0U;
  }
  EXPECT_EQ(changed, 1U);
  EXPECT_EQ(LoadXlsb(package).sheet(1).view().tab_selected, !was_selected);
}

TEST(ChartsheetViewRoundtrip, PrintSettingsAreReadFromThePart) {
  for (const char* ext : {"xlsx", "xlsb"}) {
    SCOPED_TRACE(ext);
    const std::string suffix = std::string(".") + ext;
    const bool xlsx = suffix == ".xlsx";
    const std::vector<std::uint8_t> red = ReadFixture("chartsheet_twins/tab_color_rgb_red" + suffix);
    const Workbook wb = xlsx ? LoadXlsx(red) : LoadXlsb(red);
    const SheetPrintSettings& print = wb.sheet(1).print_settings();
    EXPECT_EQ(print.tab_color.kind, ColorSpec::Kind::kRgb);
    EXPECT_EQ(print.tab_color.rgb, 0xFFFF0000U);
    EXPECT_DOUBLE_EQ(print.page_margins.left, 0.7);
    EXPECT_DOUBLE_EQ(print.page_margins.top, 0.75);
    EXPECT_FALSE(print.page_margins_xml.empty());

    const std::vector<std::uint8_t> header = ReadFixture("chartsheet_twins/header_center_HDR" + suffix);
    const Workbook with_header = xlsx ? LoadXlsx(header) : LoadXlsb(header);
    EXPECT_NE(with_header.sheet(1).print_settings().header_footer_xml.find("HDR"), std::string::npos);
  }
}

TEST(ChartsheetViewRoundtrip, XlsxUnchangedChartsheetsResaveByteEqual) {
  for (const char* fixture : {kFit, kZoom80}) {
    SCOPED_TRACE(fixture);
    const std::vector<std::uint8_t> source = ReadFixture(fixture);
    const std::vector<std::uint8_t> saved = SaveXlsx(LoadXlsx(source));
    for (const char* part : kChartsheetParts) {
      EXPECT_EQ(Part(saved, part), Part(source, part)) << part;
    }
  }
}

TEST(ChartsheetViewRoundtrip, XlsbZoomSetRewritesOnlyTheViewRecord) {
  const std::vector<std::uint8_t> source = ReadFixture("chartsheet_twins/printer_settings.xlsb");
  const std::vector<std::uint8_t> stored = Part(source, "xl/chartsheets/sheet1.bin");
  ASSERT_EQ(io::xlsb::chartsheet_view_from_records(SpanOf(stored)), (ChartsheetView{100U, true}));

  Workbook wb = LoadXlsb(source);
  ASSERT_EQ(wb.sheet(1).kind(), SheetKind::kChartsheet);
  EXPECT_EQ(wb.sheet(1).view().zoom_scale, 100U);
  EXPECT_TRUE(wb.sheet(1).view().zoom_to_fit);
  EXPECT_EQ(Part(SaveXlsb(wb), "xl/chartsheets/sheet1.bin"), stored);

  SetZoom(wb, 1, 80U);
  const std::vector<std::uint8_t> saved = Part(SaveXlsb(wb), "xl/chartsheets/sheet1.bin");
  EXPECT_EQ(io::xlsb::chartsheet_view_from_records(SpanOf(saved)), (ChartsheetView{80U, false}));
  ASSERT_EQ(saved.size(), stored.size());
  // The flags word and the zoom are six adjacent payload bytes.
  std::size_t first = saved.size();
  std::size_t last = 0;
  for (std::size_t i = 0; i < saved.size(); ++i) {
    if (saved[i] != stored[i]) {
      first = std::min(first, i);
      last = i;
    }
  }
  ASSERT_LT(first, saved.size());
  EXPECT_LT(last - first, 6U);
}

TEST(ChartsheetViewRoundtrip, ZoomCarriesAcrossContainers) {
  // xlsx -> xlsb: the converted records state the set zoom; untouched sheets keep fit.
  Workbook from_xlsx = LoadXlsx(ReadFixture(kFit));
  SetZoom(from_xlsx, 1, 80U);
  const Workbook as_xlsb = LoadXlsb(SaveXlsb(from_xlsx));
  EXPECT_EQ(as_xlsb.sheet(1).view().zoom_scale, 80U);
  EXPECT_FALSE(as_xlsb.sheet(1).view().zoom_to_fit);
  EXPECT_EQ(as_xlsb.sheet(2).view().zoom_scale, 125U);
  EXPECT_TRUE(as_xlsb.sheet(2).view().zoom_to_fit);

  // xlsb -> xlsx: the converted `<sheetView>` states the set zoom without zoomToFit.
  Workbook from_xlsb = LoadXlsb(ReadFixture("chartsheet_twins/printer_settings.xlsb"));
  SetZoom(from_xlsb, 1, 80U);
  const std::vector<std::uint8_t> as_xlsx = SaveXlsx(from_xlsb);
  const std::string xml = Text(Part(as_xlsx, "xl/chartsheets/sheet1.xml"));
  EXPECT_NE(xml.find("<sheetView zoomScale=\"80\" workbookViewId=\"0\"/>"), std::string::npos) << xml;
  const Workbook reread = LoadXlsx(as_xlsx);
  EXPECT_EQ(reread.sheet(1).view().zoom_scale, 80U);
  EXPECT_FALSE(reread.sheet(1).view().zoom_to_fit);
}

TEST(ChartsheetViewRoundtrip, XmlRewriteTouchesOnlyTheViewAttributes) {
  const std::string fit =
      "<x:chartsheet><x:sheetPr/><x:sheetViews><x:sheetView workbookViewId=\"0\" zoomToFit=\"1\"/>"
      "</x:sheetViews></x:chartsheet>";
  EXPECT_EQ(io::xlsb::chartsheet_view_from_xml(fit), (ChartsheetView{100U, true}));
  const std::string zoomed = io::xlsb::chartsheet_xml_with_view(fit, {80U, false});
  EXPECT_EQ(zoomed,
            "<x:chartsheet><x:sheetPr/><x:sheetViews><x:sheetView zoomScale=\"80\" workbookViewId=\"0\"/>"
            "</x:sheetViews></x:chartsheet>");
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_view(zoomed, {100U, true}), fit);
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_view(fit, {100U, true}), fit);
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_view("<chartsheet/>", {80U, false}), "<chartsheet/>");
  EXPECT_EQ(io::xlsb::chartsheet_view_from_xml("<chartsheet/>"), std::nullopt);

  // tabSelected leads, zoomScale follows it, both before workbookViewId.
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_view(fit, {80U, false, true}),
            "<x:chartsheet><x:sheetPr/><x:sheetViews><x:sheetView tabSelected=\"1\" zoomScale=\"80\" "
            "workbookViewId=\"0\"/></x:sheetViews></x:chartsheet>");
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_view(io::xlsb::chartsheet_xml_with_view(fit, {80U, false, true}),
                                               {80U, false, false}),
            zoomed);
}

}  // namespace
}  // namespace formulon
