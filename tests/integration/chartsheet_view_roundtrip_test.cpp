// A chartsheet's view (tab selection and zoom), print settings and protection are read from its part. The view, the
// tab colour, the margins, the page setup, the header and footer and the protection are written back into it: only
// the first `<sheetView>` / BrtBeginCsView, the `<sheetPr>` / BrtCsProp colour and the element or records of a
// setting the model changed differ, every other byte of the part is kept.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <functional>
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
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
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

constexpr const char* kRedTwin = "chartsheet_twins/tab_color_rgb_red";
constexpr const char* kThemeTwin = "chartsheet_twins/tab_color_theme";

ColorSpec RgbColor(std::uint32_t argb) {
  ColorSpec color;
  color.kind = ColorSpec::Kind::kRgb;
  color.rgb = argb;
  return color;
}

ColorSpec ThemeColor(std::uint32_t theme) {
  ColorSpec color;
  color.kind = ColorSpec::Kind::kTheme;
  color.theme = theme;
  return color;
}

std::vector<std::uint8_t> ChartsheetPart(const std::vector<std::uint8_t>& package, bool xlsx) {
  return Part(package, xlsx ? "xl/chartsheets/sheet1.xml" : "xl/chartsheets/sheet1.bin");
}

// Loads `twin` in its own container, sets the chartsheet's tab colour and saves it in the same container.
std::vector<std::uint8_t> ResavedWithTabColor(const std::string& twin, bool xlsx, const ColorSpec& color) {
  Workbook wb = xlsx ? LoadXlsx(ReadFixture(twin)) : LoadXlsb(ReadFixture(twin));
  EXPECT_EQ(wb.sheet(1).kind(), SheetKind::kChartsheet);
  wb.sheet(1).mutable_print_settings().tab_color = color;
  return xlsx ? SaveXlsx(wb) : SaveXlsb(wb);
}

TEST(ChartsheetViewRoundtrip, TabColorSetWritesWhatExcelWrites) {
  for (const bool xlsx : {true, false}) {
    const std::string suffix = xlsx ? ".xlsx" : ".xlsb";
    SCOPED_TRACE(suffix);
    const std::vector<std::uint8_t> red = ChartsheetPart(ReadFixture(kRedTwin + suffix), xlsx);
    const std::vector<std::uint8_t> theme = ChartsheetPart(ReadFixture(kThemeTwin + suffix), xlsx);
    ASSERT_NE(red, theme);

    EXPECT_EQ(ChartsheetPart(ResavedWithTabColor(kRedTwin + suffix, xlsx, RgbColor(0xFFFF0000U)), xlsx), red);
    EXPECT_EQ(ChartsheetPart(ResavedWithTabColor(kRedTwin + suffix, xlsx, ThemeColor(4U)), xlsx), theme);
    const std::vector<std::uint8_t> saved = ResavedWithTabColor(kThemeTwin + suffix, xlsx, RgbColor(0xFFFF0000U));
    EXPECT_EQ(ChartsheetPart(saved, xlsx), red);
    const Workbook reread = xlsx ? LoadXlsx(saved) : LoadXlsb(saved);
    EXPECT_EQ(reread.sheet(1).print_settings().tab_color.kind, ColorSpec::Kind::kRgb);
    EXPECT_EQ(reread.sheet(1).print_settings().tab_color.rgb, 0xFFFF0000U);
  }
}

TEST(ChartsheetViewRoundtrip, TabColorClearWritesAnEmptySheetPr) {
  // Excel's part for a chartsheet without a tab colour: `<sheetPr/>`, BrtCsProp colour automatic.
  std::string xml = Text(ChartsheetPart(ReadFixture(kRedTwin + std::string(".xlsx")), true));
  const std::string stated = "<sheetPr><tabColor rgb=\"FFFF0000\"/></sheetPr>";
  ASSERT_NE(xml.find(stated), std::string::npos);
  xml.replace(xml.find(stated), stated.size(), "<sheetPr/>");
  const std::vector<std::uint8_t> cleared_xlsx = ResavedWithTabColor(kRedTwin + std::string(".xlsx"), true, {});
  EXPECT_EQ(Text(ChartsheetPart(cleared_xlsx, true)), xml);
  EXPECT_EQ(LoadXlsx(cleared_xlsx).sheet(1).print_settings().tab_color.kind, ColorSpec::Kind::kNone);

  std::vector<std::uint8_t> bin = ChartsheetPart(ReadFixture(kRedTwin + std::string(".xlsb")), false);
  const std::vector<std::uint8_t> red_bytes = {0x05, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF};
  const auto at = std::search(bin.begin(), bin.end(), red_bytes.begin(), red_bytes.end());
  ASSERT_NE(at, bin.end());
  const std::vector<std::uint8_t> automatic = {0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
  std::copy(automatic.begin(), automatic.end(), at);
  const std::vector<std::uint8_t> cleared_xlsb = ResavedWithTabColor(kRedTwin + std::string(".xlsb"), false, {});
  EXPECT_EQ(ChartsheetPart(cleared_xlsb, false), bin);
  EXPECT_EQ(LoadXlsb(cleared_xlsb).sheet(1).print_settings().tab_color.kind, ColorSpec::Kind::kNone);
}

TEST(ChartsheetViewRoundtrip, TabColorCarriesAcrossContainers) {
  // A colour set before a container change converts as Excel's twin with that colour converts.
  for (const bool from_xlsx : {true, false}) {
    const std::string suffix = from_xlsx ? ".xlsx" : ".xlsb";
    SCOPED_TRACE(suffix);
    Workbook edited = from_xlsx ? LoadXlsx(ReadFixture(kRedTwin + suffix)) : LoadXlsb(ReadFixture(kRedTwin + suffix));
    edited.sheet(1).mutable_print_settings().tab_color = ThemeColor(4U);
    const Workbook twin =
        from_xlsx ? LoadXlsx(ReadFixture(kThemeTwin + suffix)) : LoadXlsb(ReadFixture(kThemeTwin + suffix));
    const std::vector<std::uint8_t> converted = from_xlsx ? SaveXlsb(edited) : SaveXlsx(edited);
    EXPECT_EQ(ChartsheetPart(converted, !from_xlsx),
              ChartsheetPart(from_xlsx ? SaveXlsb(twin) : SaveXlsx(twin), !from_xlsx));
    const Workbook reread = from_xlsx ? LoadXlsb(converted) : LoadXlsx(converted);
    EXPECT_EQ(reread.sheet(1).print_settings().tab_color.kind, ColorSpec::Kind::kTheme);
    EXPECT_EQ(reread.sheet(1).print_settings().tab_color.theme, 4U);
  }
}

// The red twin's .xlsb with a code name "C" in its chartsheet's BrtCsProp, beside the tab colour.
std::vector<std::uint8_t> RedXlsbWithCodeName() {
  const std::vector<std::uint8_t> source = ReadFixture(kRedTwin + std::string(".xlsb"));
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(source))));
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const std::string& path : zip.list_entries()) {
    std::vector<std::uint8_t> bytes = zip.read_entry(path).value();
    if (path == "xl/chartsheets/sheet1.bin") {
      const std::vector<std::uint8_t> plain = {0x8B, 0x05, 0x0E, 0x01, 0x00, 0x05, 0xFF, 0x00, 0x00,
                                               0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00, 0x00};
      const std::vector<std::uint8_t> named = {0x8B, 0x05, 0x10, 0x01, 0x00, 0x05, 0xFF, 0x00, 0x00, 0xFF,
                                               0x00, 0x00, 0xFF, 0x01, 0x00, 0x00, 0x00, 0x43, 0x00};
      const auto at = std::search(bytes.begin(), bytes.end(), plain.begin(), plain.end());
      EXPECT_NE(at, bytes.end());
      const auto offset = at - bytes.begin();
      bytes.erase(at, at + static_cast<std::ptrdiff_t>(plain.size()));
      bytes.insert(bytes.begin() + offset, named.begin(), named.end());
    }
    EXPECT_NE(mz_zip_writer_add_mem(&writer, path.c_str(), bytes.data(), bytes.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
  }
  void* archive = nullptr;
  std::size_t size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + size);
  mz_free(archive);
  return out;
}

TEST(ChartsheetViewRoundtrip, XlsbTabColorBesideACodeNameIsReadAndRestated) {
  const std::vector<std::uint8_t> source = RedXlsbWithCodeName();
  Workbook wb = LoadXlsb(source);
  EXPECT_EQ(wb.sheet(1).print_settings().tab_color.kind, ColorSpec::Kind::kRgb);
  EXPECT_EQ(wb.sheet(1).print_settings().tab_color.rgb, 0xFFFF0000U);
  const std::vector<std::uint8_t> stored = ChartsheetPart(source, false);
  EXPECT_EQ(ChartsheetPart(SaveXlsb(wb), false), stored);

  // Clearing rewrites only the colour bytes; the code name stays.
  wb.sheet(1).mutable_print_settings().tab_color = ColorSpec{};
  const std::vector<std::uint8_t> cleared = SaveXlsb(wb);
  std::vector<std::uint8_t> expected = stored;
  const std::vector<std::uint8_t> red_bytes = {0x05, 0xFF, 0x00, 0x00, 0xFF, 0x00, 0x00, 0xFF};
  const auto at = std::search(expected.begin(), expected.end(), red_bytes.begin(), red_bytes.end());
  ASSERT_NE(at, expected.end());
  const std::vector<std::uint8_t> automatic = {0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
  std::copy(automatic.begin(), automatic.end(), at);
  EXPECT_EQ(ChartsheetPart(cleared, false), expected);
  EXPECT_EQ(LoadXlsb(cleared).sheet(1).print_settings().tab_color.kind, ColorSpec::Kind::kNone);

  // A colour carried to .xlsx is written beside the code name.
  wb.sheet(1).mutable_print_settings().tab_color = ThemeColor(4U);
  EXPECT_EQ(LoadXlsb(SaveXlsb(wb)).sheet(1).print_settings().tab_color.theme, 4U);
  const std::string xml = Text(ChartsheetPart(SaveXlsx(wb), true));
  EXPECT_NE(xml.find("<sheetPr codeName=\"C\"><tabColor theme=\"4\"/></sheetPr>"), std::string::npos) << xml;
}

TEST(ChartsheetViewRoundtrip, XmlTabColorRewriteKeepsTheRestOfThePart) {
  const std::string bare = "<x:chartsheet a=\"1\"><x:sheetViews/></x:chartsheet>";
  const std::string red =
      "<x:chartsheet a=\"1\"><x:sheetPr><x:tabColor rgb=\"FFFF0000\"/></x:sheetPr><x:sheetViews/>"
      "</x:chartsheet>";
  // A part without `<sheetPr>` gains one, in its root's namespace prefix, as the root's first child.
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_tab_color(bare, {}), bare);
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_tab_color(bare, RgbColor(0xFFFF0000U)), red);

  const std::string plain = "<chartsheet><sheetPr codeName=\"C\"/><sheetViews/></chartsheet>";
  ColorSpec tinted = ThemeColor(5U);
  tinted.tint = -0.25;
  const std::string with_theme = io::xlsb::chartsheet_xml_with_tab_color(plain, tinted);
  EXPECT_EQ(with_theme,
            "<chartsheet><sheetPr codeName=\"C\"><tabColor theme=\"5\" tint=\"-0.25\"/></sheetPr><sheetViews/>"
            "</chartsheet>");
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_tab_color(with_theme, tinted), with_theme);
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_tab_color(with_theme, {}), plain);
  EXPECT_EQ(io::xlsb::chartsheet_xml_with_tab_color("<chartsheet/>", RgbColor(0xFFFF0000U)), "<chartsheet/>");
}

// ---- Margins, page setup, header and footer, protection ----

constexpr const char* kMarginsXml =
    "<pageMargins left=\"0.5\" right=\"0.75\" top=\"1\" bottom=\"1.25\" header=\"0.25\" footer=\"0.375\"/>";
constexpr const char* kLandscapeXml = "<pageSetup paperSize=\"9\" orientation=\"landscape\"/>";
constexpr const char* kLegalXml = "<pageSetup paperSize=\"5\" orientation=\"portrait\"/>";
constexpr const char* kHeaderFooterXml =
    "<headerFooter><oddHeader>&amp;RR</oddHeader><oddFooter>&amp;LL&amp;P&amp;CC</oddFooter></headerFooter>";
constexpr const char* kDifferentFirstXml =
    "<headerFooter differentFirst=\"1\"><oddHeader>&amp;CO</oddHeader><firstHeader>&amp;CF</firstHeader>"
    "</headerFooter>";

ChartsheetProtection HashedProtection() {
  ChartsheetProtection protection;
  protection.algorithm_name = "SHA-512";
  protection.hash_value = "FEOoOgNU+lu/gQaYkcJsMsJCAdolOUZsju7crdLGfmo6yGCo62jHPQ2LfHdvg+umnjLNWsGZoZgG16ju/+taFQ==";
  protection.salt_value = "BAMwdck9CSCcwdaFLEANaA==";
  protection.spin_count = 100000U;
  protection.content_locked = true;
  protection.objects_locked = true;
  return protection;
}

std::string TwinPath(const char* name, const char* ext) {
  return std::string("chartsheet_twins/") + name + "." + ext;
}

Workbook LoadTwin(const char* name, const char* ext) {
  const std::vector<std::uint8_t> bytes = ReadFixture(TwinPath(name, ext));
  return std::string(ext) == "xlsx" ? LoadXlsx(bytes) : LoadXlsb(bytes);
}

std::vector<std::uint8_t> TwinPart(const char* name, const char* ext) {
  return ChartsheetPart(ReadFixture(TwinPath(name, ext)), std::string(ext) == "xlsx");
}

// The part's records, without the revision-id block Excel appends (BrtFRTBegin, the uid, BrtFRTEnd).
std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>> SettingRecords(const std::vector<std::uint8_t>& part) {
  std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>> out;
  io::ByteSpan cursor = SpanOf(part);
  while (cursor.size > 0) {
    const Expected<io::xlsb::XlsbRecord, Error> rec = io::xlsb::read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "truncated record";
      break;
    }
    const std::uint16_t type = rec.value().type;
    if (type != 37U && type != 3072U && type != 38U) {
      const io::ByteSpan payload = rec.value().payload;
      out.emplace_back(type, std::vector<std::uint8_t>(payload.data, payload.data + payload.size));
    }
  }
  return out;
}

// The body of a `<chartsheet>` part from `<sheetPr>` on, without the root's namespaces and revision id.
std::string XmlBody(const std::vector<std::uint8_t>& part) {
  const std::string xml = Text(part);
  const std::size_t at = xml.find("<sheetPr");
  return at == std::string::npos ? xml : xml.substr(at);
}

TEST(ChartsheetOwnSettings, ModelHoldsWhatTheFileStates) {
  for (const char* ext : {"xlsx", "xlsb"}) {
    SCOPED_TRACE(ext);
    const Workbook control = LoadTwin("cs_control", ext);
    EXPECT_FALSE(control.sheet(1).chartsheet_protection().has_value());
    EXPECT_TRUE(control.sheet(1).print_settings().page_setup_xml.empty());
    EXPECT_TRUE(control.sheet(1).print_settings().header_footer_xml.empty());

    const SheetPrintSettings margins = LoadTwin("cs_margins", ext).sheet(1).print_settings();
    EXPECT_EQ(margins.page_margins_xml, kMarginsXml);
    EXPECT_DOUBLE_EQ(margins.page_margins.left, 0.5);
    EXPECT_DOUBLE_EQ(margins.page_margins.right, 0.75);
    EXPECT_DOUBLE_EQ(margins.page_margins.top, 1.0);
    EXPECT_DOUBLE_EQ(margins.page_margins.bottom, 1.25);
    EXPECT_DOUBLE_EQ(margins.page_margins.header, 0.25);
    EXPECT_DOUBLE_EQ(margins.page_margins.footer, 0.375);

    EXPECT_EQ(LoadTwin("cs_page_setup_landscape_a4", ext).sheet(1).print_settings().page_setup_xml, kLandscapeXml);
    EXPECT_EQ(LoadTwin("cs_page_setup_legal", ext).sheet(1).print_settings().page_setup_xml, kLegalXml);
    EXPECT_EQ(LoadTwin("cs_header_footer", ext).sheet(1).print_settings().header_footer_xml, kHeaderFooterXml);
    EXPECT_EQ(LoadTwin("cs_header_footer_different_first", ext).sheet(1).print_settings().header_footer_xml,
              kDifferentFirstXml);

    EXPECT_EQ(LoadTwin("cs_protect_hash", ext).sheet(1).chartsheet_protection(), HashedProtection());
    ChartsheetProtection content_only;
    content_only.content_locked = true;
    EXPECT_EQ(LoadTwin("cs_protect_content_only", ext).sheet(1).chartsheet_protection(), content_only);
  }
}

// Loads `from` in each container, applies `edit` to the model and saves in both containers: each save holds what
// Excel's save of the `to` book holds.
void ExpectEditWritesTwin(const char* from, const char* to, const std::function<void(Workbook&)>& edit) {
  {
    Workbook wb = LoadTwin(from, "xlsx");
    edit(wb);
    EXPECT_EQ(Text(ChartsheetPart(SaveXlsx(wb), true)), Text(TwinPart(to, "xlsx"))) << "xlsx to xlsx";
    EXPECT_EQ(SettingRecords(ChartsheetPart(SaveXlsb(wb), false)), SettingRecords(TwinPart(to, "xlsb")))
        << "xlsx to xlsb";
  }
  {
    Workbook wb = LoadTwin(from, "xlsb");
    edit(wb);
    EXPECT_EQ(ChartsheetPart(SaveXlsb(wb), false), TwinPart(to, "xlsb")) << "xlsb to xlsb";
    EXPECT_EQ(XmlBody(ChartsheetPart(SaveXlsx(wb), true)), XmlBody(TwinPart(to, "xlsx"))) << "xlsb to xlsx";
  }
}

TEST(ChartsheetOwnSettings, MarginsSetWriteWhatExcelWrites) {
  ExpectEditWritesTwin("cs_control", "cs_margins",
                       [](Workbook& wb) { wb.sheet(1).mutable_print_settings().page_margins_xml = kMarginsXml; });
}

TEST(ChartsheetOwnSettings, PageSetupSetWritesWhatExcelWrites) {
  ExpectEditWritesTwin("cs_control", "cs_page_setup_landscape_a4",
                       [](Workbook& wb) { wb.sheet(1).mutable_print_settings().page_setup_xml = kLandscapeXml; });
  ExpectEditWritesTwin("cs_control", "cs_page_setup_legal",
                       [](Workbook& wb) { wb.sheet(1).mutable_print_settings().page_setup_xml = kLegalXml; });
  ExpectEditWritesTwin("cs_page_setup_landscape_a4", "cs_page_setup_legal",
                       [](Workbook& wb) { wb.sheet(1).mutable_print_settings().page_setup_xml = kLegalXml; });
  ExpectEditWritesTwin("cs_page_setup_legal", "cs_control",
                       [](Workbook& wb) { wb.sheet(1).mutable_print_settings().page_setup_xml.clear(); });
}

TEST(ChartsheetOwnSettings, HeaderFooterSetWritesWhatExcelWrites) {
  ExpectEditWritesTwin("cs_control", "cs_header_footer",
                       [](Workbook& wb) { wb.sheet(1).mutable_print_settings().header_footer_xml = kHeaderFooterXml; });
  ExpectEditWritesTwin("cs_header_footer", "cs_header_footer_different_first", [](Workbook& wb) {
    wb.sheet(1).mutable_print_settings().header_footer_xml = kDifferentFirstXml;
  });
  ExpectEditWritesTwin("cs_header_footer", "cs_control",
                       [](Workbook& wb) { wb.sheet(1).mutable_print_settings().header_footer_xml.clear(); });
}

TEST(ChartsheetOwnSettings, ProtectionSetWritesWhatExcelWrites) {
  ExpectEditWritesTwin("cs_control", "cs_protect_hash",
                       [](Workbook& wb) { wb.sheet(1).set_chartsheet_protection(HashedProtection()); });
  ExpectEditWritesTwin("cs_control", "cs_protect_content_only", [](Workbook& wb) {
    ChartsheetProtection protection;
    protection.content_locked = true;
    wb.sheet(1).set_chartsheet_protection(protection);
  });
  ExpectEditWritesTwin("cs_protect_hash", "cs_protect_content_only", [](Workbook& wb) {
    ChartsheetProtection protection;
    protection.content_locked = true;
    wb.sheet(1).set_chartsheet_protection(protection);
  });
  ExpectEditWritesTwin("cs_protect_hash", "cs_control",
                       [](Workbook& wb) { wb.sheet(1).set_chartsheet_protection(std::nullopt); });
}

TEST(ChartsheetOwnSettings, UnchangedChartsheetsResaveByteEqual) {
  for (const char* name :
       {"cs_control", "cs_margins", "cs_page_setup_landscape_a4", "cs_page_setup_legal", "cs_header_footer",
        "cs_header_footer_different_first", "cs_protect_hash", "cs_protect_content_only", "cs_zoom_80",
        "cs_page_setup_legal_unoriented", "cs_zoom_fit_explicit_off", "cs_page_setup_first_page",
        "cs_page_setup_orientation_only", "cs_code_name", "cs_unpublished", "cs_code_name_unpublished"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(ChartsheetPart(SaveXlsx(LoadTwin(name, "xlsx")), true), TwinPart(name, "xlsx"));
  }
  for (const char* name :
       {"cs_control", "cs_margins", "cs_page_setup_landscape_a4", "cs_page_setup_legal", "cs_header_footer",
        "cs_header_footer_different_first", "cs_protect_hash", "cs_protect_content_only", "cs_zoom_80",
        "cs_zoom_fit_explicit_off", "cs_page_setup_first_page", "cs_page_setup_orientation_only", "cs_code_name",
        "cs_unpublished", "cs_code_name_unpublished"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(ChartsheetPart(SaveXlsb(LoadTwin(name, "xlsb")), false), TwinPart(name, "xlsb"));
  }
}

// An .xlsx `<pageSetup>` always carries `orientation`: Excel adds the portrait default after `paperSize`.
TEST(ChartsheetOwnSettings, PageSetupWithoutOrientationIsWrittenPortrait) {
  const Workbook authored = LoadTwin("cs_page_setup_legal_unoriented", "xlsx");
  EXPECT_EQ(authored.sheet(1).print_settings().page_setup_xml, "<pageSetup paperSize=\"5\"/>");
  EXPECT_EQ(SettingRecords(ChartsheetPart(SaveXlsb(authored), false)),
            SettingRecords(TwinPart("cs_page_setup_legal", "xlsb")));
  const Workbook from_xlsb = LoadTwin("cs_page_setup_legal", "xlsb");
  EXPECT_EQ(XmlBody(ChartsheetPart(SaveXlsx(from_xlsb), true)), XmlBody(TwinPart("cs_page_setup_legal", "xlsx")));
}

// A .xlsb chartsheet always carries BrtCsProtection between BrtEndCsViews and BrtMargins, all zero when unprotected;
// a hash adds BrtCsProtectionIso before it with the same content and objects flags.
TEST(ChartsheetOwnSettings, XlsbProtectionRecordsFollowTheMeasuredOrder) {
  const auto between = [](const std::vector<std::uint8_t>& part) {
    std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>> run;
    bool inside = false;
    for (auto& record : SettingRecords(part)) {
      if (record.first == 476U) {
        inside = false;
      }
      if (inside) {
        run.push_back(record);
      }
      if (record.first == 140U) {
        inside = true;
      }
    }
    return run;
  };
  const auto unprotected = between(ChartsheetPart(SaveXlsb(LoadTwin("cs_control", "xlsx")), false));
  ASSERT_EQ(unprotected.size(), 1U);
  EXPECT_EQ(unprotected[0].first, 669U);
  EXPECT_EQ(unprotected[0].second, std::vector<std::uint8_t>(10U, 0U));

  const auto hashed = between(ChartsheetPart(SaveXlsb(LoadTwin("cs_protect_hash", "xlsx")), false));
  ASSERT_EQ(hashed.size(), 2U);
  EXPECT_EQ(hashed[0].first, 679U);
  EXPECT_EQ(hashed[1].first, 669U);
  EXPECT_EQ(hashed[1].second, (std::vector<std::uint8_t>{0, 0, 1, 0, 0, 0, 1, 0, 0, 0}));
}

// An explicit `zoomToFit="0"` is not written back once the view is restated, in either container.
TEST(ChartsheetOwnSettings, ExplicitZoomToFitOffIsNotWrittenBack) {
  const Workbook authored = LoadTwin("cs_zoom_fit_explicit_off", "xlsx");
  EXPECT_FALSE(authored.sheet(1).view().zoom_to_fit);
  EXPECT_EQ(SettingRecords(ChartsheetPart(SaveXlsb(authored), false)),
            SettingRecords(TwinPart("cs_zoom_fit_explicit_off", "xlsb")));
  const Workbook from_xlsb = LoadTwin("cs_zoom_fit_explicit_off", "xlsb");
  EXPECT_EQ(Text(ChartsheetPart(SaveXlsx(from_xlsb), true)).find("zoomToFit"), std::string::npos);

  Workbook zoomed = LoadTwin("cs_zoom_fit_explicit_off", "xlsx");
  SetZoom(zoomed, 1, 80U);
  EXPECT_EQ(Text(ChartsheetPart(SaveXlsx(zoomed), true)), Text(TwinPart("cs_zoom_80", "xlsx")));
}

// `orientation` goes in its schema position and no `paperSize` is added: after `firstPageNumber`, or first.
TEST(ChartsheetOwnSettings, PageSetupWithoutPaperSizeKeepsSchemaOrder) {
  constexpr const char* kFirstPageXml = "<pageSetup firstPageNumber=\"3\" useFirstPageNumber=\"1\"/>";
  constexpr const char* kOrientationOnlyXml = "<pageSetup orientation=\"landscape\"/>";
  ExpectEditWritesTwin("cs_control", "cs_page_setup_first_page",
                       [&](Workbook& wb) { wb.sheet(1).mutable_print_settings().page_setup_xml = kFirstPageXml; });
  ExpectEditWritesTwin("cs_control", "cs_page_setup_orientation_only", [&](Workbook& wb) {
    wb.sheet(1).mutable_print_settings().page_setup_xml = kOrientationOnlyXml;
  });
  ExpectEditWritesTwin("cs_page_setup_first_page", "cs_page_setup_orientation_only", [&](Workbook& wb) {
    wb.sheet(1).mutable_print_settings().page_setup_xml = kOrientationOnlyXml;
  });
  ExpectEditWritesTwin("cs_page_setup_orientation_only", "cs_page_setup_first_page",
                       [&](Workbook& wb) { wb.sheet(1).mutable_print_settings().page_setup_xml = kFirstPageXml; });

  // Excel's own .xlsx converts to Excel's own .xlsb.
  for (const char* name : {"cs_page_setup_first_page", "cs_page_setup_orientation_only"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(SettingRecords(ChartsheetPart(SaveXlsb(LoadTwin(name, "xlsx")), false)),
              SettingRecords(TwinPart(name, "xlsb")));
  }
}

TEST(ChartsheetOwnSettings, SheetPrAttributesAreReadFromThePart) {
  const std::pair<const char*, const char*> twins[] = {
      {"cs_code_name", "<sheetPr codeName=\"ChartA\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"},
      {"cs_unpublished", "<sheetPr published=\"0\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"},
      {"cs_code_name_unpublished",
       "<sheetPr published=\"0\" codeName=\"ChartA\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"},
  };
  for (const char* ext : {"xlsx", "xlsb"}) {
    for (const auto& [name, xml] : twins) {
      SCOPED_TRACE(std::string(name) + "." + ext);
      EXPECT_EQ(LoadTwin(name, ext).sheet(1).print_settings().sheet_pr_xml, xml);
    }
  }
}

// `published` and `codeName` of a chartsheet's `<sheetPr>` are restated, `published` first, the tab colour kept.
TEST(ChartsheetOwnSettings, SheetPrSetWritesWhatExcelWrites) {
  const auto set_pr = [](const char* xml) {
    return [xml](Workbook& wb) { wb.sheet(1).mutable_print_settings().sheet_pr_xml = xml; };
  };
  ExpectEditWritesTwin("cs_control", "cs_code_name",
                       set_pr("<sheetPr codeName=\"ChartA\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"));
  ExpectEditWritesTwin("cs_control", "cs_unpublished",
                       set_pr("<sheetPr published=\"0\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"));
  ExpectEditWritesTwin("cs_control", "cs_code_name_unpublished",
                       set_pr("<sheetPr codeName=\"ChartA\" published=\"0\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"));
  ExpectEditWritesTwin("cs_code_name", "cs_code_name_unpublished",
                       set_pr("<sheetPr published=\"0\" codeName=\"ChartA\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"));
  ExpectEditWritesTwin("cs_code_name_unpublished", "cs_unpublished",
                       set_pr("<sheetPr published=\"0\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"));
  ExpectEditWritesTwin("cs_code_name_unpublished", "cs_control",
                       set_pr("<sheetPr><tabColor rgb=\"FFFF0000\"/></sheetPr>"));
}

}  // namespace
}  // namespace formulon
