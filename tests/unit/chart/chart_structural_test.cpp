#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "drawing/drawing_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "passthrough_part.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kChartRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
constexpr const char* kStyleRel = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
constexpr const char* kColorsRel = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
constexpr const char* kCtChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
constexpr const char* kCtRels = "application/vnd.openxmlformats-package.relationships+xml";
constexpr const char* kChartData =
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart "
    "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{rId}\"/>"
    "</a:graphicData>";
constexpr const char* kChartPath = "xl/charts/chart1.xml";
constexpr const char* kStylePath = "xl/charts/style1.xml";
constexpr const char* kColorsPath = "xl/charts/colors1.xml";
constexpr const char* kChartRelsPath = "xl/charts/_rels/chart1.xml.rels";

std::string ChartXml(const std::string& cat, const std::string& val) {
  return "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea>"
         "<c:barChart><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:cat><c:strRef><c:f>" +
         cat +
         "</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>a</c:v></c:pt></c:strCache></c:strRef></c:cat>"
         "<c:val><c:numRef><c:f>" +
         val +
         "</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>4</c:v></c:pt>"
         "</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";
}

Expected<void, Error> AddPart(Workbook& wb, const char* path, const char* content_type, const std::string& text) {
  return wb.add_passthrough_part(PassthroughPart(path, content_type, Bytes(text.begin(), text.end())));
}

/// Sheet1 (index 0) hosts one chart with style/colors parts; "Other" is index 1.
Workbook WithChart(const std::string& chart_xml) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Other");
  ImageInsertOptions placement;
  placement.width_emu = 4572000;
  placement.height_emu = 2743200;
  EXPECT_TRUE(static_cast<bool>(insert_graphic_frame(wb, 0, placement, kChartData, kChartPath, kChartRel)));
  EXPECT_TRUE(static_cast<bool>(AddPart(wb, kChartPath, kCtChart, chart_xml)));
  EXPECT_TRUE(
      static_cast<bool>(AddPart(wb, kStylePath, "application/vnd.ms-office.chartstyle+xml", "<cs:chartStyle/>")));
  EXPECT_TRUE(
      static_cast<bool>(AddPart(wb, kColorsPath, "application/vnd.ms-office.chartcolorstyle+xml", "<cs:colorStyle/>")));
  const std::string rels = std::string(
                               "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                               "<Relationship Id=\"rId1\" Type=\"") +
                           kStyleRel + "\" Target=\"style1.xml\"/><Relationship Id=\"rId2\" Type=\"" + kColorsRel +
                           "\" Target=\"colors1.xml\"/></Relationships>";
  EXPECT_TRUE(static_cast<bool>(AddPart(wb, kChartRelsPath, kCtRels, rels)));
  return wb;
}

bool HasPart(const Workbook& wb, const std::string& path) {
  return io::ooxml::find_passthrough_part(wb, path) != nullptr;
}

std::string PartText(const Workbook& wb, const std::string& path = kChartPath) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  EXPECT_NE(part, nullptr);
  return part == nullptr ? std::string() : std::string(part->bytes.begin(), part->bytes.end());
}

bool Has(const std::string& text, const std::string& needle) {
  return text.find(needle) != std::string::npos;
}

Bytes ReadFile(const std::string& name) {
  Bytes out;
  std::FILE* file = std::fopen((std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name).c_str(), "rb");
  if (file == nullptr) {
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  out.resize(static_cast<std::size_t>(std::ftell(file)));
  std::fseek(file, 0, SEEK_SET);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    out.clear();
  }
  std::fclose(file);
  return out;
}

Workbook LoadFixture(const std::string& name) {
  const Bytes bytes = ReadFile(name);
  EXPECT_FALSE(bytes.empty()) << name;
  const io::ByteSpan span{bytes.data(), bytes.size()};
  if (name.size() > 5 && name.substr(name.size() - 5) == ".xlsb") {
    auto read = io::xlsb::read_xlsb(span);
    EXPECT_TRUE(static_cast<bool>(read));
    return read ? std::move(read.value().workbook) : Workbook::create_empty();
  }
  auto read = io::read_ooxml(span);
  EXPECT_TRUE(static_cast<bool>(read));
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

Workbook Reload(const Workbook& wb) {
  auto saved = wb.save();
  EXPECT_TRUE(static_cast<bool>(saved));
  if (!saved) {
    return Workbook::create_empty();
  }
  const io::ByteSpan span{saved.value().data(), saved.value().size()};
  auto read = io::read_ooxml(span);
  EXPECT_TRUE(static_cast<bool>(read));
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

}  // namespace

TEST(ChartStructural, RowInsertAndDeleteMoveRefs) {
  Workbook wb = WithChart(ChartXml("Sheet1!$A$2:$A$5", "Sheet1!$B$2:$B$5"));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 1)));
  EXPECT_TRUE(Has(PartText(wb), "<c:f>Sheet1!$A$3:$A$6</c:f>"));
  EXPECT_TRUE(Has(PartText(wb), "<c:f>Sheet1!$B$3:$B$6</c:f>"));
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 0, 1)));
  EXPECT_TRUE(Has(PartText(wb), "<c:f>Sheet1!$A$2:$A$5</c:f>"));
}

TEST(ChartStructural, ColumnInsertMovesRefs) {
  Workbook wb = WithChart(ChartXml("Sheet1!$A$2:$A$5", "Sheet1!$B$2:$B$5"));
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0, 1)));
  EXPECT_TRUE(Has(PartText(wb), "<c:f>Sheet1!$B$2:$B$5</c:f>"));
  EXPECT_TRUE(Has(PartText(wb), "<c:f>Sheet1!$C$2:$C$5</c:f>"));
}

TEST(ChartStructural, EditOnAnUnrelatedSheetLeavesPartsIdentical) {
  Workbook wb = WithChart(ChartXml("Sheet1!$A$2:$A$5", "Sheet1!$B$2:$B$5"));
  const std::string before = PartText(wb);
  const std::string rels_before = PartText(wb, kChartRelsPath);
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(1, 0, 3)));
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(1, 0, 1)));
  EXPECT_EQ(PartText(wb), before);
  EXPECT_EQ(PartText(wb, kChartRelsPath), rels_before);
}

TEST(ChartStructural, RenameUpdatesRefs) {
  Workbook wb = WithChart(ChartXml("Sheet1!$A$2:$A$5", "Sheet1!$B$2:$B$5"));
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "My Data")));
  const std::string after = PartText(wb);
  EXPECT_TRUE(Has(after, "<c:f>'My Data'!$A$2:$A$5</c:f>"));
  EXPECT_FALSE(Has(after, "Sheet1"));
}

TEST(ChartStructural, AlternatePrefixRefsFollowWorkbookEdits) {
  const std::string chart =
      "<q:chartSpace xmlns:q=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><q:chart><q:plotArea>"
      "<q:barChart><q:ser><q:idx val=\"0\"/><q:order val=\"0\"/><q:cat><q:strRef><q:f>Sheet1!$A$2:$A$5</q:f>"
      "</q:strRef></q:cat><q:val><q:numRef><q:f>Sheet1!$B$2:$B$5</q:f></q:numRef></q:val></q:ser></q:barChart>"
      "</q:plotArea></q:chart></q:chartSpace>";
  Workbook wb = WithChart(chart);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 1)));
  EXPECT_TRUE(Has(PartText(wb), "<q:f>Sheet1!$A$3:$A$6</q:f>"));
  EXPECT_TRUE(Has(PartText(wb), "<q:f>Sheet1!$B$3:$B$6</q:f>"));
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "Data")));
  EXPECT_TRUE(Has(PartText(wb), "<q:f>Data!$A$3:$A$6</q:f>"));
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 0, 1)));
  EXPECT_TRUE(Has(PartText(wb), "<q:f>Data!$A$2:$A$5</q:f>"));
  EXPECT_TRUE(Has(PartText(wb), "<q:f>Data!$B$2:$B$5</q:f>"));
}

TEST(ChartStructural, RemovingTheReferencedSheetYieldsRefSites) {
  Workbook wb = WithChart(ChartXml("Other!$A$2:$A$5", "Other!$B$2:$B$5"));
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));
  const std::string after = PartText(wb);
  EXPECT_TRUE(Has(after, "<c:f>#REF!</c:f>"));
  EXPECT_FALSE(Has(after, "Other"));
  EXPECT_FALSE(Has(after, "<c:strCache>"));
  EXPECT_TRUE(Has(after, "<c:formatCode>General</c:formatCode>"));
  EXPECT_TRUE(HasPart(wb, kChartPath));
}

TEST(ChartStructural, RemovingTheHostSheetDropsItsChartParts) {
  Workbook wb = WithChart(ChartXml("Other!$A$2:$A$5", "Other!$B$2:$B$5"));
  ASSERT_TRUE(HasPart(wb, kChartPath));
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));
  EXPECT_FALSE(HasPart(wb, kChartPath));
  EXPECT_FALSE(HasPart(wb, kChartRelsPath));
  EXPECT_FALSE(HasPart(wb, kStylePath));
  EXPECT_FALSE(HasPart(wb, kColorsPath));

  const Workbook reloaded = Reload(wb);
  EXPECT_EQ(reloaded.sheet_count(), 1U);
  EXPECT_FALSE(HasPart(reloaded, kChartPath));
  EXPECT_FALSE(HasPart(reloaded, kStylePath));
}

TEST(ChartStructural, RemovingAnotherSheetKeepsTheChartParts) {
  Workbook wb = WithChart(ChartXml("Sheet1!$A$2:$A$5", "Sheet1!$B$2:$B$5"));
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));
  EXPECT_TRUE(HasPart(wb, kChartPath));
  EXPECT_TRUE(HasPart(wb, kChartRelsPath));
  EXPECT_TRUE(HasPart(wb, kStylePath));
  EXPECT_TRUE(HasPart(wb, kColorsPath));
}

TEST(ChartStructural, FixtureRefsFollowEditAndSurviveReload) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::string before = PartText(wb, "xl/charts/chart1.xml");
  ASSERT_TRUE(Has(before, "<c:f>Sheet1!$B$2:$B$5</c:f>"));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));
  EXPECT_TRUE(Has(PartText(wb, "xl/charts/chart1.xml"), "<c:f>Sheet1!$B$4:$B$7</c:f>"));

  const Workbook reloaded = Reload(wb);
  EXPECT_TRUE(Has(PartText(reloaded, "xl/charts/chart1.xml"), "<c:f>Sheet1!$B$4:$B$7</c:f>"));
}

TEST(ChartStructural, XlsbLoadedChartRefsFollow) {
  Workbook wb = LoadFixture("chart_xlsb.xlsb");
  ASSERT_TRUE(Has(PartText(wb, "xl/charts/chart1.xml"), "<c:f>Sheet1!$B$2:$B$5</c:f>"));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 1)));
  EXPECT_TRUE(Has(PartText(wb, "xl/charts/chart1.xml"), "<c:f>Sheet1!$B$3:$B$6</c:f>"));
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "Data")));
  EXPECT_TRUE(Has(PartText(wb, "xl/charts/chart1.xml"), "<c:f>Data!$B$3:$B$6</c:f>"));
}

}  // namespace formulon
