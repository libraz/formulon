#include "chart/chart_ref.h"

#include <cstdint>
#include <sstream>
#include <string>
#include <string_view>
#include <vector>

#include "chart/chart_ns.h"
#include "drawing/drawing_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/xml_utils.h"
#include "parser/ref_transforms.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kChartRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
constexpr const char* kCtChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
constexpr const char* kChartData =
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart "
    "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{rId}\"/>"
    "</a:graphicData>";
constexpr const char* kChartPath = "xl/charts/chart1.xml";

constexpr const char* kChartOpen =
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
    "xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\"><c:chart><c:plotArea><c:barChart>";
constexpr const char* kChartClose = "</c:barChart></c:plotArea></c:chart></c:chartSpace>";

std::string StrRef(const std::string& f) {
  return "<c:strRef><c:f>" + f +
         "</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>old</c:v></c:pt>"
         "</c:strCache></c:strRef>";
}

std::string NumRef(const std::string& f) {
  return "<c:numRef><c:f>" + f +
         "</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val=\"2\"/>"
         "<c:pt idx=\"0\"><c:v>4</c:v></c:pt><c:pt idx=\"1\"><c:v>5</c:v></c:pt></c:numCache></c:numRef>";
}

std::string Series(const std::string& name, const std::string& cat, const std::string& val) {
  return "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx>" + StrRef(name) + "</c:tx><c:cat>" + StrRef(cat) +
         "</c:cat><c:val>" + NumRef(val) + "</c:val></c:ser>";
}

std::string MultiLevelSeries(const std::string& cat) {
  return "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx>" + StrRef("Sheet1!$C$1") +
         "</c:tx><c:cat><c:multiLvlStrRef><c:f>" + cat +
         "</c:f><c:multiLvlStrCache><c:ptCount val=\"2\"/><c:lvl><c:pt idx=\"0\"><c:v>a</c:v></c:pt>"
         "<c:pt idx=\"1\"><c:v>b</c:v></c:pt></c:lvl><c:lvl><c:pt idx=\"0\"><c:v>g1</c:v></c:pt></c:lvl>"
         "</c:multiLvlStrCache></c:multiLvlStrRef></c:cat><c:val>" +
         NumRef("Sheet1!$C$2:$C$3") + "</c:val></c:ser>";
}

pugi::xml_document Parse(const std::string& xml) {
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_string(xml.c_str()));
  return doc;
}

std::string_view NamespaceUri(const pugi::xml_node& node) {
  const std::string_view name = node.name();
  const std::size_t colon = name.find(':');
  const std::string_view prefix = colon == std::string_view::npos ? std::string_view() : name.substr(0, colon);
  return io::find_namespace_binding(node, prefix);
}

pugi::xml_node FirstElementChild(const pugi::xml_node& parent) {
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element) {
      return child;
    }
  }
  return {};
}

std::vector<std::string> Texts(const pugi::xml_document& doc) {
  std::vector<std::string> out;
  for (const pugi::xml_node& site : chart_ref::enumerate_sites(doc)) {
    out.emplace_back(site.child_value());
  }
  return out;
}

/// A workbook whose Sheet1 holds one chart frame showing `chart_xml`.
Workbook WithChart(const std::string& chart_xml) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Other");
  ImageInsertOptions placement;
  placement.width_emu = 4572000;
  placement.height_emu = 2743200;
  EXPECT_TRUE(static_cast<bool>(insert_graphic_frame(wb, 0, placement, kChartData, kChartPath, kChartRel)));
  EXPECT_TRUE(static_cast<bool>(
      wb.add_passthrough_part(PassthroughPart(kChartPath, kCtChart, Bytes(chart_xml.begin(), chart_xml.end())))));
  return wb;
}

std::string PartText(const Workbook& wb, const std::string& path = kChartPath) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  EXPECT_NE(part, nullptr);
  return part == nullptr ? std::string() : std::string(part->bytes.begin(), part->bytes.end());
}

parser::RowColShiftTransform InsertRows(std::uint32_t index, std::uint32_t count) {
  return parser::RowColShiftTransform("Sheet1", parser::RowColAxis::kRow, parser::RowColEdit::kInsert, index, count);
}

}  // namespace

TEST(ChartRef, EnumeratesEverySiteKind) {
  const std::string xml =
      std::string(kChartOpen) + "<c:ser><c:tx>" + StrRef("S!$B$1") + "</c:tx><c:cat>" + StrRef("S!$A$2:$A$3") +
      "</c:cat><c:val>" + NumRef("S!$B$2:$B$3") + "</c:val><c:xVal>" + NumRef("S!$C$2") + "</c:xVal><c:yVal>" +
      NumRef("S!$D$2") + "</c:yVal><c:bubbleSize>" + NumRef("S!$E$2") +
      "</c:bubbleSize></c:ser>"
      "<c:title><c:tx>" +
      StrRef("S!$F$1") +
      "</c:tx></c:title>"
      "<c:ext><c15:f>S!$G$1</c15:f><c15:fullRef><c15:sqref>S!$H$1:$H$4</c15:sqref></c15:fullRef>"
      "<c15:formulaRef><c15:sqref>S!$I$1</c15:sqref></c15:formulaRef><c15:sqref>S!$J$1</c15:sqref></c:ext>" +
      kChartClose;
  const std::vector<std::string> expected = {"S!$B$1", "S!$A$2:$A$3", "S!$B$2:$B$3", "S!$C$2",      "S!$D$2",
                                             "S!$E$2", "S!$F$1",      "S!$G$1",      "S!$H$1:$H$4", "S!$I$1"};
  EXPECT_EQ(Texts(Parse(xml)), expected);
}

TEST(ChartRef, AlternateAndDefaultNamespacePrefixesFollowStructuralEdits) {
  const std::string xml = std::string("<root xmlns:q=\"") + chart_ns::kNsC + "\" xmlns:x=\"" + chart_ns::kNsC15 +
                          "\" xmlns:z=\"http://schemas.microsoft.com/office/drawing/2014/chartex\">" +
                          "<q:chartSpace><q:chart><q:f>Sheet1!$A$1</q:f><q:ext><x:f>Sheet1!$B$1</x:f>"
                          "<z:f>Sheet1!$F$1</z:f><x:fullRef>"
                          "<x:sqref>Sheet1!$C$1:$C$2</x:sqref></x:fullRef></q:ext></q:chart></q:chartSpace>"
                          "<chartSpace xmlns=\"" +
                          chart_ns::kNsC +
                          "\"><chart><f>Sheet1!$D$1</f><ext><x:formulaRef>"
                          "<x:sqref>Sheet1!$E$1</x:sqref></x:formulaRef></ext></chart></chartSpace></root>";
  pugi::xml_document doc = Parse(xml);
  const std::vector<std::string> initial = {"Sheet1!$A$1",      "Sheet1!$B$1", "Sheet1!$F$1",
                                            "Sheet1!$C$1:$C$2", "Sheet1!$D$1", "Sheet1!$E$1"};
  EXPECT_EQ(Texts(doc), initial);

  EXPECT_EQ(chart_ref::rewrite_sites(doc, InsertRows(0, 2)), 6U);
  const parser::SheetRenameTransform rename("Sheet1", "Data Set");
  EXPECT_EQ(chart_ref::rewrite_sites(doc, rename), 6U);
  const std::vector<std::string> expected = {"'Data Set'!$A$3",      "'Data Set'!$B$3", "'Data Set'!$F$3",
                                             "'Data Set'!$C$3:$C$4", "'Data Set'!$D$3", "'Data Set'!$E$3"};
  EXPECT_EQ(Texts(doc), expected);
}

TEST(ChartRef, AlternatePrefixKeepsCacheAndFilteredTitleNamespacesValid) {
  const std::string xml =
      std::string("<q:chartSpace xmlns:q=\"") + chart_ns::kNsC + "\" xmlns:x=\"" + chart_ns::kNsC15 + "\">" +
      "<q:chart><q:plotArea><q:barChart><q:ser xmlns:c15=\"urn:shadowed\"><q:idx val=\"0\"/><q:order val=\"0\"/>"
      "<q:tx><q:strRef><q:f>Other!$B$1</q:f><c15:foreign><c15:leaf/></c15:foreign>"
      "<q:strCache><q:ptCount val=\"1\"/><q:pt idx=\"0\"><q:v>name</q:v>"
      "</q:pt></q:strCache></q:strRef></q:tx><q:cat><q:strRef><q:f>Other!$A$2:$A$4</q:f></q:strRef></q:cat>"
      "<q:val><q:numRef><q:f>Other!$B$2:$B$4</q:f><q:numCache><q:formatCode>0.0</q:formatCode>"
      "<q:ptCount "
      "val=\"1\"/></q:numCache></q:numRef></q:val></q:ser></q:barChart></q:plotArea></q:chart></q:chartSpace>";
  pugi::xml_document doc = Parse(xml);
  const std::vector<std::string_view> order = {"Sheet1", "Other"};
  const parser::SheetRemovalTransform removal(order, 1);

  EXPECT_EQ(chart_ref::rewrite_sites(doc, removal), 3U);
  const std::vector<std::string> expected = {"#REF!", "#REF!", "#REF!"};
  EXPECT_EQ(Texts(doc), expected);

  std::ostringstream serialized;
  doc.document_element().print(serialized, "", pugi::format_raw);
  const std::string result = serialized.str();
  EXPECT_NE(result.find("<q:extLst>"), std::string::npos) << result;
  EXPECT_EQ(result.find("<c:extLst>"), std::string::npos) << result;
  EXPECT_NE(result.find("<c15ext:filteredSeriesTitle>"), std::string::npos) << result;
  EXPECT_NE(result.find("<c15ext:filteredCategoryTitle>"), std::string::npos) << result;
  EXPECT_NE(result.find("<c15ext:tx><q:strRef>"), std::string::npos) << result;
  EXPECT_NE(result.find("<c15ext:cat><q:strRef>"), std::string::npos) << result;
  EXPECT_NE(result.find("xmlns:c15ext=\"http://schemas.microsoft.com/office/drawing/2012/chart\""), std::string::npos)
      << result;
  EXPECT_NE(result.find("<c15ext:formulaRef>"), std::string::npos) << result;
  EXPECT_NE(result.find("<c15:foreign><c15:leaf/></c15:foreign>"), std::string::npos) << result;
  EXPECT_NE(result.find("<q:numCache><q:formatCode>General</q:formatCode>"), std::string::npos) << result;
  EXPECT_NE(result.find("xmlns:c15=\"urn:shadowed\""), std::string::npos) << result;

  pugi::xml_document reparsed;
  EXPECT_TRUE(reparsed.load_string(result.c_str())) << result;
  const pugi::xml_node series = reparsed.select_node("//q:ser").node();
  const pugi::xml_node ext_list = series.child("q:extLst");
  const pugi::xml_node title_ext = FirstElementChild(ext_list);
  const pugi::xml_node title = FirstElementChild(title_ext);
  const pugi::xml_node title_ref = FirstElementChild(FirstElementChild(title));
  EXPECT_EQ(NamespaceUri(title_ext), chart_ns::kNsC);
  EXPECT_EQ(NamespaceUri(title), chart_ns::kNsC15);
  EXPECT_EQ(NamespaceUri(title_ref), chart_ns::kNsC);
  const pugi::xml_node foreign = title_ref.child("c15:foreign");
  ASSERT_TRUE(foreign);
  EXPECT_EQ(NamespaceUri(foreign), "urn:shadowed");
  EXPECT_EQ(NamespaceUri(foreign.child("c15:leaf")), "urn:shadowed");
  const pugi::xml_node formula_ext = FirstElementChild(title_ref.child("q:extLst"));
  const pugi::xml_node formula_ref = FirstElementChild(formula_ext);
  EXPECT_EQ(NamespaceUri(formula_ext), chart_ns::kNsC);
  EXPECT_EQ(NamespaceUri(formula_ref), chart_ns::kNsC15);
}

TEST(ChartRef, ChartPrefixNamedC15DoesNotGetReboundByFilteredTitle) {
  const std::string xml = std::string("<c15:chartSpace xmlns:c15=\"") + chart_ns::kNsC + "\" xmlns:x=\"" +
                          chart_ns::kNsC15 + "\">" +
                          "<c15:chart><c15:plotArea><c15:barChart><c15:ser><c15:idx val=\"0\"/><c15:order val=\"0\"/>"
                          "<c15:tx><c15:strRef><c15:f>Other!$B$1</c15:f></c15:strRef></c15:tx>"
                          "<c15:cat><c15:strRef><c15:f>Other!$A$2:$A$4</c15:f></c15:strRef></c15:cat>"
                          "<c15:val><c15:numRef><c15:f>Other!$B$2:$B$4</c15:f></c15:numRef></c15:val>"
                          "</c15:ser></c15:barChart></c15:plotArea></c15:chart></c15:chartSpace>";
  pugi::xml_document doc = Parse(xml);
  const std::vector<std::string_view> order = {"Sheet1", "Other"};
  const parser::SheetRemovalTransform removal(order, 1);

  EXPECT_EQ(chart_ref::rewrite_sites(doc, removal), 3U);
  std::ostringstream serialized;
  doc.document_element().print(serialized, "", pugi::format_raw);
  const std::string result = serialized.str();
  EXPECT_NE(result.find("<c15:ext uri=\"{02D57815-91ED-43cb-92C2-25804820EDAC}\" xmlns:c15ext=\""), std::string::npos)
      << result;
  EXPECT_NE(result.find("<c15ext:filteredSeriesTitle>"), std::string::npos) << result;
  EXPECT_NE(result.find("<c15ext:formulaRef>"), std::string::npos) << result;

  pugi::xml_document reparsed;
  ASSERT_TRUE(reparsed.load_string(result.c_str())) << result;
  const pugi::xml_node series = reparsed.select_node("//c15:ser").node();
  const pugi::xml_node title_ext = FirstElementChild(series.child("c15:extLst"));
  const pugi::xml_node title = FirstElementChild(title_ext);
  const pugi::xml_node title_holder = FirstElementChild(title);
  const pugi::xml_node title_ref = FirstElementChild(title_holder);
  EXPECT_EQ(NamespaceUri(title_ext), chart_ns::kNsC);
  EXPECT_EQ(NamespaceUri(title), chart_ns::kNsC15);
  EXPECT_EQ(NamespaceUri(title_holder), chart_ns::kNsC15);
  EXPECT_EQ(NamespaceUri(title_ref), chart_ns::kNsC);
  const pugi::xml_node formula_ext = FirstElementChild(title_ref.child("c15:extLst"));
  const pugi::xml_node formula_ref = FirstElementChild(formula_ext);
  EXPECT_EQ(NamespaceUri(formula_ext), chart_ns::kNsC);
  EXPECT_EQ(NamespaceUri(formula_ref), chart_ns::kNsC15);
}

TEST(ChartRef, ExtListDefaultNamespaceDoesNotChangeMovedReference) {
  const std::string xml = std::string("<q:chartSpace xmlns:q=\"") + chart_ns::kNsC +
                          "\"><q:chart><q:plotArea>"
                          "<q:barChart><q:ser xmlns=\"" +
                          chart_ns::kNsC +
                          "\"><q:tx><strRef><f>Other!$B$1</f></strRef></q:tx>"
                          "<q:val><q:numRef><q:f>Other!$B$2</q:f></q:numRef></q:val>"
                          "<q:extLst xmlns=\"urn:extension-default\" xmlns:aux=\"urn:ext-only\"/>"
                          "</q:ser></q:barChart></q:plotArea></q:chart></q:chartSpace>";
  pugi::xml_document doc = Parse(xml);
  const std::vector<std::string_view> order = {"Sheet1", "Other"};
  const parser::SheetRemovalTransform removal(order, 1);

  EXPECT_EQ(chart_ref::rewrite_sites(doc, removal), 2U);
  std::ostringstream serialized;
  doc.document_element().print(serialized, "", pugi::format_raw);
  const std::string result = serialized.str();
  EXPECT_EQ(result.find("xmlns:aux=\"\""), std::string::npos) << result;

  pugi::xml_document reparsed;
  ASSERT_TRUE(reparsed.load_string(result.c_str())) << result;
  const pugi::xml_node series = reparsed.select_node("//q:ser").node();
  const pugi::xml_node title_ext = FirstElementChild(series.child("q:extLst"));
  const pugi::xml_node title = FirstElementChild(title_ext);
  const pugi::xml_node title_ref = FirstElementChild(FirstElementChild(title));
  EXPECT_EQ(NamespaceUri(title_ref), chart_ns::kNsC);
  const pugi::xml_node cache = title_ref.child("strCache");
  ASSERT_TRUE(cache);
  EXPECT_EQ(NamespaceUri(cache), chart_ns::kNsC);
  EXPECT_EQ(NamespaceUri(cache.child("ptCount")), chart_ns::kNsC);
  const pugi::xml_node formula_ext = FirstElementChild(title_ref.child("extLst"));
  EXPECT_EQ(NamespaceUri(formula_ext), chart_ns::kNsC);
  EXPECT_EQ(NamespaceUri(FirstElementChild(formula_ext)), chart_ns::kNsC15);
}

TEST(ChartRef, DefaultNamespaceCategoryFlattensAfterLevelDeletion) {
  const std::string xml =
      std::string("<chartSpace xmlns=\"") + chart_ns::kNsC +
      "\"><chart><plotArea><barChart><ser>"
      "<tx><strRef><f>Sheet1!$C$1</f></strRef></tx><cat><multiLvlStrRef><f>Sheet1!$A$2:$B$3</f>"
      "<multiLvlStrCache><ptCount val=\"2\"/><lvl/><lvl/></multiLvlStrCache></multiLvlStrRef></cat>"
      "<val><numRef><f>Sheet1!$C$2:$C$3</f></numRef></val></ser></barChart></plotArea></chart></chartSpace>";
  pugi::xml_document doc = Parse(xml);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kCol, parser::RowColEdit::kDelete, 0, 1);

  EXPECT_EQ(chart_ref::rewrite_sites(doc, del), 3U);
  const std::vector<std::string> expected = {"Sheet1!$B$1", "Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3"};
  EXPECT_EQ(Texts(doc), expected);
  std::ostringstream serialized;
  doc.document_element().print(serialized, "", pugi::format_raw);
  EXPECT_NE(serialized.str().find("<cat><strRef><f>Sheet1!$A$2:$A$3</f>"), std::string::npos) << serialized.str();
  EXPECT_EQ(serialized.str().find("multiLvlStr"), std::string::npos) << serialized.str();
}

TEST(ChartRef, ForeignAndShadowedFormulaElementsAreIgnored) {
  const std::string xml =
      std::string("<q:chartSpace xmlns:q=\"") + chart_ns::kNsC + "\" xmlns:c=\"" + chart_ns::kNsC + "\">" +
      "<q:chart><q:plotArea><q:barChart><q:ser><q:cat><q:strRef><q:f>Sheet1!$A$1</q:f></q:strRef></q:cat>"
      "<q:extLst><q:ext xmlns:c=\"urn:foreign\"><c:f>Sheet1!$B$1</c:f><q:f xmlns:q=\"urn:foreign\">"
      "Sheet1!$C$1</q:f></q:ext></q:extLst></q:ser></q:barChart></q:plotArea></q:chart></q:chartSpace>";
  pugi::xml_document doc = Parse(xml);
  EXPECT_EQ(Texts(doc), (std::vector<std::string>{"Sheet1!$A$1"}));
  EXPECT_EQ(chart_ref::rewrite_sites(doc, InsertRows(0, 1)), 1U);
  EXPECT_EQ(Texts(doc), (std::vector<std::string>{"Sheet1!$A$2"}));

  std::ostringstream serialized;
  doc.document_element().print(serialized, "", pugi::format_raw);
  EXPECT_NE(serialized.str().find("<c:f>Sheet1!$B$1</c:f>"), std::string::npos) << serialized.str();
  EXPECT_NE(serialized.str().find("Sheet1!$C$1</q:f>"), std::string::npos) << serialized.str();
}

TEST(ChartRef, ForeignCacheWithShadowedPrefixSurvivesReset) {
  const std::string xml =
      std::string("<q:chartSpace xmlns:q=\"") + chart_ns::kNsC +
      "\"><q:chart><q:plotArea><q:barChart><q:ser>"
      "<q:val><q:numRef><q:f>Other!$B$1</q:f><q:numCache xmlns:q=\"urn:foreign\"><q:formatCode>keep</q:formatCode>"
      "</q:numCache></q:numRef></q:val></q:ser></q:barChart></q:plotArea></q:chart></q:chartSpace>";
  pugi::xml_document doc = Parse(xml);
  const std::vector<std::string_view> order = {"Sheet1", "Other"};
  const parser::SheetRemovalTransform removal(order, 1);

  EXPECT_EQ(chart_ref::rewrite_sites(doc, removal), 1U);
  EXPECT_EQ(Texts(doc), (std::vector<std::string>{"#REF!"}));
  std::ostringstream serialized;
  doc.document_element().print(serialized, "", pugi::format_raw);
  EXPECT_NE(serialized.str().find("<q:numCache><q:formatCode>General</q:formatCode>"), std::string::npos)
      << serialized.str();
  EXPECT_NE(serialized.str().find("<q:numCache xmlns:q=\"urn:foreign\"><q:formatCode>keep</q:formatCode>"),
            std::string::npos)
      << serialized.str();
}

TEST(ChartRef, FilteredTitlesKeepHolderLocalNamespaceDeclarations) {
  const std::string xml =
      std::string("<q:chartSpace xmlns:q=\"") + chart_ns::kNsC +
      "\"><q:chart><q:plotArea><q:barChart>"
      "<q:ser><m:tx xmlns:m=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><m:strRef>"
      "<m:f>Other!$A$1</m:f></m:strRef></m:tx><q:val><q:numRef><q:f>Other!$A$2</q:f></q:numRef></q:val></q:ser>"
      "<q:ser><q:tx xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><strRef><f>Other!$B$1</f>"
      "</strRef></q:tx><q:val><q:numRef><q:f>Other!$B$2</q:f></q:numRef></q:val></q:ser>"
      "</q:barChart></q:plotArea></q:chart></q:chartSpace>";
  pugi::xml_document doc = Parse(xml);
  const std::vector<std::string_view> order = {"Sheet1", "Other"};
  const parser::SheetRemovalTransform removal(order, 1);

  EXPECT_EQ(chart_ref::rewrite_sites(doc, removal), 4U);
  EXPECT_EQ(Texts(doc), (std::vector<std::string>{"#REF!", "#REF!", "#REF!", "#REF!"}));
  std::ostringstream serialized;
  doc.document_element().print(serialized, "", pugi::format_raw);
  const std::string result = serialized.str();
  EXPECT_NE(result.find("<c15:tx><m:strRef xmlns:m=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"),
            std::string::npos)
      << result;
  EXPECT_NE(result.find("<c15:tx><strRef xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"),
            std::string::npos)
      << result;

  pugi::xml_document reparsed;
  EXPECT_TRUE(reparsed.load_string(result.c_str())) << result;
}

TEST(ChartRef, EnumeratesChartexNameSites) {
  const pugi::xml_document doc = Parse(
      "<cx:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chart><cx:plotArea>"
      "<cx:data id=\"0\"><cx:strDim type=\"cat\"><cx:f>_xlchart.v1.0</cx:f></cx:strDim>"
      "<cx:numDim type=\"val\"><cx:f dir=\"row\">Sheet1!$B$2:$B$5</cx:f></cx:numDim></cx:data></cx:plotArea></cx:chart>"
      "</cx:chartSpace>");
  const std::vector<std::string> expected = {"_xlchart.v1.0", "Sheet1!$B$2:$B$5"};
  EXPECT_EQ(Texts(doc), expected);
}

TEST(ChartRef, RowInsertShiftsEverySiteAndKeepsCaches) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$B$1", "Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3") + kChartClose);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, InsertRows(0, 2)), 3U);
  const std::vector<std::string> expected = {"Sheet1!$B$3", "Sheet1!$A$4:$A$5", "Sheet1!$B$4:$B$5"};
  EXPECT_EQ(Texts(doc), expected);
  EXPECT_STREQ(doc.select_node("//c:tx//c:v").node().child_value(), "old");
  EXPECT_STREQ(doc.select_node("//c:val//c:formatCode").node().child_value(), "0.0");
}

TEST(ChartRef, RowDeleteShrinksARange) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$B$1", "Sheet1!$A$2:$A$5", "Sheet1!$B$2:$B$5") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kRow, parser::RowColEdit::kDelete, 2, 1);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, del), 2U);
  const std::vector<std::string> expected = {"Sheet1!$B$1", "Sheet1!$A$2:$A$4", "Sheet1!$B$2:$B$4"};
  EXPECT_EQ(Texts(doc), expected);
}

TEST(ChartRef, SheetRenameRewritesQualifiers) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$B$1", "Sheet1!$A$2:$A$3", "Other!$B$2") + kChartClose);
  const parser::SheetRenameTransform rename("Sheet1", "Data Set");
  EXPECT_EQ(chart_ref::rewrite_sites(doc, rename), 2U);
  const std::vector<std::string> expected = {"'Data Set'!$B$1", "'Data Set'!$A$2:$A$3", "Other!$B$2"};
  EXPECT_EQ(Texts(doc), expected);
}

TEST(ChartRef, FilteredSeriesAndFullRefSitesFollow) {
  const std::string xml =
      std::string(kChartOpen) +
      "<c:ser><c:cat><c:strRef><c:extLst><c:ext><c15:fullRef><c15:sqref>Sheet1!$A$2:$A$5</c15:sqref></c15:fullRef>"
      "</c:ext></c:extLst><c:f>(Sheet1!$A$2,Sheet1!$A$4:$A$5)</c:f></c:strRef></c:cat></c:ser>"
      "<c:extLst><c:ext><c15:filteredBarSeries><c15:ser><c:tx><c:strRef><c:extLst><c:ext><c15:formulaRef>"
      "<c15:sqref>Sheet1!$C$1</c15:sqref></c15:formulaRef></c:ext></c:extLst></c:strRef></c:tx></c15:ser>"
      "</c15:filteredBarSeries></c:ext></c:extLst>" +
      kChartClose;
  pugi::xml_document doc = Parse(xml);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, InsertRows(0, 1)), 3U);
  const std::vector<std::string> expected = {"Sheet1!$A$3:$A$6", "(Sheet1!$A$3,Sheet1!$A$5:$A$6)", "Sheet1!$C$2"};
  EXPECT_EQ(Texts(doc), expected);
}

TEST(ChartRef, UnparseableSiteStaysWhileOthersContinue) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$B$1", "Sheet1!$A$", "Sheet1!$B$2:$B$3") + kChartClose);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, InsertRows(0, 1)), 2U);
  const std::vector<std::string> expected = {"Sheet1!$B$2", "Sheet1!$A$", "Sheet1!$B$3:$B$4"};
  EXPECT_EQ(Texts(doc), expected);
}

TEST(ChartRef, ChartexNameSiteIsLeftAlone) {
  pugi::xml_document doc = Parse(
      "<cx:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chart>"
      "<cx:f>_xlchart.v1.0</cx:f><cx:f>Sheet1!$B$2:$B$5</cx:f></cx:chart></cx:chartSpace>");
  EXPECT_EQ(chart_ref::rewrite_sites(doc, InsertRows(0, 1)), 1U);
  const std::vector<std::string> expected = {"_xlchart.v1.0", "Sheet1!$B$3:$B$6"};
  EXPECT_EQ(Texts(doc), expected);
}

TEST(ChartRef, DeletedSheetResetsCachesToTheExcelShape) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Other!$B$1", "Other!$A$2:$A$3", "Other!$B$2:$B$3") + kChartClose);
  const std::vector<std::string_view> order = {"Sheet1", "Other"};
  const parser::SheetRemovalTransform removal(order, 1);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, removal), 3U);
  const std::vector<std::string> expected = {"#REF!", "#REF!", "#REF!"};
  EXPECT_EQ(Texts(doc), expected);

  EXPECT_STREQ(doc.select_node("//c15:tx/c:strRef/c:strCache/c:ptCount/@val").attribute().value(), "1");
  EXPECT_STREQ(doc.select_node("//c15:tx//c:pt[@idx='0']/c:v").node().child_value(), "#REF!");
  EXPECT_FALSE(doc.select_node("//c15:cat//c:strCache"));
  EXPECT_FALSE(doc.select_node("//c15:cat//c:numCache"));
  EXPECT_STREQ(doc.select_node("//c:val/c:numRef/c:numCache/c:formatCode").node().child_value(), "General");
  EXPECT_STREQ(doc.select_node("//c:val//c:ptCount/@val").attribute().value(), "1");
  EXPECT_EQ(doc.select_nodes("//c:val//c:pt").size(), 1U);
  EXPECT_STREQ(doc.select_node("//c:val//c:pt[@idx='0']/c:v").node().child_value(), "1");
}

TEST(ChartRef, DeletedNameAndCategoryMoveToFilteredTitles) {
  const std::string c16 =
      "<c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"><c16:uniqueId val=\"{00000002}\"/></c:ext>";
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx>" + StrRef("Other!$B$1") +
            "</c:tx><c:invertIfNegative val=\"0\"/><c:cat>" + StrRef("Other!$A$2:$A$4") + "</c:cat><c:val>" +
            NumRef("Other!$B$2:$B$4") + "</c:val><c:extLst>" + c16 + "</c:extLst></c:ser>" + kChartClose);
  const std::vector<std::string_view> order = {"Sheet1", "Other"};
  const parser::SheetRemovalTransform removal(order, 1);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, removal), 3U);
  std::ostringstream xml;
  doc.select_node("//c:ser").node().print(xml, "", pugi::format_raw);
  // The shape Excel for Mac 16.113 saves after deleting the referenced sheet.
  const std::string c15 = "uri=\"{02D57815-91ED-43cb-92C2-25804820EDAC}\"";
  const std::string formula_ref =
      "<c:extLst><c:ext " + c15 + "><c15:formulaRef><c15:sqref>#REF!</c15:sqref></c15:formulaRef></c:ext></c:extLst>";
  const std::string expected =
      "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:invertIfNegative val=\"0\"/><c:val><c:numRef><c:f>#REF!</c:f>"
      "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt>"
      "</c:numCache></c:numRef></c:val><c:extLst><c:ext " +
      c15 +
      " xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\"><c15:filteredSeriesTitle><c15:tx>"
      "<c:strRef>" +
      formula_ref +
      "<c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>#REF!</c:v></c:pt>"
      "</c:strCache></c:strRef></c15:tx></c15:filteredSeriesTitle></c:ext><c:ext " +
      c15 +
      " xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\"><c15:filteredCategoryTitle><c15:cat>"
      "<c:strRef>" +
      formula_ref + "</c:strRef></c15:cat></c15:filteredCategoryTitle></c:ext>" + c16 + "</c:extLst></c:ser>";
  EXPECT_EQ(xml.str(), expected);
}

TEST(ChartRef, InSheetDeletionKeepsTheSheetQualifier) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$B$1", "Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kRow, parser::RowColEdit::kDelete, 1, 2);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, del), 2U);
  const std::vector<std::string> expected = {"Sheet1!$B$1", "Sheet1!#REF!", "Sheet1!#REF!"};
  EXPECT_EQ(Texts(doc), expected);
  EXPECT_TRUE(doc.select_node("//c:ser/c:tx"));
  EXPECT_FALSE(doc.select_node("//c:ser/c:cat"));
  EXPECT_TRUE(doc.select_node("//c15:filteredCategoryTitle/c15:cat/c:strRef"));
  EXPECT_FALSE(doc.select_node("//c15:filteredSeriesTitle"));
}

TEST(ChartRef, DeletedNameKeepsAQuotedQualifier) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("'My Data'!$B$1", "'My Data'!$A$2", "'My Data'!$B$2") + kChartClose);
  const parser::RowColShiftTransform del("My Data", parser::RowColAxis::kRow, parser::RowColEdit::kDelete, 1, 1);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, del), 2U);
  const std::vector<std::string> expected = {"'My Data'!$B$1", "'My Data'!#REF!", "'My Data'!#REF!"};
  EXPECT_EQ(Texts(doc), expected);
}

TEST(ChartRef, OneDeletedAreaMovesAUnionCategory) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) +
            Series("Sheet1!$B$1", "(Sheet1!$A$2:$A$3,Sheet1!$A$5)", "(Sheet1!$B$2:$B$3,Sheet1!$B$5)") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kRow, parser::RowColEdit::kDelete, 4, 1);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, del), 2U);
  const std::vector<std::string> expected = {"Sheet1!$B$1", "(Sheet1!$B$2:$B$3,Sheet1!#REF!)",
                                             "(Sheet1!$A$2:$A$3,Sheet1!#REF!)"};
  EXPECT_EQ(Texts(doc), expected);
  EXPECT_FALSE(doc.select_node("//c:ser/c:cat"));
  EXPECT_FALSE(doc.select_node("//c15:cat//c:strCache"));
  EXPECT_STREQ(doc.select_node("//c:val//c:formatCode").node().child_value(), "General");
  EXPECT_STREQ(doc.select_node("//c:val//c:ptCount/@val").attribute().value(), "1");
}

TEST(ChartRef, NameOfASeriesThatLostItsValuesStaysWhileAnotherKeepsValues) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$C$1", "Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5") +
            Series("Sheet1!$D$1", "Sheet1!$B$2:$B$5", "Sheet1!$D$2:$D$5") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kCol, parser::RowColEdit::kDelete, 3, 1);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, del), 2U);
  const std::vector<std::string> expected = {"Sheet1!$C$1",  "Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5",
                                             "Sheet1!#REF!", "Sheet1!$B$2:$B$5", "Sheet1!#REF!"};
  EXPECT_EQ(Texts(doc), expected);
  EXPECT_FALSE(doc.select_node("//c15:filteredSeriesTitle"));
  EXPECT_STREQ(doc.select_node("(//c:ser)[2]/c:tx//c:v").node().child_value(), "#REF!");
}

TEST(ChartRef, NamesMoveWhenEverySeriesLostItsValues) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$C$1", "Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5") +
            Series("Sheet1!$D$1", "Sheet1!$B$2:$B$5", "Sheet1!$D$2:$D$5") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kCol, parser::RowColEdit::kDelete, 2, 2);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, del), 4U);
  EXPECT_FALSE(doc.select_node("//c:ser/c:tx"));
  EXPECT_EQ(doc.select_nodes("//c15:filteredSeriesTitle").size(), 2U);
}

TEST(ChartRef, CategoryOfASeriesThatLostItsValuesStillMoves) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$C$1", "Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5") +
            Series("Sheet1!$D$1", "Sheet1!$B$2:$B$5", "Sheet1!$D$2:$D$5") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kCol, parser::RowColEdit::kDelete, 1, 2);
  chart_ref::rewrite_sites(doc, del);
  EXPECT_STREQ(doc.select_node("(//c:ser)[1]/c:tx/c:strRef/c:f").node().child_value(), "Sheet1!#REF!");
  EXPECT_FALSE(doc.select_node("//c:ser/c:cat"));
  EXPECT_EQ(doc.select_nodes("//c15:filteredCategoryTitle").size(), 2U);
  EXPECT_FALSE(doc.select_node("//c15:filteredSeriesTitle"));
}

TEST(ChartRef, DeletedTitleCellStaysInTheTitle) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$C$1", "Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5") +
            "<c:title><c:tx>" + StrRef("Sheet1!$E$1") + "</c:tx><c:overlay val=\"0\"/></c:title>" + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kCol, parser::RowColEdit::kDelete, 4, 1);
  EXPECT_EQ(chart_ref::rewrite_sites(doc, del), 1U);
  std::ostringstream xml;
  doc.select_node("//c:title").node().print(xml, "", pugi::format_raw);
  // Excel for Mac 16.113 keeps a deleted title link in place.
  EXPECT_EQ(xml.str(),
            "<c:title><c:tx><c:strRef><c:f>Sheet1!#REF!</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\">"
            "<c:v>#REF!</c:v></c:pt></c:strCache></c:strRef></c:tx><c:overlay val=\"0\"/></c:title>");
}

TEST(ChartRef, MultiLevelCategoryDownToOneLevelBecomesAStrRef) {
  pugi::xml_document doc = Parse(std::string(kChartOpen) + MultiLevelSeries("Sheet1!$A$2:$B$3") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kCol, parser::RowColEdit::kDelete, 0, 1);
  chart_ref::rewrite_sites(doc, del);
  EXPECT_STREQ(doc.select_node("//c:cat/c:strRef/c:f").node().child_value(), "Sheet1!$A$2:$A$3");
  EXPECT_FALSE(doc.select_node("//c:multiLvlStrRef"));
  EXPECT_FALSE(doc.select_node("//c:multiLvlStrCache"));
}

TEST(ChartRef, MultiLevelCategoryShiftKeepsItsLevels) {
  pugi::xml_document doc = Parse(std::string(kChartOpen) + MultiLevelSeries("Sheet1!$A$2:$B$3") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kRow, parser::RowColEdit::kDelete, 2, 1);
  chart_ref::rewrite_sites(doc, del);
  EXPECT_STREQ(doc.select_node("//c:cat/c:multiLvlStrRef/c:f").node().child_value(), "Sheet1!$A$2:$B$2");
  EXPECT_TRUE(doc.select_node("//c:multiLvlStrCache"));
}

TEST(ChartRef, DeletedMultiLevelCategoryMovesWithoutACache) {
  pugi::xml_document doc = Parse(std::string(kChartOpen) + MultiLevelSeries("Sheet1!$A$2:$B$3") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kCol, parser::RowColEdit::kDelete, 0, 2);
  chart_ref::rewrite_sites(doc, del);
  EXPECT_FALSE(doc.select_node("//c:ser/c:cat"));
  EXPECT_STREQ(doc.select_node("//c15:filteredCategoryTitle/c15:cat/c:multiLvlStrRef//c15:sqref").node().child_value(),
               "Sheet1!#REF!");
  EXPECT_FALSE(doc.select_node("//c:multiLvlStrCache"));
}

TEST(ChartRef, PartialRefLossKeepsTheCache) {
  pugi::xml_document doc =
      Parse(std::string(kChartOpen) + Series("Sheet1!$B$1", "Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3") + kChartClose);
  const parser::RowColShiftTransform del("Sheet1", parser::RowColAxis::kRow, parser::RowColEdit::kDelete, 1, 1);
  chart_ref::rewrite_sites(doc, del);
  EXPECT_STREQ(doc.select_node("//c:cat//c:f").node().child_value(), "Sheet1!$A$2:$A$2");
  EXPECT_STREQ(doc.select_node("//c:cat//c:v").node().child_value(), "old");
  EXPECT_STREQ(doc.select_node("//c:val//c:formatCode").node().child_value(), "0.0");
}

TEST(ChartRef, RewriteAllStoresOnlyChangedParts) {
  const std::string chart =
      std::string(kChartOpen) + Series("Sheet1!$B$1", "Sheet1!$A$2:$A$3", "Other!$B$2") + kChartClose;
  Workbook wb = WithChart(chart);
  const std::string before = PartText(wb);

  const parser::RowColShiftTransform other("Other", parser::RowColAxis::kRow, parser::RowColEdit::kInsert, 0, 1);
  const parser::RowColShiftTransform unrelated("Missing", parser::RowColAxis::kRow, parser::RowColEdit::kInsert, 0, 1);
  ASSERT_TRUE(static_cast<bool>(chart_ref::rewrite_all(wb, unrelated)));
  EXPECT_EQ(PartText(wb), before);

  ASSERT_TRUE(static_cast<bool>(chart_ref::rewrite_all(wb, InsertRows(0, 1))));
  const std::string after = PartText(wb);
  EXPECT_NE(after, before);
  EXPECT_NE(after.find("<c:f>Sheet1!$B$2</c:f>"), std::string::npos);
  EXPECT_NE(after.find("<c:f>Sheet1!$A$3:$A$4</c:f>"), std::string::npos);
  EXPECT_NE(after.find("<c:f>Other!$B$2</c:f>"), std::string::npos);

  ASSERT_TRUE(static_cast<bool>(chart_ref::rewrite_all(wb, other)));
  EXPECT_NE(PartText(wb).find("<c:f>Other!$B$3</c:f>"), std::string::npos);
}

TEST(ChartRef, RewriteAllFollowsASheetRename) {
  Workbook wb = WithChart(std::string(kChartOpen) + Series("Sheet1!$B$1", "Sheet1!$A$2", "Sheet1!$B$2") + kChartClose);
  const parser::SheetRenameTransform rename("Sheet1", "Sales");
  ASSERT_TRUE(static_cast<bool>(chart_ref::rewrite_all(wb, rename)));
  const std::string after = PartText(wb);
  EXPECT_NE(after.find("<c:f>Sales!$B$1</c:f>"), std::string::npos);
  EXPECT_EQ(after.find("Sheet1"), std::string::npos);
}

}  // namespace formulon
