// Round trip of the pivot `<filters>` block, and the load -> save sweep over
// every `<pivotTableDefinition>` attribute and child the reader keeps.

#include <cstddef>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_table_reader.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_table_reader_test_helpers.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {
using namespace pivot_table_reader_test_support;

// Region on rows, an unplaced Date field, and Amount summed twice.
pivot::PivotTable FilterTable() {
  pivot::PivotTable table;
  table.set_name("P");
  table.set_pivot_cache_id(1U);
  table.set_anchor(0U, 0U, 4U, 3U);
  const std::pair<const char*, pivot::PivotAxis> fields[] = {
      {"Region", pivot::PivotAxis::Row}, {"Date", pivot::PivotAxis::None}, {"Amount", pivot::PivotAxis::Value}};
  for (const auto& [name, axis] : fields) {
    pivot::PivotField field;
    field.source_name = name;
    field.axis = axis;
    table.mutable_fields().push_back(std::move(field));
  }
  table.mutable_row_field_order().push_back(0U);
  for (const char* name : {"Sum of Amount", "Max of Amount"}) {
    pivot::PivotDataField data;
    data.name = name;
    data.field_index = 2U;
    table.mutable_data_fields().push_back(std::move(data));
  }
  table.set_col_values_position(0U);
  return table;
}

pivot::PivotFilter MakeFilter(pivot::FilterType type, std::string field, std::variant<int, double, std::string> value,
                              std::variant<std::monostate, int, double> high = {}, std::uint32_t data_field = 0U) {
  pivot::PivotFilter filter;
  filter.axis = pivot::PivotAxis::Row;
  filter.field_name = std::move(field);
  filter.type = type;
  filter.value = std::move(value);
  filter.value_high = high;
  filter.data_field_index = data_field;
  return filter;
}

std::string FiltersBlock(const std::string& xml) {
  const std::size_t begin = xml.find("<filters");
  const std::size_t end = xml.rfind("</filters>");
  if (begin == std::string::npos || end == std::string::npos) {
    return {};
  }
  return xml.substr(begin, end + std::string_view("</filters>").size() - begin);
}

struct GeneratedCase {
  const char* label;
  pivot::PivotFilter filter;
  std::string_view expected;
};

// Every public type is written as Excel writes it from the filter dialogs
// (the value and date forms match `pivot_value_date_filters.xlsx` byte for
// byte bar the `id`), and reads back into `active_filters()` unchanged.
TEST(PivotFiltersRoundTrip, EachPublicTypeWritesExcelsSpellingAndReadsBack) {
  const GeneratedCase cases[] = {
      {"contains", MakeFilter(pivot::FilterType::LabelContains, "Region", std::string("or")),
       "<filter fld=\"0\" type=\"captionContains\" evalOrder=\"-1\" id=\"1\" stringValue1=\"or\"><autoFilter "
       "ref=\"A1\"><filterColumn colId=\"0\"><customFilters><customFilter val=\"*or*\"/></customFilters>"
       "</filterColumn></autoFilter></filter>"},
      {"begins", MakeFilter(pivot::FilterType::LabelBeginsWith, "Region", std::string("a&b")),
       "<filter fld=\"0\" type=\"captionBeginsWith\" evalOrder=\"-1\" id=\"1\" stringValue1=\"a&amp;b\"><autoFilter "
       "ref=\"A1\"><filterColumn colId=\"0\"><customFilters><customFilter val=\"a&amp;b*\"/></customFilters>"
       "</filterColumn></autoFilter></filter>"},
      {"top", MakeFilter(pivot::FilterType::ValueTop10, "Region", 2),
       "<filter fld=\"0\" type=\"count\" evalOrder=\"-1\" id=\"1\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">"
       "<filterColumn colId=\"0\"><top10 val=\"2\" filterVal=\"2\"/></filterColumn></autoFilter></filter>"},
      {"greater", MakeFilter(pivot::FilterType::ValueGreaterThan, "Region", 100.0, {}, 1U),
       "<filter fld=\"0\" type=\"valueGreaterThan\" evalOrder=\"-1\" id=\"1\" iMeasureFld=\"1\"><autoFilter "
       "ref=\"A1\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"greaterThan\" val=\"100\"/>"
       "</customFilters></filterColumn></autoFilter></filter>"},
      {"between", MakeFilter(pivot::FilterType::ValueBetween, "Region", 100.0, 175.5),
       "<filter fld=\"0\" type=\"valueBetween\" evalOrder=\"-1\" id=\"1\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">"
       "<filterColumn colId=\"0\"><customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" "
       "val=\"100\"/><customFilter operator=\"lessThanOrEqual\" val=\"175.5\"/></customFilters></filterColumn>"
       "</autoFilter></filter>"},
      {"date", MakeFilter(pivot::FilterType::LabelDate, "Date", 45292.0, 45382.0),
       "<filter fld=\"1\" type=\"dateBetween\" evalOrder=\"-1\" id=\"1\"><autoFilter ref=\"A1\"><filterColumn "
       "colId=\"0\"><customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"45292\"/>"
       "<customFilter operator=\"lessThanOrEqual\" val=\"45382\"/></customFilters></filterColumn></autoFilter>"
       "</filter>"},
  };
  for (const GeneratedCase& c : cases) {
    SCOPED_TRACE(c.label);
    pivot::PivotTable table = FilterTable();
    table.mutable_active_filters().push_back(c.filter);
    const std::string written = write_pivot_table_definition(table);
    EXPECT_EQ(FiltersBlock(written), "<filters count=\"1\">" + std::string(c.expected) + "</filters>");
    // Schema order: after the style slot, before the tail extension list.
    EXPECT_LT(written.find("<filters"), written.find("<extLst"));

    auto back_or = read_pivot_table_definition(Bytes(written));
    ASSERT_TRUE(static_cast<bool>(back_or)) << back_or.error().message;
    const pivot::PivotTable& back = back_or.value();
    ASSERT_EQ(back.active_filters().size(), 1U);
    const pivot::PivotFilter& f = back.active_filters()[0];
    EXPECT_EQ(f.type, c.filter.type);
    EXPECT_EQ(f.value, c.filter.value);
    EXPECT_EQ(f.value_high, c.filter.value_high);
    EXPECT_EQ(f.data_field_index, c.filter.data_field_index);
    EXPECT_EQ(f.axis, pivot::PivotAxis::Row);
    ASSERT_TRUE(f.source.has_value());
    EXPECT_EQ(f.source->id, 1U);
    EXPECT_TRUE(back.retained_filters().empty());
    EXPECT_TRUE(back.authored_value_filters().empty());
    EXPECT_TRUE(back.authored_caption_filters().empty());
    EXPECT_EQ(FiltersBlock(write_pivot_table_definition(back)), FiltersBlock(written));
  }
}

// A value filter naming something off its axis prunes the axis's innermost
// field, so that is the field the file names.
TEST(PivotFiltersRoundTrip, ValueFilterOffItsAxisIsWrittenAgainstTheInnermostField) {
  pivot::PivotTable table = FilterTable();
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::ValueTop10, "Sum of Amount", 3));
  EXPECT_NE(write_pivot_table_definition(table).find("<filter fld=\"0\" type=\"count\""), std::string::npos);
}

// An entry that filters nothing and has no file spelling is not written.
TEST(PivotFiltersRoundTrip, EntriesThatFilterNothingAreNotWritten) {
  pivot::PivotTable table = FilterTable();
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::ValueBetween, "Region", 1.0));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelContains, "Nowhere", std::string("x")));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::ValueTop10, "Region", 1, {}, 7U));
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelDate, "Date", std::string("x"), 1.0));
  EXPECT_EQ(write_pivot_table_definition(table).find("<filters"), std::string::npos);
}

std::string MixedFiltersDefinition() {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\" dataCaption=\"Values\">");
  xml.append("<location ref=\"A1:C6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  xml.append(
      "<pivotFields count=\"3\"><pivotField axis=\"axisRow\" showAll=\"0\"/><pivotField showAll=\"0\"/>"
      "<pivotField dataField=\"1\" showAll=\"0\"/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<dataFields count=\"1\"><dataField name=\"Sum\" fld=\"2\" subtotal=\"sum\"/></dataFields>");
  xml.append(
      "<filters count=\"5\">"
      "<filter fld=\"0\" type=\"captionEqual\" evalOrder=\"-1\" id=\"4\" stringValue1=\"North\"><autoFilter "
      "ref=\"A1\"><filterColumn colId=\"0\"><filters><filter val=\"North\"/></filters></filterColumn></autoFilter>"
      "</filter>"
      "<filter fld=\"0\" type=\"count\" evalOrder=\"-1\" id=\"2\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><top10 val=\"2\" filterVal=\"2\"/></filterColumn></autoFilter></filter>"
      "<filter fld=\"0\" type=\"percent\" evalOrder=\"-1\" id=\"7\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><top10 percent=\"1\" val=\"70\" filterVal=\"70\"/></filterColumn></autoFilter>"
      "</filter>"
      "<filter fld=\"1\" type=\"Q2\" evalOrder=\"-1\" id=\"3\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<dynamicFilter type=\"Q2\"/></filterColumn></autoFilter></filter>"
      "<filter fld=\"1\" type=\"someFutureFamily\" evalOrder=\"-1\" id=\"5\" name=\"x\"><autoFilter ref=\"A1\"/>"
      "<extLst><ext uri=\"{00000000-0000-0000-0000-000000000001}\"/></extLst></filter>"
      "</filters>");
  xml.append("</pivotTableDefinition>");
  return xml;
}

// A loaded block keeps every entry, in order and byte for byte, whatever the
// model makes of it: the public one is listed, the rest are retained.
TEST(PivotFiltersRoundTrip, LoadedBlockIsReEmittedVerbatim) {
  const std::string source = MixedFiltersDefinition();
  auto table_or = read_pivot_table_definition(Bytes(source));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.active_filters().size(), 1U);
  EXPECT_EQ(table.active_filters()[0].type, pivot::FilterType::ValueTop10);
  ASSERT_EQ(table.retained_filters().size(), 4U);
  EXPECT_EQ(table.retained_filters()[0].active_before, 0U);
  EXPECT_EQ(table.retained_filters()[1].active_before, 1U);
  EXPECT_EQ(table.retained_filters()[3].active_before, 1U);
  EXPECT_EQ(table.authored_caption_filters().size(), 1U);
  ASSERT_EQ(table.authored_value_filters().size(), 1U);
  EXPECT_EQ(table.authored_value_filters()[0].active_before, 1U);
  EXPECT_EQ(table.authored_recurring_filters().size(), 1U);
  EXPECT_EQ(FiltersBlock(write_pivot_table_definition(table)), FiltersBlock(source));
}

// Removing the public entry keeps the retained ones where they were; a new
// entry is appended under an id no loaded entry uses.
TEST(PivotFiltersRoundTrip, EditingThePublicListKeepsRetainedEntriesInPlace) {
  auto table_or = read_pivot_table_definition(Bytes(MixedFiltersDefinition()));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable& table = table_or.value();
  table.erase_active_filters_if([](const pivot::PivotFilter&) { return true; });
  EXPECT_EQ(table.retained_filters()[1].active_before, 0U);
  EXPECT_EQ(table.authored_value_filters()[0].active_before, 0U);
  std::string written = write_pivot_table_definition(table);
  std::string block = FiltersBlock(written);
  EXPECT_EQ(block.find("type=\"count\""), std::string::npos) << block;
  EXPECT_EQ(block.rfind("<filters count=\"4\">", 0), 0U) << block;
  EXPECT_LT(block.find("captionEqual"), block.find("percent"));
  EXPECT_LT(block.find("percent"), block.find("Q2"));

  table.mutable_fields()[0].source_name = "Region";
  table.mutable_active_filters().push_back(MakeFilter(pivot::FilterType::LabelBeginsWith, "Region", std::string("S")));
  written = write_pivot_table_definition(table);
  block = FiltersBlock(written);
  EXPECT_EQ(block.rfind("<filters count=\"5\">", 0), 0U) << block;
  EXPECT_NE(block.find("type=\"captionBeginsWith\" evalOrder=\"-1\" id=\"8\""), std::string::npos) << block;
  EXPECT_LT(block.find("someFutureFamily"), block.find("captionBeginsWith"));
}

// A decoded entry whose field moved is re-serialised against the new index
// instead of re-emitting a stale `fld`.
TEST(PivotFiltersRoundTrip, DecodedEntryFollowsItsFieldWhenTheIndexChanges) {
  auto table_or = read_pivot_table_definition(Bytes(MixedFiltersDefinition()));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable& table = table_or.value();
  table.mutable_fields()[0].source_name = "Region";
  table.mutable_fields()[1].source_name = "Date";
  table.mutable_active_filters()[0].field_name = "Region";
  // Unchanged: the source bytes, `filterVal` included.
  EXPECT_NE(FiltersBlock(write_pivot_table_definition(table)).find("id=\"2\" iMeasureFld=\"0\""), std::string::npos);
  // Region moves to index 1.
  std::swap(table.mutable_fields()[0], table.mutable_fields()[1]);
  table.mutable_row_field_order() = {1U};
  const std::string block = FiltersBlock(write_pivot_table_definition(table));
  EXPECT_NE(block.find("<filter fld=\"1\" type=\"count\" evalOrder=\"-1\" id=\"8\""), std::string::npos) << block;
}

// Entries the public vocabulary cannot express exactly stay retained.
TEST(PivotFiltersRoundTrip, InexactPublicShapesStayRetained) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"2\"><pivotField axis=\"axisRow\"/><pivotField/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<filters count=\"4\">");
  // Bottom N.
  xml.append(
      "<filter fld=\"0\" type=\"count\" id=\"1\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 top=\"0\" val=\"2\"/></filterColumn></autoFilter></filter>");
  // A pattern that is not `*stringValue1*`.
  xml.append(
      "<filter fld=\"0\" type=\"captionContains\" id=\"2\" stringValue1=\"a\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter val=\"*b*\"/></customFilters></filterColumn>"
      "</autoFilter></filter>");
  // A value filter on a field that sits on no axis.
  xml.append(
      "<filter fld=\"1\" type=\"valueGreaterThan\" id=\"3\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"greaterThan\" val=\"1\"/></customFilters></filterColumn>"
      "</autoFilter></filter>");
  // No `fld` at all.
  xml.append("<filter type=\"captionBeginsWith\" id=\"4\" stringValue1=\"a\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters></pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().active_filters().empty());
  EXPECT_EQ(table_or.value().retained_filters().size(), 4U);
  EXPECT_EQ(FiltersBlock(write_pivot_table_definition(table_or.value())), FiltersBlock(xml));
}

// ---------------------------------------------------------------------------
// Attribute sweep
// ---------------------------------------------------------------------------

std::map<std::string, std::string> Attributes(const pugi::xml_node& node) {
  std::map<std::string, std::string> out;
  for (const pugi::xml_attribute attr : node.attributes()) {
    const std::string_view name = attr.name();
    if (name == "xmlns" || name.rfind("xmlns:", 0) == 0) {
      continue;  // Declarations move to wherever the writer binds them.
    }
    out.emplace(attr.name(), attr.value());
  }
  return out;
}

// Same element names, same attribute sets (order-insensitive), same element
// children in the same order.
void ExpectSameTree(const pugi::xml_node& want, const pugi::xml_node& got, const std::string& path) {
  const std::string here = path + "/" + want.name();
  ASSERT_STREQ(want.name(), got.name()) << path;
  EXPECT_EQ(Attributes(want), Attributes(got)) << here;
  std::vector<pugi::xml_node> want_children;
  std::vector<pugi::xml_node> got_children;
  for (const pugi::xml_node child : want.children()) {
    if (child.type() == pugi::node_element) {
      want_children.push_back(child);
    }
  }
  for (const pugi::xml_node child : got.children()) {
    if (child.type() == pugi::node_element) {
      got_children.push_back(child);
    }
  }
  ASSERT_EQ(want_children.size(), got_children.size()) << here;
  for (std::size_t i = 0; i < want_children.size(); ++i) {
    ExpectSameTree(want_children[i], got_children[i], here + "[" + std::to_string(i) + "]");
  }
}

constexpr std::string_view kX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

// One definition exercising every attribute and child kept across a load ->
// save with no edit in between: modelled state, attribute passthrough on the
// root / fields / items, per-field and table-level report forms that depart
// from the derived one, field children, data-field extensions, the
// position-keyed raw bins, and a mixed `<filters>` block.
std::string SweepDefinition() {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs);
  xml.append(
      " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"xr\" "
      "xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" "
      "xr:uid=\"{11111111-2222-3333-4444-555555555555}\" name=\"PT\" cacheId=\"3\" applyNumberFormats=\"0\" "
      "applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" "
      "applyWidthHeightFormats=\"1\" dataCaption=\"Values\" grandTotalCaption=\"All\" errorCaption=\"#\" "
      "showError=\"1\" missingCaption=\"-\" updatedVersion=\"8\" minRefreshableVersion=\"3\" "
      "useAutoFormatting=\"1\" itemPrintTitles=\"1\" createdVersion=\"8\" indent=\"0\" compact=\"1\" "
      "compactData=\"1\" outline=\"0\" outlineData=\"0\" multipleFieldFilters=\"0\" rowGrandTotals=\"0\" "
      "subtotalHiddenItems=\"1\" chartFormat=\"2\" rowHeaderCaption=\"Rows\">");
  xml.append(
      "<location ref=\"A3:E9\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\" rowPageCount=\"1\" "
      "colPageCount=\"1\"/>");
  xml.append("<pivotFields count=\"5\">");
  xml.append(
      "<pivotField axis=\"axisRow\" compact=\"0\" showAll=\"0\" sortType=\"descending\" defaultSubtotal=\"0\" "
      "insertBlankRow=\"1\" numFmtId=\"14\" name=\"Area\"><items count=\"4\"><item x=\"1\" sd=\"0\"/>"
      "<item t=\"blank\"/><item x=\"0\" n=\"Nord\" h=\"1\"/><item t=\"grand\"/></items><autoSortScope><pivotArea "
      "dataOnly=\"0\" outline=\"0\" "
      "fieldPosition=\"0\"><references count=\"1\"><reference field=\"4294967294\" count=\"1\" selected=\"0\">"
      "<x v=\"0\"/></reference></references></pivotArea></autoSortScope><extLst><ext "
      "uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\" xmlns:x14=\"");
  xml.append(kX14).append("\"><x14:pivotField fillDownLabels=\"1\"/></ext></extLst></pivotField>");
  xml.append(
      "<pivotField axis=\"axisCol\" compact=\"0\" outline=\"0\" showAll=\"0\" sumSubtotal=\"1\" "
      "countASubtotal=\"1\"><items count=\"4\"><item x=\"0\"/><item x=\"1\"/><item t=\"sum\"/>"
      "<item t=\"countA\"/></items></pivotField>");
  xml.append(
      "<pivotField axis=\"axisPage\" compact=\"0\" showAll=\"0\" sortType=\"ascending\"><items count=\"3\">"
      "<item x=\"0\"/><item x=\"1\" h=\"1\"/><item t=\"default\" sd=\"0\"/></items></pivotField>");
  xml.append("<pivotField dataField=\"1\" compact=\"0\" showAll=\"0\"/>");
  xml.append("<pivotField compact=\"0\" showAll=\"0\" sortType=\"ascending\" dragToRow=\"0\"/>");
  xml.append("</pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<rowItems count=\"2\"><i><x/></i><i t=\"grand\"><x/></i></rowItems>");
  xml.append("<colFields count=\"2\"><field x=\"1\"/><field x=\"-2\"/></colFields>");
  xml.append("<colItems count=\"1\"><i/></colItems>");
  xml.append("<pageFields count=\"1\"><pageField fld=\"2\" item=\"0\" hier=\"-1\"/></pageFields>");
  xml.append(
      "<dataFields count=\"2\"><dataField name=\"Sum of Amt\" fld=\"3\" subtotal=\"sum\" "
      "showDataAs=\"percentOfTotal\" baseField=\"0\" baseItem=\"0\" numFmtId=\"10\"><extLst>"
      "<ext uri=\"{E15A36E0-9728-4e99-A89B-3F7291B0FE68}\" xmlns:x14=\"");
  xml.append(kX14).append(
      "\"><x14:dataField sourceField=\"0\"/></ext></extLst></dataField>"
      "<dataField name=\"Count of Amt\" fld=\"3\" subtotal=\"count\" baseField=\"0\" baseItem=\"0\"><extLst>"
      "<ext uri=\"{E15A36E0-9728-4e99-A89B-3F7291B0FE68}\" xmlns:x14=\"");
  xml.append(kX14).append(
      "\"><x14:dataField pivotShowAs=\"percentOfParentRow\" sourceField=\"3\" uniqueName=\"[u]\"><x14:extLst>"
      "<x14:ext uri=\"{BBBBBBBB-0000-0000-0000-000000000000}\"/></x14:extLst></x14:dataField></ext><ext "
      "uri=\"{AAAAAAAA-0000-0000-0000-000000000000}\" "
      "xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:dataField "
      "isCountDistinct=\"1\"/></ext></extLst></dataField></dataFields>");
  xml.append(
      "<formats count=\"1\"><format dxfId=\"0\"><pivotArea type=\"all\" dataOnly=\"0\" outline=\"0\" "
      "fieldPosition=\"0\"/></format></formats>");
  xml.append(
      "<conditionalFormats count=\"1\"><conditionalFormat priority=\"1\"><pivotAreas count=\"1\"><pivotArea "
      "type=\"data\" outline=\"0\" fieldPosition=\"0\"/></pivotAreas></conditionalFormat></conditionalFormats>");
  xml.append(
      "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" "
      "outline=\"0\" fieldPosition=\"0\"/></chartFormat></chartFormats>");
  xml.append(
      "<pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" "
      "showRowStripes=\"0\" showColStripes=\"1\" showLastColumn=\"1\"/>");
  xml.append(
      "<filters count=\"3\"><filter fld=\"0\" type=\"captionEqual\" evalOrder=\"-1\" id=\"1\" "
      "stringValue1=\"North\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\"><filters><filter val=\"North\"/>"
      "</filters></filterColumn></autoFilter></filter><filter fld=\"0\" type=\"count\" evalOrder=\"-1\" id=\"2\" "
      "iMeasureFld=\"0\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\"><top10 val=\"2\" filterVal=\"2\"/>"
      "</filterColumn></autoFilter></filter><filter fld=\"4\" type=\"Q2\" evalOrder=\"-1\" id=\"3\"><autoFilter "
      "ref=\"A1\"><filterColumn colId=\"0\"><dynamicFilter type=\"Q2\"/></filterColumn></autoFilter></filter>"
      "</filters>");
  xml.append("<rowHierarchiesUsage count=\"1\"><rowHierarchyUsage hierarchyUsage=\"0\"/></rowHierarchiesUsage>");
  xml.append("<extLst><ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\" xmlns:x14=\"");
  xml.append(kX14).append(
      "\"><x14:pivotTableDefinition hideValuesRow=\"1\" "
      "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/></ext></extLst>");
  xml.append("</pivotTableDefinition>");
  return xml;
}

TEST(PivotDefinitionRoundTrip, LoadSaveKeepsEveryAttributeAndChild) {
  const std::string source = SweepDefinition();
  auto table_or = read_pivot_table_definition(Bytes(source));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const std::string written = write_pivot_table_definition(table_or.value());

  pugi::xml_document want;
  pugi::xml_document got;
  ASSERT_TRUE(want.load_buffer(source.data(), source.size()));
  ASSERT_TRUE(got.load_buffer(written.data(), written.size())) << written;
  ExpectSameTree(want.document_element(), got.document_element(), "");

  auto again_or = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(again_or)) << again_or.error().message;
  EXPECT_EQ(write_pivot_table_definition(again_or.value()), written);
}

// A report form read from the file is the file's until the layout is set
// again, which resets the table and every field to the derived form.
TEST(PivotDefinitionRoundTrip, SettingTheLayoutDropsTheFormsReadFromTheFile) {
  auto table_or = read_pivot_table_definition(Bytes(SweepDefinition()));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.layout(), pivot::PivotLayout::Outline);
  table.set_layout(pivot::PivotLayout::Tabular);
  const std::string written = write_pivot_table_definition(table);
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(written.data(), written.size())) << written;
  const std::map<std::string, std::string> root = Attributes(doc.document_element());
  EXPECT_EQ(root.at("compact"), "0");
  EXPECT_EQ(root.at("compactData"), "0");
  EXPECT_EQ(root.count("outline"), 0U);
  EXPECT_EQ(root.count("outlineData"), 0U);
  EXPECT_EQ(CountOccurrences(written, " compact=\"0\" outline=\"0\""), 5U) << written;
}

// A subtotal selection edited in the model regenerates the subtotal markers;
// markers with no model counterpart (`blank`, `grand`) keep their place.
TEST(PivotDefinitionRoundTrip, EditedSubtotalsRegenerateMarkersAndKeepTheRest) {
  auto table_or = read_pivot_table_definition(Bytes(SweepDefinition()));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable& table = table_or.value();
  table.mutable_fields()[0].subtotal_fns = {pivot::SubtotalFn::Max};
  table.mutable_fields()[2].subtotal_fns = {pivot::SubtotalFn::Sum};
  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<items count=\"5\"><item x=\"1\" sd=\"0\"/><item t=\"blank\"/><item x=\"0\" h=\"1\" "
                         "n=\"Nord\"/><item t=\"max\"/><item t=\"grand\"/></items>"),
            std::string::npos)
      << written;
  EXPECT_NE(written.find("<item x=\"1\" h=\"1\"/><item t=\"sum\"/></items>"), std::string::npos) << written;
  // The x14 data-field extension follows a show-as edit and keeps the rest.
  table.mutable_data_fields()[1].show_as = pivot::ShowValuesAs::RankAscending;
  const std::string edited = write_pivot_table_definition(table);
  EXPECT_NE(edited.find("<x14:dataField pivotShowAs=\"rankAscending\" sourceField=\"3\" uniqueName=\"[u]\">"),
            std::string::npos)
      << edited;
}

}  // namespace
}  // namespace formulon::io
