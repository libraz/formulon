// Pivot-table definition reader tests grouped by XML concern.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_table_reader.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_table_reader_test_helpers.h"
#include "utils/error.h"

namespace formulon::io {
namespace {
using namespace pivot_table_reader_test_support;
TEST(PivotTableReader, ShowDataAsAttributesAreParsed) {
  // <dataField> exercises showDataAs / baseField / baseItem decoding.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"2\">");
  xml.append("  <pivotField axis=\"axisRow\" name=\"R\"/>");
  xml.append("  <pivotField dataField=\"1\" name=\"V\"/>");
  xml.append("</pivotFields>");
  xml.append("<dataFields count=\"1\">");
  xml.append(
      "  <dataField name=\"X\" fld=\"0\" subtotal=\"sum\""
      " showDataAs=\"difference\" baseField=\"1\" baseItem=\"3\"/>");
  xml.append("</dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.data_fields().size(), 1U);
  const pivot::PivotDataField& df = table.data_fields()[0];
  EXPECT_EQ(df.show_as, pivot::ShowValuesAs::DifferenceFrom);
  ASSERT_TRUE(df.show_as_base_field.has_value());
  EXPECT_EQ(*df.show_as_base_field, 1U);
  ASSERT_TRUE(df.show_as_base_item.has_value());
  EXPECT_EQ(*df.show_as_base_item, 3U);
}
TEST(PivotTableReader, EmptyPassthroughWhenNoExtensionsPresent) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or));
  EXPECT_TRUE(table_or.value().raw_passthrough_xml().empty());
}
TEST(PivotTableReader, GrandTotalsDefaultTrueWhenAbsent) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_TRUE(table_or.value().grand_totals_rows());
  EXPECT_TRUE(table_or.value().grand_totals_cols());
}
TEST(PivotTableReader, GrandTotalsOffIsRead) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs);
  xml.append(" name=\"P\" cacheId=\"0\" rowGrandTotals=\"0\" colGrandTotals=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_FALSE(table_or.value().grand_totals_rows());
  EXPECT_FALSE(table_or.value().grand_totals_cols());
}
TEST(PivotTableReader, LocationRequiredAttributesAreRead) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append(
      "<location ref=\"A3:D10\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"3\" "
      "rowPageCount=\"4\" colPageCount=\"5\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_TRUE(table.location_first_header_row().has_value());
  EXPECT_EQ(*table.location_first_header_row(), 1U);
  ASSERT_TRUE(table.location_first_data_row().has_value());
  EXPECT_EQ(*table.location_first_data_row(), 2U);
  ASSERT_TRUE(table.location_first_data_col().has_value());
  EXPECT_EQ(*table.location_first_data_col(), 3U);
  ASSERT_TRUE(table.location_row_page_count().has_value());
  EXPECT_EQ(*table.location_row_page_count(), 4U);
  ASSERT_TRUE(table.location_col_page_count().has_value());
  EXPECT_EQ(*table.location_col_page_count(), 5U);
}
TEST(PivotTableReader, LocationOptionalAttributesAbsentStayEmpty) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  // Only `ref` present; all offset attributes absent.
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  EXPECT_FALSE(table.location_first_header_row().has_value());
  EXPECT_FALSE(table.location_first_data_row().has_value());
  EXPECT_FALSE(table.location_first_data_col().has_value());
  EXPECT_FALSE(table.location_row_page_count().has_value());
  EXPECT_FALSE(table.location_col_page_count().has_value());
}
TEST(PivotTableReader, CustomSubtotalsAndDefaultSubtotalSurviveRoundTrip) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"1\">");
  // defaultSubtotal turned OFF, with explicit Average + Max custom
  // subtotals selected.
  xml.append("  <pivotField axis=\"axisRow\" name=\"R\" defaultSubtotal=\"0\" avgSubtotal=\"1\" maxSubtotal=\"1\"/>");
  xml.append("</pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  const pivot::PivotField& f = table.fields()[0];
  EXPECT_FALSE(f.default_subtotal);
  ASSERT_EQ(f.subtotal_fns.size(), 2U);
  EXPECT_EQ(f.subtotal_fns[0], pivot::SubtotalFn::Average);
  EXPECT_EQ(f.subtotal_fns[1], pivot::SubtotalFn::Max);

  // Write -> read again: the custom selection must not revert to default.
  const std::string round = write_pivot_table_definition(table);
  auto reparsed_or = read_pivot_table_definition(Bytes(round));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
  const pivot::PivotField& f2 = reparsed_or.value().fields()[0];
  EXPECT_FALSE(f2.default_subtotal);
  ASSERT_EQ(f2.subtotal_fns.size(), 2U);
  EXPECT_EQ(f2.subtotal_fns[0], pivot::SubtotalFn::Average);
  EXPECT_EQ(f2.subtotal_fns[1], pivot::SubtotalFn::Max);
}
TEST(PivotTableReader, DefaultSubtotalDefaultsTrueWhenAbsent) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"1\"><pivotField axis=\"axisRow\" name=\"R\"/></pivotFields>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotField& f = table_or.value().fields()[0];
  EXPECT_TRUE(f.default_subtotal);
  EXPECT_TRUE(f.subtotal_fns.empty());
}
TEST(PivotTableReader, RowItemsAndColItemsSurviveRoundTrip) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("<rowItems count=\"2\"><i><x/></i><i t=\"grand\"><x/></i></rowItems>");
  xml.append("<colItems count=\"1\"><i><x/></i></colItems>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  // rowItems / colItems are binned into their schema-position buffers so
  // the writer can re-emit them before <dataFields> (a single tail buffer
  // would place them after <dataFields> and trip Excel's repair).
  const pivot::PivotTable& table = table_or.value();
  EXPECT_NE(table.raw_passthrough_after_row_fields().find("<rowItems"), std::string::npos);
  EXPECT_NE(table.raw_passthrough_after_row_fields().find("t=\"grand\""), std::string::npos);
  EXPECT_NE(table.raw_col_items().find("<colItems"), std::string::npos);
  // They must NOT leak into the tail buffer.
  EXPECT_EQ(table.raw_passthrough_xml().find("<rowItems"), std::string::npos);

  // Write -> read again: the layout-item cache must still be present, and
  // rowItems must precede colItems in the emitted bytes.
  const std::string round = write_pivot_table_definition(table);
  const std::size_t p_rowitems = round.find("<rowItems");
  const std::size_t p_colitems = round.find("<colItems");
  ASSERT_NE(p_rowitems, std::string::npos) << round;
  ASSERT_NE(p_colitems, std::string::npos) << round;
  EXPECT_LT(p_rowitems, p_colitems);
  auto reparsed_or = read_pivot_table_definition(Bytes(round));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
  const pivot::PivotTable& reparsed = reparsed_or.value();
  EXPECT_NE(reparsed.raw_passthrough_after_row_fields().find("<rowItems"), std::string::npos);
  EXPECT_NE(reparsed.raw_col_items().find("<colItems"), std::string::npos);
}

// ---------------------------------------------------------------------------
// Report options, item captions / collapse state and the x14 field extension
// ---------------------------------------------------------------------------

constexpr std::string_view kX14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

std::string OptionsDefinition(std::string_view root_attrs, std::string_view field_xml) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\" dataCaption=\"Values\"");
  xml.append(root_attrs).append(">");
  xml.append("<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  xml.append("<pivotFields count=\"1\">").append(field_xml).append("</pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("</pivotTableDefinition>");
  return xml;
}

bool HasPassthrough(const std::vector<std::pair<std::string, std::string>>& attrs, std::string_view name) {
  for (const auto& attr : attrs) {
    if (attr.first == name) {
      return true;
    }
  }
  return false;
}

TEST(PivotTableReader, ReportOptionAttributesAreModelled) {
  const std::string xml = OptionsDefinition(" grandTotalCaption=\"TOTAL\" useAutoFormatting=\"1\" enableDrill=\"0\"",
                                            "<pivotField axis=\"axisRow\"/>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_TRUE(table.grand_total_caption().has_value());
  EXPECT_EQ(*table.grand_total_caption(), "TOTAL");
  EXPECT_TRUE(table.auto_formatting());
  EXPECT_FALSE(table.enable_drill());
  EXPECT_FALSE(HasPassthrough(table.passthrough_attrs(), "grandTotalCaption"));
  EXPECT_FALSE(HasPassthrough(table.passthrough_attrs(), "useAutoFormatting"));
  EXPECT_FALSE(HasPassthrough(table.passthrough_attrs(), "enableDrill"));
}

TEST(PivotTableReader, AbsentReportOptionsReadAsTheirSchemaDefaults) {
  // A table read without `useAutoFormatting` does not autofit, although a
  // table created in memory does.
  auto table_or = read_pivot_table_definition(Bytes(OptionsDefinition("", "<pivotField axis=\"axisRow\"/>")));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_FALSE(table_or.value().grand_total_caption().has_value());
  EXPECT_FALSE(table_or.value().auto_formatting());
  EXPECT_TRUE(table_or.value().enable_drill());
  EXPECT_TRUE(pivot::PivotTable().auto_formatting());
}

TEST(PivotTableReader, InsertBlankRowAndItemCaptionAndCollapseAreModelled) {
  const std::string xml = OptionsDefinition(
      "",
      "<pivotField axis=\"axisRow\" showAll=\"0\" insertBlankRow=\"1\"><items count=\"4\">"
      "<item sd=\"0\" x=\"2\"/><item n=\"zz\" x=\"0\" s=\"1\"/><item x=\"1\"/><item t=\"default\" sd=\"0\"/>"
      "</items></pivotField>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotField& field = table_or.value().fields()[0];
  EXPECT_TRUE(field.insert_blank_row);
  EXPECT_FALSE(HasPassthrough(field.passthrough_attrs, "insertBlankRow"));
  ASSERT_EQ(field.items.size(), 3U);
  EXPECT_FALSE(field.items[0].show_detail);
  EXPECT_TRUE(field.items[0].caption.empty());
  EXPECT_TRUE(field.items[1].show_detail);
  EXPECT_EQ(field.items[1].caption, "zz");
  EXPECT_FALSE(HasPassthrough(field.items[1].passthrough_attrs, "n"));
  EXPECT_FALSE(HasPassthrough(field.items[0].passthrough_attrs, "sd"));
  EXPECT_TRUE(HasPassthrough(field.items[1].passthrough_attrs, "s"));
  ASSERT_EQ(field.marker_items.size(), 1U);
  EXPECT_FALSE(field.marker_items[0].show_detail);
}

TEST(PivotTableReader, X14PivotFieldExtensionIsSplitAndOtherExtensionsKeptVerbatim) {
  std::string field_xml =
      "<pivotField axis=\"axisRow\" showAll=\"0\"><items count=\"1\"><item t=\"default\"/></items>"
      "<autoSortScope><pivotArea dataOnly=\"0\" outline=\"0\" fieldPosition=\"0\"/></autoSortScope><extLst>"
      "<ext uri=\"{AAAAAAAA-0000-0000-0000-000000000000}\"><other/></ext>"
      "<ext uri=\"{2946ED86-A175-432A-8AC1-64E0C546D7DE}\" xmlns:x14=\"";
  field_xml.append(kX14Ns).append(
      "\"><x14:pivotField fillDownLabels=\"1\" ignore=\"1\"><x14:child/></x14:pivotField></ext></extLst>"
      "</pivotField>");
  auto table_or = read_pivot_table_definition(Bytes(OptionsDefinition("", field_xml)));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotField& field = table_or.value().fields()[0];
  // The upper-case URI spelling is the same extension.
  EXPECT_TRUE(field.fill_down_labels);
  ASSERT_EQ(field.x14_attrs.size(), 1U);
  EXPECT_EQ(field.x14_attrs[0].first, "ignore");
  EXPECT_NE(field.x14_children_xml.find("<x14:child"), std::string::npos);
  EXPECT_NE(field.raw_ext_xml.find("{AAAAAAAA-0000-0000-0000-000000000000}"), std::string::npos);
  EXPECT_EQ(field.raw_ext_xml.find("2946ED86"), std::string::npos);
  EXPECT_NE(field.raw_children_xml.find("<autoSortScope"), std::string::npos);
  EXPECT_EQ(field.raw_children_xml.find("<extLst"), std::string::npos);
}

TEST(PivotTableReader, ColItemsAndPageFieldsLandInSeparateBins) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  xml.append("<pivotFields count=\"2\"><pivotField axis=\"axisCol\"/><pivotField axis=\"axisPage\"/></pivotFields>");
  xml.append("<colFields count=\"1\"><field x=\"0\"/></colFields>");
  xml.append("<colItems count=\"1\"><i/></colItems>");
  xml.append("<pageFields count=\"1\"><pageField fld=\"1\" hier=\"-1\"/></pageFields>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_NE(table_or.value().raw_col_items().find("<colItems"), std::string::npos);
  EXPECT_EQ(table_or.value().raw_col_items().find("<pageFields"), std::string::npos);
  EXPECT_NE(table_or.value().raw_passthrough_after_col_fields().find("<pageFields"), std::string::npos);
  EXPECT_EQ(table_or.value().raw_passthrough_after_col_fields().find("<colItems"), std::string::npos);
}

}  // namespace
}  // namespace formulon::io
