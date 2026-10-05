// Pivot-table definition reader tests grouped by XML concern.

#include "io/pivot_table_reader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_table_reader_test_helpers.h"
#include "utils/error.h"

namespace formulon::io {
namespace {
using namespace pivot_table_reader_test_support;
TEST(PivotTableReader, MinimalHappyPath) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P1\" cacheId=\"3\">");
  xml.append("  <location ref=\"A3:D10\"/>");
  xml.append("  <pivotFields count=\"0\"/>");
  xml.append("  <dataFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  EXPECT_EQ(table.name(), "P1");
  EXPECT_EQ(table.pivot_cache_id(), 3U);
  EXPECT_EQ(table.anchor_row(), 2U);
  EXPECT_EQ(table.anchor_col(), 0U);
  EXPECT_EQ(table.span_rows(), 8U);
  EXPECT_EQ(table.span_cols(), 4U);
  EXPECT_TRUE(table.fields().empty());
  EXPECT_TRUE(table.data_fields().empty());
}
TEST(PivotTableReader, MultipleFieldsAxisDecoding) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"3\">");
  xml.append("    <pivotField axis=\"axisRow\" name=\"R\"/>");
  xml.append("    <pivotField axis=\"axisCol\"/>");
  xml.append("    <pivotField dataField=\"1\"/>");  // unset axis -> Value
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 3U);
  EXPECT_EQ(table.fields()[0].axis, pivot::PivotAxis::Row);
  EXPECT_EQ(table.fields()[0].custom_name, "R");
  EXPECT_EQ(table.fields()[1].axis, pivot::PivotAxis::Col);
  EXPECT_EQ(table.fields()[2].axis, pivot::PivotAxis::Value);
}
TEST(PivotTableReader, ItemsDecodingSkipsSubtotalMarkers) {
  // Mix of real items (with `x` indices), a hidden item (`h="1"`), and a
  // subtotal marker (`t="default"`). The marker must be skipped; the
  // hidden item must round-trip with `visible = false`.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\">");
  xml.append("      <items count=\"4\">");
  xml.append("        <item x=\"0\"/>");
  xml.append("        <item h=\"1\" x=\"1\"/>");
  xml.append("        <item x=\"2\"/>");
  xml.append("        <item t=\"default\"/>");
  xml.append("      </items>");
  xml.append("    </pivotField>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  const auto& items = table.fields()[0].items;
  ASSERT_EQ(items.size(), 3U);
  EXPECT_TRUE(items[0].visible);
  EXPECT_FALSE(items[1].visible);
  EXPECT_TRUE(items[2].visible);
}
TEST(PivotTableReader, ItemWithMalformedCacheIndexIsRejected) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\"><items><item x=\"not-an-index\"/></items></pivotField>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_NE(table_or.error().message.find("invalid x attribute"), std::string::npos);
}
TEST(PivotTableReader, RowAndColFieldOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"4\">");
  xml.append("    <pivotField axis=\"axisRow\"/>");
  xml.append("    <pivotField axis=\"axisRow\"/>");
  xml.append("    <pivotField axis=\"axisCol\"/>");
  xml.append("    <pivotField/>");
  xml.append("  </pivotFields>");
  xml.append("  <rowFields count=\"2\">");
  xml.append("    <field x=\"0\"/>");
  xml.append("    <field x=\"1\"/>");
  xml.append("  </rowFields>");
  xml.append("  <colFields count=\"1\">");
  xml.append("    <field x=\"2\"/>");
  xml.append("  </colFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.row_field_order().size(), 2U);
  EXPECT_EQ(table.row_field_order()[0], 0U);
  EXPECT_EQ(table.row_field_order()[1], 1U);
  ASSERT_EQ(table.col_field_order().size(), 1U);
  EXPECT_EQ(table.col_field_order()[0], 2U);
}
TEST(PivotTableReader, SortTypeDescendingReversesTheField) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\" sortType=\"descending\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_FALSE(table.fields()[0].sort.ascending);
  EXPECT_FALSE(table.fields()[0].sort.manual);
}
TEST(PivotTableReader, SortTypeManualDefersToItemOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\" sortType=\"manual\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_TRUE(table.fields()[0].sort.manual);
}
TEST(PivotTableReader, MissingSortTypeKeepsAscendingDefault) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_TRUE(table.fields()[0].sort.ascending);
  EXPECT_FALSE(table.fields()[0].sort.manual);
}
TEST(PivotTableReader, ValuesFieldMarkerIsKeptOutOfFieldOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"2\">");
  xml.append("    <pivotField axis=\"axisCol\"/>");
  xml.append("    <pivotField dataField=\"1\"/>");
  xml.append("  </pivotFields>");
  xml.append("  <colFields count=\"2\">");
  xml.append("    <field x=\"0\"/>");
  xml.append("    <field x=\"-2\"/>");
  xml.append("  </colFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.col_field_order().size(), 1U);
  EXPECT_EQ(table.col_field_order()[0], 0U);
  ASSERT_TRUE(table.col_values_position().has_value());
  EXPECT_EQ(table.col_values_position().value(), 1U);
  EXPECT_FALSE(table.row_values_position().has_value());
}
TEST(PivotTableReader, MultipleDataFieldsOnSameSourceField) {
  // GETPIVOTDATA's display-name lookup requires that two data fields on
  // the same source column (Sum + Average of "Amount") survive as two
  // distinct PivotDataField entries.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"4\">");
  xml.append("    <pivotField/>");
  xml.append("    <pivotField/>");
  xml.append("    <pivotField/>");
  xml.append("    <pivotField dataField=\"1\"/>");
  xml.append("  </pivotFields>");
  xml.append("  <dataFields count=\"2\">");
  xml.append("    <dataField name=\"Sum of Amount\" fld=\"3\" subtotal=\"sum\"/>");
  xml.append("    <dataField name=\"Average of Amount\" fld=\"3\" subtotal=\"average\"/>");
  xml.append("  </dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.data_fields().size(), 2U);
  EXPECT_EQ(table.data_fields()[0].name, "Sum of Amount");
  EXPECT_EQ(table.data_fields()[0].field_index, 3U);
  EXPECT_EQ(table.data_fields()[0].aggregation, pivot::Aggregation::Sum);
  EXPECT_EQ(table.data_fields()[1].name, "Average of Amount");
  EXPECT_EQ(table.data_fields()[1].field_index, 3U);
  EXPECT_EQ(table.data_fields()[1].aggregation, pivot::Aggregation::Average);
}
TEST(PivotTableReader, DataFieldAggregationMappingExhaustive) {
  // One <dataField> per Aggregation enum value the OOXML subtotal
  // attribute can name.
  struct Case {
    const char* attr;
    pivot::Aggregation expected;
  };
  const Case cases[] = {
      {"sum", pivot::Aggregation::Sum},
      {"count", pivot::Aggregation::Count},
      {"average", pivot::Aggregation::Average},
      {"max", pivot::Aggregation::Max},
      {"min", pivot::Aggregation::Min},
      {"product", pivot::Aggregation::Product},
      {"countNums", pivot::Aggregation::CountNumbers},
      {"stdDev", pivot::Aggregation::StdDev},
      {"var", pivot::Aggregation::Var},
  };

  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <dataFields count=\"9\">");
  for (const Case& c : cases) {
    xml.append("    <dataField name=\"agg_")
        .append(c.attr)
        .append("\" fld=\"0\" subtotal=\"")
        .append(c.attr)
        .append("\"/>");
  }
  xml.append("  </dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.data_fields().size(), sizeof(cases) / sizeof(cases[0]));
  for (std::size_t i = 0; i < table.data_fields().size(); ++i) {
    EXPECT_EQ(table.data_fields()[i].aggregation, cases[i].expected) << "case=" << cases[i].attr;
  }
}
TEST(PivotTableReader, UnknownSubtotalFallsBackToSum) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <dataFields count=\"1\">");
  xml.append("    <dataField name=\"Weird of X\" fld=\"0\" subtotal=\"weirdo\"/>");
  xml.append("  </dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  ASSERT_EQ(table_or.value().data_fields().size(), 1U);
  EXPECT_EQ(table_or.value().data_fields()[0].aggregation, pivot::Aggregation::Sum);
}
TEST(PivotTableReader, PageFieldsCarryTheSelectedItem) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"7\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"1\"><pivotField axis=\"axisPage\"/></pivotFields>");
  xml.append("  <pageFields count=\"1\">");
  xml.append("    <pageField fld=\"0\" item=\"2\"/>");
  xml.append("  </pageFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  EXPECT_EQ(table.name(), "P");
  EXPECT_EQ(table.pivot_cache_id(), 7U);
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_EQ(table.fields()[0].axis, pivot::PivotAxis::Page);

  ASSERT_EQ(table.page_fields().size(), 1U);
  EXPECT_EQ(table.page_fields()[0].field_index, 0U);
  ASSERT_TRUE(table.page_fields()[0].item_index.has_value());
  EXPECT_EQ(*table.page_fields()[0].item_index, 2U);

  // Decoding does not take the block off the writer's hands: the authored
  // bytes still ride in the pre-dataFields bin so the round trip re-emits
  // them untouched.
  EXPECT_NE(table.raw_passthrough_after_col_fields().find("<pageField "), std::string::npos);
}
TEST(PivotTableReader, PageFieldWithoutASelectionLeavesTheItemAbsent) {
  // The unfiltered state: Excel writes no `item`, and the field is showing
  // every item rather than item 0.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"1\"><pivotField axis=\"axisPage\"/></pivotFields>");
  xml.append("  <pageFields count=\"1\"><pageField fld=\"0\" hier=\"-1\"/></pageFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  ASSERT_EQ(table_or.value().page_fields().size(), 1U);
  EXPECT_FALSE(table_or.value().page_fields()[0].item_index.has_value());
}
TEST(PivotTableReader, PageFieldWithoutAFieldIndexIsSkipped) {
  // `fld` names the field; absent, the entry designates nothing. Defaulting
  // it to 0 would draw a page header over a field the user never filtered by.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"2\"><pivotField axis=\"axisPage\"/><pivotField axis=\"axisRow\"/></pivotFields>");
  xml.append("  <pageFields count=\"2\"><pageField item=\"1\"/><pageField fld=\"0\"/></pageFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  ASSERT_EQ(table_or.value().page_fields().size(), 1U);
  EXPECT_EQ(table_or.value().page_fields()[0].field_index, 0U);
}
TEST(PivotTableReader, PageFieldOrderOverridesDocumentOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"2\"><pivotField axis=\"axisPage\"/><pivotField axis=\"axisPage\"/></pivotFields>");
  xml.append("  <pageFields count=\"2\"><pageField fld=\"1\"/><pageField fld=\"0\"/></pageFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_EQ(table_or.value().page_field_order(), (std::vector<std::uint32_t>{1, 0}));
}
TEST(PivotTableReader, PageAxisWithoutAPageFieldsBlockFallsBackToDocumentOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"3\">");
  xml.append("    <pivotField axis=\"axisRow\"/><pivotField axis=\"axisPage\"/><pivotField axis=\"axisPage\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_TRUE(table_or.value().page_fields().empty());
  EXPECT_EQ(table_or.value().page_field_order(), (std::vector<std::uint32_t>{1, 2}));
}

}  // namespace
}  // namespace formulon::io
