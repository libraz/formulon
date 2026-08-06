//
// Unit tests for `formulon::io::read_pivot_table_definition`. Each test
// feeds a hand-rolled OOXML byte vector into the reader and asserts the
// populated `pivot::PivotTable` shape; the workbook integration (rels
// resolution, xlsx fixture loading) is covered separately.

#include "io/pivot_table_reader.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"

namespace formulon::io {
namespace {

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kPivotNs = " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

// ---------------------------------------------------------------------------
// Happy paths
// ---------------------------------------------------------------------------

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

TEST(PivotTableReader, PageFieldsSilentlySkipped) {
  // `<pageFields>` carries `<pageField>` entries we do not yet model.
  // The reader must not fail and must not surface anything from them.
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
}

// ---------------------------------------------------------------------------
// Error paths
// ---------------------------------------------------------------------------

TEST(PivotTableReader, MissingDataFieldNameIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <dataFields count=\"1\">");
  xml.append("    <dataField fld=\"1\" subtotal=\"sum\"/>");
  xml.append("  </dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(PivotTableReader, MissingLocationIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(PivotTableReader, UnparseableLocationRefIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"not-a-range\"/>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(PivotTableReader, MalformedRootIsContentTypeInvalid) {
  std::string xml(kXmlDecl);
  xml.append("<foo").append(kPivotNs).append("><location ref=\"A1\"/></foo>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(PivotTableReader, MalformedXmlIsParseError) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append("><location");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXmlParse);
}

// ---------------------------------------------------------------------------
// Passthrough capture for unmodelled extensions.
// ---------------------------------------------------------------------------

TEST(PivotTableReader, UnmodelledChildrenCapturedAsPassthrough) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<calculatedItems count=\"1\"><calculatedItem name=\"Avg\" formula=\"=A1/B1\"/></calculatedItems>");
  xml.append("<pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\"/>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const std::string& passthrough = table_or.value().raw_passthrough_xml();
  // Both unrecognised elements survive verbatim (order-preserving).
  EXPECT_NE(passthrough.find("<calculatedItems"), std::string::npos);
  EXPECT_NE(passthrough.find("formula=\"=A1/B1\""), std::string::npos);
  EXPECT_NE(passthrough.find("<pivotTableStyleInfo"), std::string::npos);
  EXPECT_NE(passthrough.find("PivotStyleLight16"), std::string::npos);
  // Recognised elements (location) are NOT in the passthrough.
  EXPECT_EQ(passthrough.find("<location"), std::string::npos);
}

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

// ---------------------------------------------------------------------------
// Per-field subtotal selection / defaultSubtotal round trip.
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// <rowItems> / <colItems> survive verbatim through the passthrough buffer.
// ---------------------------------------------------------------------------

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
  EXPECT_NE(table.raw_passthrough_after_col_fields().find("<colItems"), std::string::npos);
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
  EXPECT_NE(reparsed.raw_passthrough_after_col_fields().find("<colItems"), std::string::npos);
}

}  // namespace
}  // namespace formulon::io
