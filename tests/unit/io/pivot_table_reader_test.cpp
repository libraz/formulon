// Copyright 2026 libraz. Licensed under the MIT License.
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

}  // namespace
}  // namespace formulon::io
