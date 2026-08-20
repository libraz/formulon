//
// Unit tests for `formulon::io::read_pivot_table_definition`. Each test
// feeds a hand-rolled OOXML byte vector into the reader and asserts the
// populated `pivot::PivotTable` shape; the workbook integration (rels
// resolution, xlsx fixture loading) is covered separately.

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
#include "utils/error.h"

namespace formulon::io {
namespace {

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

std::size_t CountOccurrences(std::string_view haystack, std::string_view needle) {
  std::size_t count = 0;
  for (std::size_t at = haystack.find(needle); at != std::string_view::npos; at = haystack.find(needle, at + 1)) {
    ++count;
  }
  return count;
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

// `<filters>` is decoded for evaluation but serialised by nobody: the block
// has to survive a read -> write round trip byte-for-byte and exactly once
// out of the passthrough tail, while the decoded view lands in
// `authored_caption_filters()`. It must still not appear in
// `active_filters()`, which is a session-only list the writer never emits
// (see `PivotTable::active_filters`).
TEST(PivotTableReader, AuthoredFiltersDecodeForEvaluationAndStillRoundTripAsPassthrough) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("<pivotTableStyleInfo name=\"PivotStyleLight16\"/>");
  xml.append(
      "<filters count=\"1\"><filter fld=\"0\" type=\"captionEqual\" evalOrder=\"-1\" id=\"1\">"
      "<autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"North\"/></customFilters>"
      "</filterColumn></autoFilter></filter></filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  EXPECT_TRUE(table.active_filters().empty());
  ASSERT_EQ(table.authored_caption_filters().size(), 1U);
  EXPECT_EQ(table.authored_caption_filters()[0].field_index, 0U);
  EXPECT_EQ(table.authored_caption_filters()[0].predicate, pivot::CaptionPredicate::Equal);
  EXPECT_EQ(table.authored_caption_filters()[0].value, "North");
  EXPECT_NE(table.raw_passthrough_xml().find("<filters"), std::string::npos);
  EXPECT_NE(table.raw_passthrough_xml().find("val=\"North\""), std::string::npos);

  const std::string round = write_pivot_table_definition(table);
  EXPECT_EQ(CountOccurrences(round, "<filters"), 1U) << round;
  EXPECT_NE(round.find("val=\"North\""), std::string::npos) << round;
  // The style block is authored before `<filters>`; re-emission keeps that
  // order, which the schema requires.
  EXPECT_LT(round.find("<pivotTableStyleInfo"), round.find("<filters"));

  auto reparsed_or = read_pivot_table_definition(Bytes(round));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
  EXPECT_TRUE(reparsed_or.value().active_filters().empty());
  // The decode is stable across the round trip, and the writer still emits
  // the block from the passthrough tail rather than from the decoded view.
  EXPECT_EQ(reparsed_or.value().authored_caption_filters().size(), 1U);
  EXPECT_EQ(write_pivot_table_definition(reparsed_or.value()), round);
}

// Excel repeats the predicate inside the criterion as an `autoFilter`
// wildcard pattern, so the decode has to strip the marker the `type`
// already names. `<top10>`-driven and value-family entries carry no
// caption criterion and are skipped rather than decoded into a wrong
// predicate.
TEST(PivotTableReader, AuthoredCaptionFilterVariantsDecodeAndNonCaptionTypesAreSkipped) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"5\">");
  xml.append(
      "<filter fld=\"1\" type=\"captionBeginsWith\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"No*\"/></customFilters>"
      "</filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"2\" type=\"captionContains\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"*or*\"/></customFilters>"
      "</filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"3\" type=\"captionEndsWith\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters><customFilter operator=\"equal\" val=\"*th\"/></customFilters>"
      "</filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"4\" type=\"captionBetween\"><autoFilter ref=\"A3:A9\"><filterColumn colId=\"0\">"
      "<customFilters and=\"1\"><customFilter operator=\"greaterThanOrEqual\" val=\"B\"/>"
      "<customFilter operator=\"lessThanOrEqual\" val=\"M\"/></customFilters>"
      "</filterColumn></autoFilter></filter>");
  // Value family: outside the caption set, so passthrough-only.
  xml.append(
      "<filter fld=\"5\" type=\"valueGreaterThan\" iMeasureFld=\"0\"><autoFilter ref=\"A3:A9\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter operator=\"greaterThan\" val=\"100\"/>"
      "</customFilters></filterColumn></autoFilter></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& decoded = table_or.value().authored_caption_filters();
  ASSERT_EQ(decoded.size(), 4U);
  EXPECT_EQ(decoded[0].field_index, 1U);
  EXPECT_EQ(decoded[0].predicate, pivot::CaptionPredicate::BeginsWith);
  EXPECT_EQ(decoded[0].value, "No");
  EXPECT_EQ(decoded[1].predicate, pivot::CaptionPredicate::Contains);
  EXPECT_EQ(decoded[1].value, "or");
  EXPECT_EQ(decoded[2].predicate, pivot::CaptionPredicate::EndsWith);
  EXPECT_EQ(decoded[2].value, "th");
  EXPECT_EQ(decoded[3].predicate, pivot::CaptionPredicate::Between);
  EXPECT_EQ(decoded[3].value, "B");
  EXPECT_EQ(decoded[3].value_high, "M");
  // Every entry, decoded or not, is still re-emitted from the passthrough.
  const std::string round = write_pivot_table_definition(table_or.value());
  EXPECT_NE(round.find("valueGreaterThan"), std::string::npos) << round;
  EXPECT_EQ(CountOccurrences(round, "<filter "), 5U) << round;
}

// A plain equality is also spelled as a `<filters>` list of `<filter val>`
// rather than a `<customFilters>` comparison; both reach the same decode.
TEST(PivotTableReader, AuthoredCaptionFilterAcceptsThePlainFilterListShape) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append(
      "<filters count=\"1\"><filter fld=\"0\" type=\"captionEqual\"><autoFilter ref=\"A3:A9\">"
      "<filterColumn colId=\"0\"><filters><filter val=\"South\"/></filters>"
      "</filterColumn></autoFilter></filter></filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  ASSERT_EQ(table_or.value().authored_caption_filters().size(), 1U);
  EXPECT_EQ(table_or.value().authored_caption_filters()[0].value, "South");
}

// Entries the value decoder must refuse rather than answer approximately,
// all of which still have to survive the round trip as passthrough.
//
// `percent` and `sum` are the other two flavours of Excel's "top ten"
// dialog. They share `count`'s nested `<top10>` element but rank on a
// share of the total rather than on an item count, so reading them as a
// count would quietly return the wrong rows -- worse than not filtering.
// An entry that names a family but no usable criterion is the remaining
// unevaluable shape: it must not reach any decoded list, and it must
// still survive the round trip so a save does not drop a filter Excel
// authored and can apply itself.
TEST(PivotTableReader, UnevaluableFilterFamiliesAreSkippedButStillRoundTrip) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"3\">");
  // A count filter with no nested `<top10>` names no quantity at all.
  xml.append("<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"/></filter>");
  // An unknown type is the forward-compatibility case: a family this
  // version does not know must be passed through rather than guessed at.
  xml.append("<filter fld=\"1\" type=\"someFutureFamily\"><autoFilter ref=\"A1\"/></filter>");
  // A caption filter whose criterion element is missing entirely.
  xml.append("<filter fld=\"2\" type=\"captionEqual\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());
  EXPECT_TRUE(table_or.value().authored_caption_filters().empty());
  EXPECT_TRUE(table_or.value().authored_period_filters().empty());
  EXPECT_TRUE(table_or.value().authored_recurring_filters().empty());

  const std::string round = write_pivot_table_definition(table_or.value());
  EXPECT_EQ(CountOccurrences(round, "<filter "), 3U) << round;
  EXPECT_NE(round.find("someFutureFamily"), std::string::npos) << round;
}

TEST(PivotTableReader, DecodedFilterFamiliesStillRoundTripVerbatim) {
  // Decoding is read-side only: the writer re-emits the `<filters>` block
  // from the verbatim passthrough tail, so gaining a decoder must not
  // change a single byte of what is written back.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"2\">");
  xml.append("<filter fld=\"2\" type=\"thisWeek\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"2\" type=\"M1\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_EQ(table_or.value().authored_period_filters().size(), 1U);
  EXPECT_EQ(table_or.value().authored_recurring_filters().size(), 1U);

  const std::string round = write_pivot_table_definition(table_or.value());
  EXPECT_EQ(CountOccurrences(round, "<filter "), 2U) << round;
  EXPECT_NE(round.find("thisWeek"), std::string::npos) << round;
  EXPECT_NE(round.find("\"M1\""), std::string::npos) << round;
}

TEST(PivotTableReader, TopNFlavoursAreDistinguishedByTheirTypeAlone) {
  // The three flavours share an identically shaped `<top10 val="N">`, so a
  // decode that ignored the type attribute would silently conflate them.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"3\">");
  xml.append(
      "<filter fld=\"0\" type=\"count\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 val=\"2\"/></filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"0\" type=\"percent\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 percent=\"1\" val=\"70\"/></filterColumn></autoFilter></filter>");
  xml.append(
      "<filter fld=\"0\" type=\"sum\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
      "<top10 val=\"100\"/></filterColumn></autoFilter></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& filters = table_or.value().authored_value_filters();
  ASSERT_EQ(filters.size(), 3U);
  for (const auto& entry : filters) {
    EXPECT_EQ(entry.type, pivot::FilterType::ValueTop10);
  }
  EXPECT_EQ(filters[0].top_n_basis, pivot::TopNBasis::Items);
  EXPECT_DOUBLE_EQ(filters[0].value, 2.0);
  EXPECT_EQ(filters[1].top_n_basis, pivot::TopNBasis::Percent);
  EXPECT_DOUBLE_EQ(filters[1].value, 70.0);
  EXPECT_EQ(filters[2].top_n_basis, pivot::TopNBasis::Sum);
  EXPECT_DOUBLE_EQ(filters[2].value, 100.0);
}

TEST(PivotTableReader, RelativePeriodFiltersDecodeFromTheTypeNameAlone) {
  // These carry no criteria: the window is implied by the type, so the
  // decode is the name mapping and nothing else.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"3\">");
  xml.append("<filter fld=\"2\" type=\"thisMonth\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"3\" type=\"yearToDate\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"1\" type=\"lastQuarter\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& periods = table_or.value().authored_period_filters();
  ASSERT_EQ(periods.size(), 3U);
  EXPECT_EQ(periods[0].field_index, 2U);
  EXPECT_EQ(periods[0].period, pivot::RelativePeriod::ThisMonth);
  EXPECT_EQ(periods[1].field_index, 3U);
  EXPECT_EQ(periods[1].period, pivot::RelativePeriod::YearToDate);
  EXPECT_EQ(periods[2].field_index, 1U);
  EXPECT_EQ(periods[2].period, pivot::RelativePeriod::LastQuarter);
  // A period entry carries no bounds, so it must not land in the list whose
  // members all do.
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());

  const std::string round = write_pivot_table_definition(table_or.value());
  EXPECT_EQ(CountOccurrences(round, "<filter "), 3U) << round;
}

TEST(PivotTableReader, RecurringPeriodFiltersDecodeToACalendarMonthRange) {
  // `M<n>` is one month and `Q<n>` its three, both year-free -- so unlike
  // the relative periods these need no clock to mean something.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"4\">");
  xml.append("<filter fld=\"2\" type=\"M1\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"2\" type=\"M12\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"3\" type=\"Q1\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"1\" type=\"Q4\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& recurring = table_or.value().authored_recurring_filters();
  ASSERT_EQ(recurring.size(), 4U);
  EXPECT_EQ(recurring[0].field_index, 2U);
  EXPECT_EQ(recurring[0].month_low, 1U);
  EXPECT_EQ(recurring[0].month_high, 1U);
  EXPECT_EQ(recurring[1].month_low, 12U);
  EXPECT_EQ(recurring[1].month_high, 12U);
  EXPECT_EQ(recurring[2].field_index, 3U);
  EXPECT_EQ(recurring[2].month_low, 1U);
  EXPECT_EQ(recurring[2].month_high, 3U);
  EXPECT_EQ(recurring[3].month_low, 10U);
  EXPECT_EQ(recurring[3].month_high, 12U);
  // A recurring entry is not a window and carries no bounds, so it must
  // reach neither sibling list.
  EXPECT_TRUE(table_or.value().authored_period_filters().empty());
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());
}

TEST(PivotTableReader, MalformedRecurringSelectorsAreSkipped) {
  // `M0`, `M13` and `Q5` are outside their families, and a bare letter or
  // a non-digit tail names nothing at all. None may decode to a range that
  // silently filters the wrong months.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"6\">");
  for (const char* type : {"M0", "M13", "Q0", "Q5", "M", "Mx"}) {
    xml.append("<filter fld=\"2\" type=\"").append(type).append("\"><autoFilter ref=\"A1\"/></filter>");
  }
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().authored_recurring_filters().empty());
}

TEST(PivotTableReader, WeekPeriodFiltersDecodeAlongsideTheOtherWindows) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<filters count=\"3\">");
  xml.append("<filter fld=\"2\" type=\"thisWeek\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"2\" type=\"lastWeek\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("<filter fld=\"2\" type=\"nextWeek\"><autoFilter ref=\"A1\"/></filter>");
  xml.append("</filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const auto& periods = table_or.value().authored_period_filters();
  ASSERT_EQ(periods.size(), 3U);
  EXPECT_EQ(periods[0].period, pivot::RelativePeriod::ThisWeek);
  EXPECT_EQ(periods[1].period, pivot::RelativePeriod::LastWeek);
  EXPECT_EQ(periods[2].period, pivot::RelativePeriod::NextWeek);
}

TEST(PivotTableReader, ValueFilterCriteriaThatAreNotNumbersAreSkipped) {
  // The criterion shares the sheet path's number lexer, so a caption-ish
  // payload under a value type is refused instead of reaching the
  // comparison as zero.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append(
      "<filters count=\"1\"><filter fld=\"0\" type=\"valueGreaterThan\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters><customFilter operator=\"greaterThan\" val=\"North\"/>"
      "</customFilters></filterColumn></autoFilter></filter></filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());
}

TEST(PivotTableReader, ValueBetweenWithOnlyOneBoundIsSkipped) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append(
      "<filters count=\"1\"><filter fld=\"0\" type=\"valueBetween\"><autoFilter ref=\"A1\">"
      "<filterColumn colId=\"0\"><customFilters and=\"1\">"
      "<customFilter operator=\"greaterThanOrEqual\" val=\"100\"/>"
      "</customFilters></filterColumn></autoFilter></filter></filters>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  EXPECT_TRUE(table_or.value().authored_value_filters().empty());
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
