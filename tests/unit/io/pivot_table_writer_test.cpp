// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for `formulon::io::write_pivot_table_definition`. Each test
// builds a hand-rolled `pivot::PivotTable`, writes the XML, and (where
// useful) pipes the bytes back through the symmetric reader to assert a
// clean round-trip on the data the reader actually consumes.

#include "io/pivot_table_writer.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_table_reader.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"

namespace formulon::io {
namespace {

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kPivotNs = " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

// ---------------------------------------------------------------------------
// Empty table
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, EmptyTableRoundTrips) {
  pivot::PivotTable empty;
  const std::string xml = write_pivot_table_definition(empty);
  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  EXPECT_TRUE(parsed.fields().empty());
  EXPECT_TRUE(parsed.row_field_order().empty());
  EXPECT_TRUE(parsed.col_field_order().empty());
  EXPECT_TRUE(parsed.data_fields().empty());
  // Default-constructed PivotTable has spans 0/0; the writer's
  // EncodeA1Range zero-span guard collapses that to A1:A1, which the
  // reader interprets as a 1x1 anchor at (0,0).
  EXPECT_EQ(parsed.anchor_row(), 0U);
  EXPECT_EQ(parsed.anchor_col(), 0U);
}

// ---------------------------------------------------------------------------
// Definition: round-trip via reader, mirroring the integration test
// fixture in tests/integration/ooxml_pivot_test.cpp.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, DefinitionRoundTripsThroughReader) {
  pivot::PivotTable table;
  table.set_name("PivotTable1");
  table.set_pivot_cache_id(0U);
  table.set_anchor(0U, 3U, 5U, 2U);

  // Field 0: Region on the row axis, two visible items.
  pivot::PivotField region;
  region.axis = pivot::PivotAxis::Row;
  region.custom_name = "Region";
  region.items.push_back(pivot::PivotItem{"", true});
  region.items.push_back(pivot::PivotItem{"", true});
  table.mutable_fields().push_back(std::move(region));

  // Field 1: Amount on the value axis (no items; data only).
  pivot::PivotField amount;
  amount.axis = pivot::PivotAxis::Value;
  amount.custom_name = "Amount";
  table.mutable_fields().push_back(std::move(amount));

  table.mutable_row_field_order().push_back(0U);

  pivot::PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1U;
  sum_amount.aggregation = pivot::Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));

  const std::string xml = write_pivot_table_definition(table);
  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();

  EXPECT_EQ(parsed.name(), "PivotTable1");
  EXPECT_EQ(parsed.pivot_cache_id(), 0U);
  EXPECT_EQ(parsed.anchor_row(), 0U);
  EXPECT_EQ(parsed.anchor_col(), 3U);
  EXPECT_EQ(parsed.span_rows(), 5U);
  EXPECT_EQ(parsed.span_cols(), 2U);

  ASSERT_EQ(parsed.fields().size(), 2U);
  EXPECT_EQ(parsed.fields()[0].axis, pivot::PivotAxis::Row);
  EXPECT_EQ(parsed.fields()[0].custom_name, "Region");
  ASSERT_EQ(parsed.fields()[0].items.size(), 2U);
  EXPECT_TRUE(parsed.fields()[0].items[0].visible);
  EXPECT_TRUE(parsed.fields()[0].items[1].visible);

  EXPECT_EQ(parsed.fields()[1].axis, pivot::PivotAxis::Value);
  EXPECT_EQ(parsed.fields()[1].custom_name, "Amount");

  ASSERT_EQ(parsed.row_field_order().size(), 1U);
  EXPECT_EQ(parsed.row_field_order()[0], 0U);
  EXPECT_TRUE(parsed.col_field_order().empty());

  ASSERT_EQ(parsed.data_fields().size(), 1U);
  EXPECT_EQ(parsed.data_fields()[0].name, "Sum of Amount");
  EXPECT_EQ(parsed.data_fields()[0].field_index, 1U);
  EXPECT_EQ(parsed.data_fields()[0].aggregation, pivot::Aggregation::Sum);
}

// ---------------------------------------------------------------------------
// Anchor encoding: range form, multi-letter columns
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, AnchorRangeEncodesAsA1) {
  pivot::PivotTable table;
  // (row=5, col=1, span_rows=3, span_cols=4) -> top-left B6, bottom-right E8.
  table.set_anchor(5U, 1U, 3U, 4U);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("ref=\"B6:E8\""), std::string::npos) << "xml=" << xml;
}

// ---------------------------------------------------------------------------
// All four axes round-trip through reader+writer.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, AllAxesRoundTrip) {
  pivot::PivotTable table;
  table.set_name("A");
  table.set_anchor(0U, 0U, 1U, 1U);
  for (const pivot::PivotAxis ax :
       {pivot::PivotAxis::Row, pivot::PivotAxis::Col, pivot::PivotAxis::Value, pivot::PivotAxis::Page}) {
    pivot::PivotField f;
    f.axis = ax;
    f.custom_name = "F";
    table.mutable_fields().push_back(std::move(f));
  }

  const std::string xml = write_pivot_table_definition(table);
  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_EQ(parsed.fields().size(), 4U);
  EXPECT_EQ(parsed.fields()[0].axis, pivot::PivotAxis::Row);
  EXPECT_EQ(parsed.fields()[1].axis, pivot::PivotAxis::Col);
  EXPECT_EQ(parsed.fields()[2].axis, pivot::PivotAxis::Value);
  EXPECT_EQ(parsed.fields()[3].axis, pivot::PivotAxis::Page);
}

// ---------------------------------------------------------------------------
// Hidden item visibility round-trips via h="1".
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, HiddenItemRoundTrips) {
  pivot::PivotTable table;
  table.set_name("T");
  table.set_anchor(0U, 0U, 1U, 1U);
  pivot::PivotField f;
  f.axis = pivot::PivotAxis::Row;
  f.custom_name = "Region";
  f.items.push_back(pivot::PivotItem{"", true});
  f.items.push_back(pivot::PivotItem{"", false});  // hidden
  table.mutable_fields().push_back(std::move(f));

  const std::string xml = write_pivot_table_definition(table);
  // Sanity: the second item carries h="1".
  EXPECT_NE(xml.find("<item x=\"1\" h=\"1\"/>"), std::string::npos) << "xml=" << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_EQ(parsed.fields().size(), 1U);
  ASSERT_EQ(parsed.fields()[0].items.size(), 2U);
  EXPECT_TRUE(parsed.fields()[0].items[0].visible);
  EXPECT_FALSE(parsed.fields()[0].items[1].visible);
}

// ---------------------------------------------------------------------------
// Every Aggregation variant round-trips through the reader's mapping.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, AllAggregationsRoundTrip) {
  pivot::PivotTable table;
  table.set_name("T");
  table.set_anchor(0U, 0U, 1U, 1U);
  // One pivot field per data-field entry; reader does not enforce the
  // shape but we keep the indices consistent.
  pivot::PivotField placeholder;
  placeholder.axis = pivot::PivotAxis::Value;
  placeholder.custom_name = "X";
  table.mutable_fields().push_back(std::move(placeholder));

  const std::vector<pivot::Aggregation> aggs = {
      pivot::Aggregation::Sum,          pivot::Aggregation::Count,  pivot::Aggregation::Average,
      pivot::Aggregation::Max,          pivot::Aggregation::Min,    pivot::Aggregation::Product,
      pivot::Aggregation::CountNumbers, pivot::Aggregation::StdDev, pivot::Aggregation::StdDevP,
      pivot::Aggregation::Var,          pivot::Aggregation::VarP,
  };
  for (std::size_t i = 0; i < aggs.size(); ++i) {
    pivot::PivotDataField df;
    df.name = "DF" + std::to_string(i);
    df.field_index = 0U;
    df.aggregation = aggs[i];
    table.mutable_data_fields().push_back(std::move(df));
  }

  const std::string xml = write_pivot_table_definition(table);
  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_EQ(parsed.data_fields().size(), aggs.size());
  for (std::size_t i = 0; i < aggs.size(); ++i) {
    EXPECT_EQ(parsed.data_fields()[i].aggregation, aggs[i]) << "i=" << i;
  }
}

// ---------------------------------------------------------------------------
// XML escaping in pivot-field name and data-field name.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, XmlEscapingApplies) {
  pivot::PivotTable table;
  table.set_name("T");
  table.set_anchor(0U, 0U, 1U, 1U);
  pivot::PivotField f;
  f.axis = pivot::PivotAxis::Row;
  f.custom_name = "<R&D>";
  table.mutable_fields().push_back(std::move(f));

  pivot::PivotDataField df;
  df.name = "Sum & Average of \"X\"";
  df.field_index = 0U;
  df.aggregation = pivot::Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(df));

  const std::string xml = write_pivot_table_definition(table);
  // Sanity: escaped sequences are present in the bytes.
  EXPECT_NE(xml.find("&lt;R&amp;D&gt;"), std::string::npos);
  EXPECT_NE(xml.find("Sum &amp; Average of &quot;X&quot;"), std::string::npos);

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_EQ(parsed.fields().size(), 1U);
  EXPECT_EQ(parsed.fields()[0].custom_name, "<R&D>");
  ASSERT_EQ(parsed.data_fields().size(), 1U);
  EXPECT_EQ(parsed.data_fields()[0].name, "Sum & Average of \"X\"");
}

// ---------------------------------------------------------------------------
// Empty optional elements: rowFields/colFields/dataFields/items omitted
// when the corresponding source vector is empty.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, EmptyOptionalElementsAreOmitted) {
  pivot::PivotTable table;
  table.set_name("T");
  table.set_anchor(0U, 0U, 1U, 1U);
  // No fields, no orders, no data fields, no items.
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_EQ(xml.find("<rowFields"), std::string::npos);
  EXPECT_EQ(xml.find("<colFields"), std::string::npos);
  EXPECT_EQ(xml.find("<dataFields"), std::string::npos);
  EXPECT_EQ(xml.find("<items"), std::string::npos);
}

// ---------------------------------------------------------------------------
// Round-trip preservation of unmodelled OOXML extensions
// (calculatedItems / pivotTableStyleInfo / etc.)
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, PassthroughRoundTrips) {
  std::string xml(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
      " name=\"P\" cacheId=\"7\">"
      "<location ref=\"A1:B2\"/>"
      "<calculatedItems count=\"1\"><calculatedItem name=\"Avg\" formula=\"=A1/B1\"/></calculatedItems>"
      "<pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\"/>"
      "</pivotTableDefinition>");
  auto first_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(first_or)) << first_or.error().message;
  const std::string written = write_pivot_table_definition(first_or.value());

  // The unmodelled extensions must reappear verbatim in the writer output.
  EXPECT_NE(written.find("<calculatedItems"), std::string::npos) << "written=" << written;
  EXPECT_NE(written.find("formula=\"=A1/B1\""), std::string::npos);
  EXPECT_NE(written.find("<pivotTableStyleInfo"), std::string::npos);
  EXPECT_NE(written.find("PivotStyleLight16"), std::string::npos);

  // A second round trip must remain stable (idempotency).
  auto second_or = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(second_or)) << second_or.error().message;
  EXPECT_EQ(write_pivot_table_definition(second_or.value()), written);
}

// ---------------------------------------------------------------------------
// showDataAs / baseField / baseItem round-trip.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, ShowDataAsRoundTrips) {
  pivot::PivotTable table;
  table.set_name("T");
  table.set_anchor(0U, 0U, 1U, 1U);
  pivot::PivotField placeholder;
  placeholder.axis = pivot::PivotAxis::Value;
  placeholder.custom_name = "X";
  table.mutable_fields().push_back(std::move(placeholder));

  pivot::PivotDataField df;
  df.name = "PoP";
  df.field_index = 0U;
  df.aggregation = pivot::Aggregation::Sum;
  df.show_as = pivot::ShowValuesAs::PercentOfParentCol;
  df.show_as_base_field = 2U;
  df.show_as_base_item = pivot::kShowAsBasePrev;
  table.mutable_data_fields().push_back(std::move(df));

  const std::string xml = write_pivot_table_definition(table);
  // Sanity: emitted attributes appear in the written bytes.
  EXPECT_NE(xml.find("showDataAs=\"percentOfParentCol\""), std::string::npos) << "xml=" << xml;
  EXPECT_NE(xml.find("baseField=\"2\""), std::string::npos);
  EXPECT_NE(xml.find("baseItem=\"1048828\""), std::string::npos);

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_EQ(parsed.data_fields().size(), 1U);
  const pivot::PivotDataField& got = parsed.data_fields()[0];
  EXPECT_EQ(got.show_as, pivot::ShowValuesAs::PercentOfParentCol);
  ASSERT_TRUE(got.show_as_base_field.has_value());
  EXPECT_EQ(*got.show_as_base_field, 2U);
  ASSERT_TRUE(got.show_as_base_item.has_value());
  EXPECT_EQ(*got.show_as_base_item, pivot::kShowAsBasePrev);
}

// ---------------------------------------------------------------------------
// Grand-total flags and <location> required attributes survive a
// read -> write -> read round trip driven from a real <pivotTableDefinition>
// fixture (the production load path), instead of silently flipping defaults
// or dropping schema-required attributes.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, GrandTotalsAndLocationAttrsSurviveRoundTrip) {
  // Fixture: grand totals explicitly OFF, full <location> attribute set.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs);
  xml.append(" name=\"P\" cacheId=\"0\" rowGrandTotals=\"0\" colGrandTotals=\"0\">");
  xml.append(
      "<location ref=\"A3:D10\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\" "
      "rowPageCount=\"1\" colPageCount=\"2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");

  // First load through the production reader.
  auto first_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(first_or)) << "read failed: " << first_or.error().message;
  const pivot::PivotTable& first = first_or.value();
  // (a) Grand totals stay OFF (default is true; must not flip back).
  EXPECT_FALSE(first.grand_totals_rows());
  EXPECT_FALSE(first.grand_totals_cols());
  // (b) <location> required + optional attributes captured.
  ASSERT_TRUE(first.location_first_header_row().has_value());
  EXPECT_EQ(*first.location_first_header_row(), 1U);
  ASSERT_TRUE(first.location_first_data_row().has_value());
  EXPECT_EQ(*first.location_first_data_row(), 2U);
  ASSERT_TRUE(first.location_first_data_col().has_value());
  EXPECT_EQ(*first.location_first_data_col(), 1U);
  ASSERT_TRUE(first.location_row_page_count().has_value());
  EXPECT_EQ(*first.location_row_page_count(), 1U);
  ASSERT_TRUE(first.location_col_page_count().has_value());
  EXPECT_EQ(*first.location_col_page_count(), 2U);

  // Write back out: the grand-total flags and location attributes must be
  // re-emitted (the flags only when OFF; the location offsets verbatim).
  const std::string written = write_pivot_table_definition(first);
  EXPECT_NE(written.find("rowGrandTotals=\"0\""), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("colGrandTotals=\"0\""), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("firstHeaderRow=\"1\""), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("firstDataRow=\"2\""), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("firstDataCol=\"1\""), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("rowPageCount=\"1\""), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("colPageCount=\"2\""), std::string::npos) << "xml=" << written;

  // Second load: same observable state survives the full round trip.
  auto second_or = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(second_or)) << "read failed: " << second_or.error().message;
  const pivot::PivotTable& second = second_or.value();
  EXPECT_FALSE(second.grand_totals_rows());
  EXPECT_FALSE(second.grand_totals_cols());
  ASSERT_TRUE(second.location_first_header_row().has_value());
  EXPECT_EQ(*second.location_first_header_row(), 1U);
  ASSERT_TRUE(second.location_first_data_row().has_value());
  EXPECT_EQ(*second.location_first_data_row(), 2U);
  ASSERT_TRUE(second.location_first_data_col().has_value());
  EXPECT_EQ(*second.location_first_data_col(), 1U);
  ASSERT_TRUE(second.location_row_page_count().has_value());
  EXPECT_EQ(*second.location_row_page_count(), 1U);
  ASSERT_TRUE(second.location_col_page_count().has_value());
  EXPECT_EQ(*second.location_col_page_count(), 2U);
}

TEST(PivotTableWriter, GrandTotalsDefaultTrueOmitsAttributes) {
  // When grand totals are ON (the default), the attributes stay absent so
  // the output matches what Excel emits for the default state.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"0\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  EXPECT_TRUE(parsed_or.value().grand_totals_rows());
  EXPECT_TRUE(parsed_or.value().grand_totals_cols());

  const std::string written = write_pivot_table_definition(parsed_or.value());
  EXPECT_EQ(written.find("rowGrandTotals"), std::string::npos) << "xml=" << written;
  EXPECT_EQ(written.find("colGrandTotals"), std::string::npos) << "xml=" << written;
  // The required <location> attributes (including the value 0) are still
  // emitted, since they were present in the source.
  EXPECT_NE(written.find("firstHeaderRow=\"0\""), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("firstDataRow=\"1\""), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("firstDataCol=\"0\""), std::string::npos) << "xml=" << written;
}

// ---------------------------------------------------------------------------
// Item cache index round-trips verbatim (item b)
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, ItemCacheIndexRoundTripsVerbatim) {
  // Excel can emit items whose x indices are neither sequential nor in
  // ascending order (e.g. a manually reordered field). The writer must
  // re-emit the captured indices, not synthesise 0,1,2.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"1\"><pivotField axis=\"axisRow\"><items count=\"3\">");
  xml.append("<item x=\"2\"/><item x=\"0\" h=\"1\"/><item x=\"1\"/>");
  xml.append("</items></pivotField></pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& table = parsed_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  ASSERT_EQ(table.fields()[0].items.size(), 3U);
  EXPECT_TRUE(table.fields()[0].items[0].has_cache_index);
  EXPECT_EQ(table.fields()[0].items[0].cache_index, 2U);
  EXPECT_EQ(table.fields()[0].items[1].cache_index, 0U);
  EXPECT_FALSE(table.fields()[0].items[1].visible);
  EXPECT_EQ(table.fields()[0].items[2].cache_index, 1U);

  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<item x=\"2\"/>"), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("<item x=\"0\" h=\"1\"/>"), std::string::npos) << "xml=" << written;
  EXPECT_NE(written.find("<item x=\"1\"/>"), std::string::npos) << "xml=" << written;
}

// ---------------------------------------------------------------------------
// Position-keyed passthrough keeps schema child order (item c)
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, PassthroughElementsKeepSchemaOrder) {
  // rowItems (after rowFields) and pageFields (after colFields) precede
  // dataFields in CT_pivotTableDefinition; pivotTableStyleInfo trails it.
  // A single tail buffer would emit rowItems/pageFields after dataFields
  // and trip Excel's repair. Verify each lands in its schema slot.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:D10\"/>");
  xml.append("<pivotFields count=\"2\"><pivotField axis=\"axisRow\"/><pivotField dataField=\"1\"/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<rowItems count=\"1\"><i><x/></i></rowItems>");
  xml.append("<colFields count=\"1\"><field x=\"-2\"/></colFields>");
  xml.append("<pageFields count=\"1\"><pageField fld=\"0\"/></pageFields>");
  xml.append("<dataFields count=\"1\"><dataField name=\"Sum of V\" fld=\"1\" subtotal=\"sum\"/></dataFields>");
  xml.append("<pivotTableStyleInfo name=\"PivotStyleLight16\"/>");
  xml.append("</pivotTableDefinition>");

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const std::string written = write_pivot_table_definition(parsed_or.value());

  const std::size_t p_rowfields = written.find("<rowFields");
  const std::size_t p_rowitems = written.find("<rowItems");
  const std::size_t p_colfields = written.find("<colFields");
  const std::size_t p_pagefields = written.find("<pageFields");
  const std::size_t p_datafields = written.find("<dataFields");
  const std::size_t p_styleinfo = written.find("<pivotTableStyleInfo");
  ASSERT_NE(p_rowitems, std::string::npos) << "xml=" << written;
  ASSERT_NE(p_pagefields, std::string::npos) << "xml=" << written;
  ASSERT_NE(p_styleinfo, std::string::npos) << "xml=" << written;
  // rowFields < rowItems < colFields < pageFields < dataFields < styleInfo.
  EXPECT_LT(p_rowfields, p_rowitems);
  EXPECT_LT(p_rowitems, p_colfields);
  EXPECT_LT(p_colfields, p_pagefields);
  EXPECT_LT(p_pagefields, p_datafields);
  EXPECT_LT(p_datafields, p_styleinfo);

  // A second round trip must be stable (idempotent passthrough binning).
  auto reparsed_or = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << "reparse failed: " << reparsed_or.error().message;
  EXPECT_EQ(write_pivot_table_definition(reparsed_or.value()), written);
}

// ---------------------------------------------------------------------------
// H-20: an unused pivotField round-trips as None (no axis, no dataField="1")
// and the required dataCaption attribute is always emitted.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, UnusedFieldRoundTripsAsNoneAndDataCaptionEmitted) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\" dataCaption=\"Vals\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"3\">");
  xml.append("<pivotField axis=\"axisRow\"/>");  // used: Row
  xml.append("<pivotField/>");                   // unused: None
  xml.append("<pivotField dataField=\"1\"/>");   // Value
  xml.append("</pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& table = parsed_or.value();
  ASSERT_EQ(table.fields().size(), 3U);
  EXPECT_EQ(table.fields()[0].axis, pivot::PivotAxis::Row);
  EXPECT_EQ(table.fields()[1].axis, pivot::PivotAxis::None);
  EXPECT_EQ(table.fields()[2].axis, pivot::PivotAxis::Value);
  EXPECT_EQ(table.data_caption(), "Vals");

  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("dataCaption=\"Vals\""), std::string::npos) << written;
  // The unused field emits neither an axis nor dataField="1".
  EXPECT_NE(written.find("<pivotField/>"), std::string::npos) << written;

  // A second round trip keeps the None field None (no dataField="1" creep).
  auto reparsed_or = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
  EXPECT_EQ(reparsed_or.value().fields()[1].axis, pivot::PivotAxis::None);
}

TEST(PivotTableWriter, DataCaptionDefaultsToValuesWhenAbsent) {
  pivot::PivotTable table;
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("dataCaption=\"Values\""), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// Report layout (compact / tabular / outline) is read and round-tripped.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, ReadsAndRoundTripsLayoutMode) {
  struct Case {
    const char* attrs;
    pivot::PivotLayout expected;
  };
  const Case cases[] = {
      {"", pivot::PivotLayout::Compact},                // defaults
      {" compact=\"0\"", pivot::PivotLayout::Tabular},  // tabular
      {" compact=\"0\" outline=\"1\"", pivot::PivotLayout::Outline},
  };
  for (const Case& c : cases) {
    std::string xml(kXmlDecl);
    xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\"");
    xml.append(c.attrs);
    xml.append("><location ref=\"A1:B2\"/></pivotTableDefinition>");
    auto parsed_or = read_pivot_table_definition(Bytes(xml));
    ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
    EXPECT_EQ(parsed_or.value().layout(), c.expected);
    // Write -> read keeps the same layout.
    const std::string written = write_pivot_table_definition(parsed_or.value());
    auto reparsed_or = read_pivot_table_definition(Bytes(written));
    ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
    EXPECT_EQ(reparsed_or.value().layout(), c.expected) << "written=" << written;
  }
}

// ---------------------------------------------------------------------------
// Unmodelled root + pivotField attributes round-trip verbatim.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, UnknownRootAndFieldAttributesRoundTrip) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs);
  xml.append(" name=\"P\" cacheId=\"0\" updatedVersion=\"8\" createdVersion=\"8\" itemPrintTitles=\"1\" indent=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"1\"><pivotField axis=\"axisRow\" compact=\"0\" showAll=\"0\"/></pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& table = parsed_or.value();
  auto has_attr = [](const std::vector<std::pair<std::string, std::string>>& v, const std::string& n,
                     const std::string& val) {
    for (const auto& [name, value] : v) {
      if (name == n) {
        return value == val;
      }
    }
    return false;
  };
  EXPECT_TRUE(has_attr(table.passthrough_attrs(), "updatedVersion", "8"));
  EXPECT_TRUE(has_attr(table.passthrough_attrs(), "createdVersion", "8"));
  EXPECT_TRUE(has_attr(table.passthrough_attrs(), "itemPrintTitles", "1"));
  EXPECT_TRUE(has_attr(table.passthrough_attrs(), "indent", "0"));
  // Modelled attributes are NOT double-captured into the passthrough list.
  for (const auto& [name, value] : table.passthrough_attrs()) {
    (void)value;
    EXPECT_NE(name, "name");
    EXPECT_NE(name, "cacheId");
  }
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_TRUE(has_attr(table.fields()[0].passthrough_attrs, "compact", "0"));
  EXPECT_TRUE(has_attr(table.fields()[0].passthrough_attrs, "showAll", "0"));

  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("updatedVersion=\"8\""), std::string::npos) << written;
  EXPECT_NE(written.find("createdVersion=\"8\""), std::string::npos) << written;
  EXPECT_NE(written.find("itemPrintTitles=\"1\""), std::string::npos) << written;
  EXPECT_NE(written.find("indent=\"0\""), std::string::npos) << written;
  EXPECT_NE(written.find("compact=\"0\""), std::string::npos) << written;
  EXPECT_NE(written.find("showAll=\"0\""), std::string::npos) << written;
  // `name="P"` appears once (not duplicated by the passthrough path).
  EXPECT_EQ(written.find("name=\"P\""), written.rfind("name=\"P\""));
}

// ---------------------------------------------------------------------------
// Attribute escaping: a field name with newline / tab / quote survives the
// attribute-context escaper round trip.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, FieldNameWithControlCharsRoundTrips) {
  pivot::PivotTable table;
  table.set_name("a\nb\tc\"d");
  table.set_anchor(0U, 0U, 1U, 1U);
  pivot::PivotField f;
  f.axis = pivot::PivotAxis::Row;
  f.custom_name = "x\ty\nz";
  table.mutable_fields().push_back(std::move(f));

  const std::string xml = write_pivot_table_definition(table);
  // The newline / tab must be emitted as numeric character references, not
  // raw control bytes (which some XML parsers silently normalise to space).
  EXPECT_NE(xml.find("&#10;"), std::string::npos) << xml;
  EXPECT_NE(xml.find("&#9;"), std::string::npos) << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  EXPECT_EQ(parsed_or.value().name(), "a\nb\tc\"d");
  ASSERT_EQ(parsed_or.value().fields().size(), 1U);
  EXPECT_EQ(parsed_or.value().fields()[0].custom_name, "x\ty\nz");
}

}  // namespace
}  // namespace formulon::io
