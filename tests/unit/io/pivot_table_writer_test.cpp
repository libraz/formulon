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

}  // namespace
}  // namespace formulon::io
