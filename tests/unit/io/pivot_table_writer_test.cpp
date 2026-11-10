//
// Unit tests for `formulon::io::write_pivot_table_definition`. Each test
// builds a hand-rolled `pivot::PivotTable`, writes the XML, and (where
// useful) pipes the bytes back through the symmetric reader to assert a
// clean round-trip on the data the reader actually consumes.

#include "io/pivot_table_writer.h"

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <variant>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/pivot_cache_writer.h"
#include "io/pivot_table_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "workbook.h"

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

TEST(PivotTableWriter, LocationRequiredAttrsDefaultIndependentlyAndPreserveZero) {
  pivot::PivotTable defaults;
  const std::string default_xml = write_pivot_table_definition(defaults);
  EXPECT_NE(default_xml.find("<location ref=\"A1:A1\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>"),
            std::string::npos)
      << "xml=" << default_xml;
  EXPECT_EQ(default_xml.find("rowPageCount"), std::string::npos) << "xml=" << default_xml;
  EXPECT_EQ(default_xml.find("colPageCount"), std::string::npos) << "xml=" << default_xml;

  pivot::PivotTable partial;
  partial.set_location_attributes(std::optional<std::uint32_t>(0U), std::nullopt, std::nullopt, std::nullopt,
                                  std::nullopt);
  const std::string partial_xml = write_pivot_table_definition(partial);
  EXPECT_NE(partial_xml.find("<location ref=\"A1:A1\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"1\"/>"),
            std::string::npos)
      << "xml=" << partial_xml;
}

// ---------------------------------------------------------------------------
// `<location ref>` against a caller-supplied rendered span.
//
// Nothing in the model keeps a span in step with the fields, so a table
// assembled in memory keeps whatever placeholder span it was created
// with. Emitting that understates the report, and Excel terminates when it
// refreshes a pivot whose `ref` is smaller than the grid it draws. The
// writer therefore raises such a span to the caller's projection. It only
// ever raises: over-sizing a `ref` is harmless, so a span deliberately set
// wider stands, and a span read out of a file is left alone entirely.
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, RenderedSpanRaisesAPlaceholderSpan) {
  pivot::PivotTable table;
  table.set_anchor(0U, 0U, 1U, 1U);  // the placeholder `fm_sheet_add_pivot` installs
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{4U, 3U});
  EXPECT_NE(xml.find("<location ref=\"A1:C4\""), std::string::npos) << "xml=" << xml;
}

TEST(PivotTableWriter, RenderedSpanIsOffsetFromTheAnchorNotFromA1) {
  pivot::PivotTable table;
  table.set_anchor(2U, 4U, 1U, 1U);  // E3
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{4U, 3U});
  EXPECT_NE(xml.find("<location ref=\"E3:G6\""), std::string::npos) << "xml=" << xml;
}

// A span read from a file describes the report Excel saved. Once an edit
// makes it stale, the projection replaces it outright, shrinking included:
// Excel's re-save of a filtered report drops the rows the filter hid.
TEST(PivotTableWriter, RenderedSpanReplacesAStaleFileSpan) {
  pivot::PivotTable table;
  table.set_anchor(2U, 8U, 6U, 2U);  // I3:J8 as read
  table.mark_span_authored();
  table.clear_span_authored();
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{2U, 2U});
  EXPECT_NE(xml.find("<location ref=\"I3:J4\""), std::string::npos) << "xml=" << xml;
}

TEST(PivotTableWriter, RenderedSpanKeepsASpanTheCallerWidened) {
  pivot::PivotTable table;
  table.set_anchor(2U, 8U, 6U, 2U);
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{2U, 2U});
  EXPECT_NE(xml.find("<location ref=\"I3:J8\""), std::string::npos) << "xml=" << xml;
}

// The ref covers the report body only. Excel's own files: one page field over
// a body at I3 is I3:J8; two page fields over-then-down is I3:J8; three
// page fields wrapped by two push the body to I4:J9.
TEST(PivotTableWriter, LocationRefExcludesTheSinglePageFieldBlock) {
  pivot::PivotTable table;
  table.set_anchor(2U, 8U, 1U, 1U);  // I3
  table.set_location_attributes(1U, 1U, 1U, 1U, 1U);
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{6U, 2U, 2U});
  EXPECT_NE(xml.find("<location ref=\"I3:J8\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" "
                     "rowPageCount=\"1\" colPageCount=\"1\"/>"),
            std::string::npos)
      << "xml=" << xml;
}

TEST(PivotTableWriter, LocationRefExcludesTwoPageFieldsOverThenDown) {
  pivot::PivotTable table;
  table.set_anchor(2U, 8U, 1U, 1U);
  table.set_location_attributes(1U, 1U, 1U, 1U, 2U);
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{6U, 2U, 2U});
  EXPECT_NE(xml.find("<location ref=\"I3:J8\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" "
                     "rowPageCount=\"1\" colPageCount=\"2\"/>"),
            std::string::npos)
      << "xml=" << xml;
}

TEST(PivotTableWriter, LocationRefStartsBelowAPageBlockTallerThanTheAnchorRoom) {
  pivot::PivotTable table;
  table.set_anchor(2U, 8U, 1U, 1U);  // I3, but the three-row page block needs I1:M2 plus a gap
  table.set_location_attributes(1U, 1U, 1U, 2U, 2U);
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{6U, 2U, 3U});
  EXPECT_NE(xml.find("<location ref=\"I4:J9\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" "
                     "rowPageCount=\"2\" colPageCount=\"2\"/>"),
            std::string::npos)
      << "xml=" << xml;
}

TEST(PivotTableWriter, TwoPageFieldsDownThenOverCarryTheirRowCount) {
  pivot::PivotTable table;
  table.set_anchor(2U, 8U, 1U, 1U);
  table.set_location_attributes(1U, 1U, 1U, 2U, 1U);
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{6U, 2U, 3U});
  EXPECT_NE(xml.find("<location ref=\"I4:J9\""), std::string::npos) << "xml=" << xml;
  EXPECT_NE(xml.find("rowPageCount=\"2\" colPageCount=\"1\"/>"), std::string::npos) << "xml=" << xml;
}

TEST(PivotTableWriter, AnAuthoredSpanIsReEmittedAndNeverReprojected) {
  // The span here came from a `ref` Excel wrote. Replacing it with our own
  // projection would make a read -> write cycle rewrite authored bytes,
  // which is exactly what the round-trip fidelity gate exists to catch --
  // so a projection that disagrees must lose.
  std::string source(kXmlDecl);
  source.append("<pivotTableDefinition");
  source.append(kPivotNs);
  source.append(" name=\"P\" cacheId=\"1\">");
  source.append("<location ref=\"A3:D10\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\"/>");
  source.append("<pivotFields count=\"0\"/>");
  source.append("</pivotTableDefinition>");

  auto parsed_or = read_pivot_table_definition(Bytes(source));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_TRUE(parsed.has_authored_span());

  const std::string xml = write_pivot_table_definition(parsed, PivotRenderedSpan{99U, 99U});
  EXPECT_NE(xml.find("<location ref=\"A3:D10\""), std::string::npos) << "xml=" << xml;
}

TEST(PivotTableWriter, AWiderDeliberateSpanOutranksASmallerProjection) {
  // `fm_sheet_set_pivot_anchor` takes the reserved rectangle, so a caller
  // that passes one means it. Over-reserving is not the failure mode --
  // Excel recomputes the range on refresh -- so the projection must not
  // shrink a range someone asked for, or that argument would be dead.
  pivot::PivotTable table;
  table.set_anchor(0U, 0U, 8U, 4U);
  const std::string xml = write_pivot_table_definition(table, PivotRenderedSpan{4U, 2U});
  EXPECT_NE(xml.find("<location ref=\"A1:D8\""), std::string::npos) << "xml=" << xml;
}

TEST(PivotTableWriter, WithoutAProjectionTheModelSpanIsEmitted) {
  // The projection is optional: a caller that cannot evaluate the pivot
  // passes nothing and the model's own span is emitted, as it always was.
  pivot::PivotTable table;
  table.set_anchor(0U, 0U, 2U, 2U);
  EXPECT_NE(write_pivot_table_definition(table).find("<location ref=\"A1:B2\""), std::string::npos);
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
// sortType
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, AscendingSortSpellsOutSortType) {
  // An absent attribute reads back as manual, so the struct default
  // (ascending, not manual) must be written explicitly.
  pivot::PivotTable table;
  pivot::PivotField field;
  field.axis = pivot::PivotAxis::Row;
  table.mutable_fields().push_back(std::move(field));

  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("sortType=\"ascending\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, ManualSortOmitsSortTypeAttribute) {
  pivot::PivotTable table;
  pivot::PivotField field;
  field.axis = pivot::PivotAxis::Row;
  field.sort.manual = true;
  table.mutable_fields().push_back(std::move(field));

  const std::string xml = write_pivot_table_definition(table);
  EXPECT_EQ(xml.find("sortType"), std::string::npos) << xml;
}

TEST(PivotTableWriter, DescendingAndManualSortRoundTrip) {
  pivot::PivotTable table;
  pivot::PivotField descending;
  descending.axis = pivot::PivotAxis::Row;
  descending.sort.ascending = false;
  table.mutable_fields().push_back(std::move(descending));

  pivot::PivotField manual;
  manual.axis = pivot::PivotAxis::Row;
  manual.sort.manual = true;
  table.mutable_fields().push_back(std::move(manual));

  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("sortType=\"descending\""), std::string::npos) << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_EQ(parsed.fields().size(), 2U);
  EXPECT_FALSE(parsed.fields()[0].sort.ascending);
  EXPECT_FALSE(parsed.fields()[0].sort.manual);
  EXPECT_TRUE(parsed.fields()[1].sort.manual);
}

// ---------------------------------------------------------------------------
// Values (Sigma) pseudo-field marker
// ---------------------------------------------------------------------------

TEST(PivotTableWriter, ValuesFieldMarkerRoundTripsAtItsDocumentPosition) {
  pivot::PivotTable table;
  pivot::PivotField field0;
  field0.axis = pivot::PivotAxis::Col;
  table.mutable_fields().push_back(std::move(field0));
  pivot::PivotField field1;
  field1.axis = pivot::PivotAxis::Value;
  table.mutable_fields().push_back(std::move(field1));

  table.mutable_col_field_order().push_back(0U);
  table.set_col_values_position(1U);  // marker sits after the real field

  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<colFields count=\"2\">"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<field x=\"0\"/><field x=\"-2\"/>"), std::string::npos) << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_EQ(parsed.col_field_order().size(), 1U);
  EXPECT_EQ(parsed.col_field_order()[0], 0U);
  ASSERT_TRUE(parsed.col_values_position().has_value());
  EXPECT_EQ(parsed.col_values_position().value(), 1U);
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
  // The Value-axis field survives only while a data field refers to it.
  pivot::PivotDataField df;
  df.name = "Sum of F";
  df.field_index = 2U;
  table.mutable_data_fields().push_back(std::move(df));

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
  EXPECT_NE(xml.find("&lt;R&amp;D>"), std::string::npos);
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

// `active_filters()` is persisted: an entry added in memory is written as a
// `<filters>` entry and decodes back into the same list on reload.
TEST(PivotTableWriter, ActiveFiltersSerializeIntoFiltersBlock) {
  pivot::PivotTable table;
  table.set_name("P");
  table.set_pivot_cache_id(1U);
  table.set_anchor(0U, 0U, 4U, 2U);
  pivot::PivotField region;
  region.source_name = "Region";
  region.axis = pivot::PivotAxis::Row;
  table.mutable_fields().push_back(std::move(region));
  table.mutable_row_field_order().push_back(0U);

  pivot::PivotFilter filter;
  filter.axis = pivot::PivotAxis::Row;
  filter.field_name = "Region";
  filter.type = pivot::FilterType::LabelBeginsWith;
  filter.value = std::string("Nor");
  table.mutable_active_filters().push_back(std::move(filter));

  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<filters count=\"1\"><filter fld=\"0\" type=\"captionBeginsWith\" evalOrder=\"-1\" "
                         "id=\"1\" stringValue1=\"Nor\"><autoFilter ref=\"A1\"><filterColumn colId=\"0\">"
                         "<customFilters><customFilter val=\"Nor*\"/></customFilters></filterColumn></autoFilter>"
                         "</filter></filters>"),
            std::string::npos)
      << written;

  auto reparsed_or = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
  ASSERT_EQ(reparsed_or.value().active_filters().size(), 1U);
  const pivot::PivotFilter& back = reparsed_or.value().active_filters()[0];
  EXPECT_EQ(back.type, pivot::FilterType::LabelBeginsWith);
  EXPECT_EQ(std::get<std::string>(back.value), "Nor");
  ASSERT_TRUE(back.source.has_value());
  EXPECT_EQ(back.source->field_index, 0U);
  EXPECT_TRUE(reparsed_or.value().retained_filters().empty());
  const std::string rewritten = write_pivot_table_definition(reparsed_or.value());
  const auto block = [](const std::string& xml) {
    return xml.substr(xml.find("<filters"), xml.rfind("</filters>") - xml.find("<filters"));
  };
  EXPECT_EQ(block(rewritten), block(written));
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
  df.name = "PoB";
  df.field_index = 0U;
  df.aggregation = pivot::Aggregation::Sum;
  df.show_as = pivot::ShowValuesAs::PercentOf;
  df.show_as_base_field = 2U;
  df.show_as_base_item = pivot::kShowAsBasePrev;
  table.mutable_data_fields().push_back(std::move(df));

  const std::string xml = write_pivot_table_definition(table);
  // Sanity: emitted attributes appear in the written bytes.
  EXPECT_NE(xml.find("showDataAs=\"percent\""), std::string::npos) << "xml=" << xml;
  EXPECT_NE(xml.find("baseField=\"2\""), std::string::npos);
  EXPECT_NE(xml.find("baseItem=\"1048828\""), std::string::npos);

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotTable& parsed = parsed_or.value();
  ASSERT_EQ(parsed.data_fields().size(), 1U);
  const pivot::PivotDataField& got = parsed.data_fields()[0];
  EXPECT_EQ(got.show_as, pivot::ShowValuesAs::PercentOf);
  ASSERT_TRUE(got.show_as_base_field.has_value());
  EXPECT_EQ(*got.show_as_base_field, 2U);
  ASSERT_TRUE(got.show_as_base_item.has_value());
  EXPECT_EQ(*got.show_as_base_item, pivot::kShowAsBasePrev);
}

// Parent, running-total-percent and rank modes have no `showDataAs` spelling:
// Excel stores them in the `<x14:dataField pivotShowAs>` extension of the
// `<dataField>`, next to `baseField` / `baseItem`.
TEST(PivotTableWriter, X14ShowAsModesRoundTripThroughTheDataFieldExtension) {
  struct Case {
    pivot::ShowValuesAs mode;
    const char* name;
  };
  const Case cases[] = {
      {pivot::ShowValuesAs::PercentOfParentRow, "percentOfParentRow"},
      {pivot::ShowValuesAs::PercentOfParentCol, "percentOfParentCol"},
      {pivot::ShowValuesAs::PercentOfParent, "percentOfParent"},
      {pivot::ShowValuesAs::PercentOfRunningTotal, "percentOfRunningTotal"},
      {pivot::ShowValuesAs::RankAscending, "rankAscending"},
      {pivot::ShowValuesAs::RankDescending, "rankDescending"},
  };
  for (const Case& c : cases) {
    pivot::PivotTable table;
    table.set_name("T");
    table.set_anchor(0U, 0U, 1U, 1U);
    pivot::PivotField placeholder;
    placeholder.axis = pivot::PivotAxis::Value;
    placeholder.custom_name = "X";
    table.mutable_fields().push_back(std::move(placeholder));
    pivot::PivotDataField df;
    df.name = "V";
    df.field_index = 0U;
    df.show_as = c.mode;
    df.show_as_base_field = 0U;
    df.show_as_base_item = 0U;
    table.mutable_data_fields().push_back(std::move(df));

    const std::string xml = write_pivot_table_definition(table);
    EXPECT_EQ(xml.find("showDataAs="), std::string::npos) << c.name << " xml=" << xml;
    EXPECT_NE(xml.find("<extLst><ext uri=\"{E15A36E0-9728-4e99-A89B-3F7291B0FE68}\" "
                       "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
                       "<x14:dataField pivotShowAs=\"" +
                       std::string(c.name) + "\"/></ext></extLst></dataField>"),
              std::string::npos)
        << c.name << " xml=" << xml;

    auto parsed_or = read_pivot_table_definition(Bytes(xml));
    ASSERT_TRUE(static_cast<bool>(parsed_or)) << c.name << ": " << parsed_or.error().message;
    ASSERT_EQ(parsed_or.value().data_fields().size(), 1U);
    const pivot::PivotDataField& got = parsed_or.value().data_fields()[0];
    EXPECT_EQ(got.show_as, c.mode) << c.name;
    EXPECT_EQ(got.show_as_base_field, std::optional<std::uint32_t>(0U));
    EXPECT_EQ(got.show_as_base_item, std::optional<std::uint32_t>(0U));
  }
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

TEST(PivotTableWriter, SubtotalHiddenItemsSurvivesRoundTripOnce) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\" subtotalHiddenItems=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  EXPECT_TRUE(parsed_or.value().subtotal_hidden_items());

  const std::string written = write_pivot_table_definition(parsed_or.value());
  const std::size_t first = written.find("subtotalHiddenItems=\"1\"");
  ASSERT_NE(first, std::string::npos) << "xml=" << written;
  EXPECT_EQ(written.find("subtotalHiddenItems", first + 1), std::string::npos) << "xml=" << written;
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
  xml.append("<pivotFields count=\"3\"><pivotField axis=\"axisRow\"/><pivotField dataField=\"1\"/>");
  xml.append("<pivotField axis=\"axisPage\"/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<rowItems count=\"1\"><i><x/></i></rowItems>");
  xml.append("<colFields count=\"1\"><field x=\"-2\"/></colFields>");
  xml.append("<pageFields count=\"1\"><pageField fld=\"2\"/></pageFields>");
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
  xml.append("<pivotField axis=\"axisRow\" showAll=\"0\"/>");  // used: Row
  xml.append("<pivotField showAll=\"0\"/>");                   // unused: None
  xml.append("<pivotField dataField=\"1\" showAll=\"0\"/>");   // Value
  xml.append("</pivotFields>");
  xml.append("<dataFields count=\"1\"><dataField name=\"X\" fld=\"2\"/></dataFields>");
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
  EXPECT_NE(written.find("<pivotField showAll=\"0\"/>"), std::string::npos) << written;

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
      {"", pivot::PivotLayout::Compact},                                  // defaults
      {" outline=\"1\" outlineData=\"1\"", pivot::PivotLayout::Compact},  // as Excel writes compact
      {" compact=\"0\" compactData=\"0\"", pivot::PivotLayout::Tabular},  // as Excel writes tabular
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

// The row fields' own `compact` / `outline` pair decides the form, as Excel
// renders it; every written field then carries that form.
TEST(PivotTableWriter, RowFieldFormDecidesLayout) {
  struct Case {
    const char* root_attrs;
    const char* field_attrs;
    pivot::PivotLayout expected;
    const char* written_field_attrs;
  };
  const Case cases[] = {
      {" outline=\"1\" outlineData=\"1\"", "", pivot::PivotLayout::Compact,
       "<pivotField axis=\"axisRow\" showAll=\"0\"/>"},
      {" compact=\"0\" compactData=\"0\"", " compact=\"0\" outline=\"0\"", pivot::PivotLayout::Tabular,
       "<pivotField axis=\"axisRow\" compact=\"0\" outline=\"0\" showAll=\"0\"/>"},
      {"", " compact=\"0\"", pivot::PivotLayout::Outline, "<pivotField axis=\"axisRow\" compact=\"0\" showAll=\"0\"/>"},
  };
  for (const Case& c : cases) {
    std::string xml(kXmlDecl);
    xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\"").append(c.root_attrs);
    xml.append("><location ref=\"A1:B2\"/><pivotFields count=\"1\"><pivotField axis=\"axisRow\" showAll=\"0\"");
    xml.append(c.field_attrs);
    xml.append("/></pivotFields>");
    xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields></pivotTableDefinition>");
    auto parsed_or = read_pivot_table_definition(Bytes(xml));
    ASSERT_TRUE(static_cast<bool>(parsed_or)) << parsed_or.error().message;
    EXPECT_EQ(parsed_or.value().layout(), c.expected) << c.field_attrs;
    const std::string written = write_pivot_table_definition(parsed_or.value());
    EXPECT_NE(written.find(c.written_field_attrs), std::string::npos) << written;
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
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields></pivotTableDefinition>");

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
  // `itemPrintTitles` and `indent` are modelled, not passed through.
  EXPECT_FALSE(has_attr(table.passthrough_attrs(), "itemPrintTitles", "1"));
  EXPECT_FALSE(has_attr(table.passthrough_attrs(), "indent", "0"));
  EXPECT_TRUE(table.item_print_titles());
  EXPECT_EQ(table.indent(), 0U);
  // Modelled attributes are NOT double-captured into the passthrough list.
  for (const auto& [name, value] : table.passthrough_attrs()) {
    (void)value;
    EXPECT_NE(name, "name");
    EXPECT_NE(name, "cacheId");
  }
  ASSERT_EQ(table.fields().size(), 1U);
  // A field's `compact` / `outline` pair is its report form, modelled as the
  // table layout rather than passed through.
  EXPECT_FALSE(has_attr(table.fields()[0].passthrough_attrs, "compact", "0"));
  EXPECT_EQ(table.layout(), pivot::PivotLayout::Outline);
  EXPECT_FALSE(table.fields()[0].show_all);

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

// ---------------------------------------------------------------------------
// Subtotal entries inside `<items>`
//
// Excel closes a field's item list with the subtotals it displays, and
// treats a file whose items lack them as damaged -- it offers to repair on
// open. Nothing else notices: every part validates against the schema
// (`<item>` is optional there), and the reader deliberately skips these
// markers, so a read-write round trip compares equal while silently
// shedding them. Pin the emitted bytes, which is the only place the loss
// is visible.
// ---------------------------------------------------------------------------

pivot::PivotTable TableWithTwoItemRowField() {
  pivot::PivotTable table;
  table.set_anchor(0U, 0U, 4U, 2U);
  pivot::PivotField region;
  region.source_name = "Region";
  region.axis = pivot::PivotAxis::Row;
  region.items.push_back(pivot::PivotItem{"", true});
  region.items.push_back(pivot::PivotItem{"", true});
  table.mutable_fields().push_back(std::move(region));
  return table;
}

TEST(PivotTableWriter, ItemsCloseWithTheImplicitSubtotalEntry) {
  const std::string xml = write_pivot_table_definition(TableWithTwoItemRowField());

  EXPECT_NE(xml.find("<items count=\"3\">"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<item t=\"default\"/></items>"), std::string::npos) << xml;
  // The count covers the marker: two real items plus the subtotal.
  EXPECT_EQ(xml.find("<items count=\"2\">"), std::string::npos) << xml;
}

TEST(PivotTableWriter, SuppressedSubtotalEmitsNoMarker) {
  pivot::PivotTable table = TableWithTwoItemRowField();
  table.mutable_fields()[0].default_subtotal = false;

  const std::string xml = write_pivot_table_definition(table);

  EXPECT_NE(xml.find("<items count=\"2\">"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<item t="), std::string::npos) << xml;
  // `defaultSubtotal="0"` is what makes the absence deliberate rather than
  // the defect this test exists for.
  EXPECT_NE(xml.find("defaultSubtotal=\"0\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, CustomSubtotalSelectionEmitsOneEntryPerFunction) {
  pivot::PivotTable table = TableWithTwoItemRowField();
  pivot::PivotField& region = table.mutable_fields()[0];
  region.default_subtotal = false;
  region.subtotal_fns = {pivot::SubtotalFn::Sum, pivot::SubtotalFn::Count};

  const std::string xml = write_pivot_table_definition(table);

  EXPECT_NE(xml.find("<items count=\"4\">"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<item t=\"sum\"/>"), std::string::npos) << xml;
  // `Count` is the count-of-values function, spelled `countA` as an item
  // token and `countASubtotal` as the selection attribute.
  EXPECT_NE(xml.find("<item t=\"countA\"/>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<item t=\"default\"/>"), std::string::npos) << xml;
}

TEST(PivotTableWriter, CustomFunctionsReplaceTheDefaultSubtotalAsExcelWritesThem) {
  pivot::PivotTable table = TableWithTwoItemRowField();
  pivot::PivotField& region = table.mutable_fields()[0];
  region.subtotal_fns = {pivot::SubtotalFn::Max, pivot::SubtotalFn::Sum};

  const std::string xml = write_pivot_table_definition(table);

  // Excel lists the functions in a fixed order and writes neither a
  // `t="default"` item nor `defaultSubtotal` next to them.
  EXPECT_NE(xml.find(" sumSubtotal=\"1\" maxSubtotal=\"1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("<items count=\"4\">"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<item t=\"sum\"/><item t=\"max\"/></items>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("t=\"default\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("defaultSubtotal"), std::string::npos) << xml;
}

TEST(PivotTableWriter, CountAndAverageUseTheExcelAttributeNames) {
  pivot::PivotTable table = TableWithTwoItemRowField();
  table.mutable_fields()[0].subtotal_fns = {pivot::SubtotalFn::Average, pivot::SubtotalFn::Count};

  const std::string xml = write_pivot_table_definition(table);

  EXPECT_NE(xml.find(" countASubtotal=\"1\" avgSubtotal=\"1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("<item t=\"countA\"/><item t=\"avg\"/></items>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("t=\"default\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, ValueFieldWithoutItemsStaysSelfClosing) {
  // A data-only field carries no item list, so it must not acquire one
  // just because subtotals default to on.
  pivot::PivotTable table;
  table.set_anchor(0U, 0U, 1U, 1U);
  pivot::PivotField amount;
  amount.source_name = "Amount";
  amount.axis = pivot::PivotAxis::Value;
  table.mutable_fields().push_back(std::move(amount));
  pivot::PivotDataField df;
  df.name = "Sum of Amount";
  df.field_index = 0U;
  table.mutable_data_fields().push_back(std::move(df));

  const std::string xml = write_pivot_table_definition(table);

  EXPECT_EQ(xml.find("<items"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<pivotField dataField=\"1\" showAll=\"0\"/>"), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// Excel-accepted definition shape for authored tables
// ---------------------------------------------------------------------------

namespace {

pivot::PivotTable TableWithRowAndDataFields(std::size_t data_field_count) {
  pivot::PivotTable table;
  table.set_name("PT");
  table.set_anchor(2U, 0U, 10U, 3U);
  pivot::PivotField region;
  region.axis = pivot::PivotAxis::Row;
  table.mutable_fields().push_back(std::move(region));
  table.mutable_row_field_order().push_back(0U);
  for (std::size_t i = 0; i < data_field_count; ++i) {
    pivot::PivotField amount;
    amount.axis = pivot::PivotAxis::Value;
    table.mutable_fields().push_back(std::move(amount));
    pivot::PivotDataField df;
    df.name = "Sum of F" + std::to_string(i);
    df.field_index = static_cast<std::uint32_t>(i + 1U);
    table.mutable_data_fields().push_back(std::move(df));
  }
  return table;
}

}  // namespace

TEST(PivotTableWriter, TwoDataFieldsPlaceValuesMarkerOnColumns) {
  const std::string xml = write_pivot_table_definition(TableWithRowAndDataFields(2U));
  EXPECT_NE(xml.find("<colFields count=\"1\"><field x=\"-2\"/></colFields>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"1\""), std::string::npos) << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << parsed_or.error().message;
  EXPECT_TRUE(parsed_or.value().col_field_order().empty());
  ASSERT_TRUE(parsed_or.value().col_values_position().has_value());
  EXPECT_EQ(parsed_or.value().col_values_position().value(), 0U);
}

TEST(PivotTableWriter, TwoDataFieldsKeepModelValuesMarkerOnRows) {
  pivot::PivotTable table = TableWithRowAndDataFields(2U);
  table.set_row_values_position(1U);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<rowFields count=\"2\"><field x=\"0\"/><field x=\"-2\"/></rowFields>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<colFields"), std::string::npos) << xml;
  EXPECT_NE(xml.find("firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, TwoDataFieldsAppendValuesMarkerAfterRealColumnFields) {
  pivot::PivotTable table = TableWithRowAndDataFields(2U);
  pivot::PivotField product;
  product.axis = pivot::PivotAxis::Col;
  table.mutable_fields().push_back(std::move(product));
  table.mutable_col_field_order().push_back(3U);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<colFields count=\"2\"><field x=\"3\"/><field x=\"-2\"/></colFields>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("firstHeaderRow=\"1\" firstDataRow=\"3\" firstDataCol=\"1\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, SingleDataFieldWritesNoValuesMarker) {
  const std::string xml = write_pivot_table_definition(TableWithRowAndDataFields(1U));
  EXPECT_EQ(xml.find("<colFields"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("-2"), std::string::npos) << xml;
  EXPECT_NE(xml.find("firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, TabularTwoRowFieldsDeriveFirstDataCol) {
  pivot::PivotTable table = TableWithRowAndDataFields(2U);
  table.set_layout(pivot::PivotLayout::Tabular);
  pivot::PivotField product;
  product.axis = pivot::PivotAxis::Row;
  table.mutable_fields().push_back(std::move(product));
  table.mutable_row_field_order().push_back(3U);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"2\""), std::string::npos) << xml;
}

namespace {

/// No row fields: `data_field_count` data fields, plus a column field when
/// `with_col_field`. Excel's location offsets for these shapes are measured.
pivot::PivotTable TableWithoutRowFields(std::size_t data_field_count, bool with_col_field) {
  pivot::PivotTable table;
  table.set_name("PT");
  table.set_anchor(2U, 8U, 3U, 3U);
  if (with_col_field) {
    pivot::PivotField product;
    product.axis = pivot::PivotAxis::Col;
    table.mutable_fields().push_back(std::move(product));
    table.mutable_col_field_order().push_back(0U);
  }
  for (std::size_t i = 0; i < data_field_count; ++i) {
    pivot::PivotField amount;
    amount.axis = pivot::PivotAxis::Value;
    table.mutable_fields().push_back(std::move(amount));
    pivot::PivotDataField df;
    df.name = "Sum of F" + std::to_string(i);
    df.field_index = static_cast<std::uint32_t>(table.fields().size() - 1U);
    table.mutable_data_fields().push_back(std::move(df));
  }
  return table;
}

}  // namespace

TEST(PivotTableWriter, ColumnFieldWithoutRowOrDataFieldsStartsDataAtTheAnchorColumn) {
  const std::string xml = write_pivot_table_definition(TableWithoutRowFields(0U, true));
  EXPECT_NE(xml.find("firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"0\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, GridDropZonesKeepTheRowLabelColumnOfAColumnFieldReport) {
  pivot::PivotTable table = TableWithoutRowFields(1U, true);
  table.set_grid_drop_zones(true);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, GridDropZonesLoneDataFieldHasNoDropZoneRow) {
  pivot::PivotTable table = TableWithoutRowFields(1U, false);
  table.set_grid_drop_zones(true);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, GridDropZonesTwoDataFieldsWithoutRowFieldsKeepTheRowLabelColumn) {
  pivot::PivotTable table = TableWithoutRowFields(2U, false);
  table.set_grid_drop_zones(true);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, AuthoredFieldsCarryShowAllZeroAndUpdatedVersion) {
  const std::string xml = write_pivot_table_definition(TableWithRowAndDataFields(1U));
  EXPECT_NE(xml.find("<pivotField axis=\"axisRow\" showAll=\"0\" sortType=\"ascending\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<pivotField dataField=\"1\" showAll=\"0\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find(" updatedVersion=\"8\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, AuthoredTableCarriesValuesRowExtension) {
  const std::string xml = write_pivot_table_definition(TableWithRowAndDataFields(1U));
  EXPECT_NE(xml.find("<extLst><ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("<x14:pivotTableDefinition hideValuesRow=\"1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("</ext></extLst></pivotTableDefinition>"), std::string::npos) << xml;
}

TEST(PivotTableWriter, CapturedExtLstIsKeptAndNotDuplicated) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/><pivotFields count=\"0\"/>");
  xml.append("<extLst><ext uri=\"{AAAA}\"/></extLst></pivotTableDefinition>");
  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << parsed_or.error().message;
  const std::string written = write_pivot_table_definition(parsed_or.value());
  EXPECT_NE(written.find("{AAAA}"), std::string::npos) << written;
  EXPECT_EQ(written.find("hideValuesRow"), std::string::npos) << written;
  EXPECT_EQ(written.find("<extLst"), written.rfind("<extLst")) << written;
}

TEST(PivotTableWriter, DataFieldFlagFollowsTheDataFieldList) {
  pivot::PivotTable table = TableWithRowAndDataFields(2U);
  // Qty (field 2) is no longer aggregated, so it must not claim dataField="1".
  table.mutable_data_fields().pop_back();
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<pivotField dataField=\"1\" showAll=\"0\"/><pivotField showAll=\"0\"/>"), std::string::npos)
      << xml;
}

TEST(PivotTableWriter, CapturedUpdatedVersionIsNotDuplicatedAndShowAllTrueIsUnwritten) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\" updatedVersion=\"6\">");
  xml.append("<location ref=\"A1:B2\"/><pivotFields count=\"1\"><pivotField axis=\"axisRow\" showAll=\"1\"/>");
  xml.append("</pivotFields><rowFields count=\"1\"><field x=\"0\"/></rowFields></pivotTableDefinition>");
  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << parsed_or.error().message;
  const std::string written = write_pivot_table_definition(parsed_or.value());
  EXPECT_NE(written.find("updatedVersion=\"6\""), std::string::npos) << written;
  EXPECT_EQ(written.find("updatedVersion=\"8\""), std::string::npos) << written;
  EXPECT_TRUE(parsed_or.value().fields()[0].show_all);
  EXPECT_EQ(written.find("showAll="), std::string::npos) << written;
}

// ---------------------------------------------------------------------------
// `<pageFields>`
// ---------------------------------------------------------------------------
//
// The element is written from the model, for a loaded table and one
// assembled in memory alike: a saved `axis="axisPage"` field with no
// page-field entry leaves Excel with no page-axis order to draw.

pivot::PivotTable BuildPageAxisTable() {
  pivot::PivotTable table;
  table.set_anchor(0U, 0U, 1U, 1U);
  pivot::PivotField region;
  region.source_name = "Region";
  region.axis = pivot::PivotAxis::Page;
  table.mutable_fields().push_back(std::move(region));
  pivot::PivotField amount;
  amount.source_name = "Amount";
  amount.axis = pivot::PivotAxis::Value;
  table.mutable_fields().push_back(std::move(amount));
  return table;
}

TEST(PivotTableWriter, PageAxisFieldGainsAPageFieldsBlock) {
  const std::string xml = write_pivot_table_definition(BuildPageAxisTable());

  EXPECT_NE(xml.find("<pageFields count=\"1\"><pageField fld=\"0\" hier=\"-1\"/></pageFields>"), std::string::npos)
      << xml;
  // Schema order: `<pageFields>` sits between `<colFields>` and
  // `<dataFields>`, so a definition that emits both must not invert them.
  EXPECT_LT(xml.find("<pageFields"), xml.find("</pivotTableDefinition>")) << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  ASSERT_EQ(parsed_or.value().page_fields().size(), 1U);
  EXPECT_EQ(parsed_or.value().page_fields()[0].field_index, 0U);
  EXPECT_FALSE(parsed_or.value().page_fields()[0].item_index.has_value());
}

TEST(PivotTableWriter, TableWithoutAPageAxisEmitsNoPageFields) {
  pivot::PivotTable table;
  table.set_anchor(0U, 0U, 1U, 1U);
  pivot::PivotField region;
  region.source_name = "Region";
  region.axis = pivot::PivotAxis::Row;
  table.mutable_fields().push_back(std::move(region));

  EXPECT_EQ(write_pivot_table_definition(table).find("<pageFields"), std::string::npos);
}

TEST(PivotTableWriter, PageFieldsAreWrittenOnceFromTheModelWithTheirAttributes) {
  pivot::PivotTable table = BuildPageAxisTable();
  pivot::PivotPageField page;
  page.field_index = 0U;
  page.item_index = 3U;
  page.hierarchy = std::nullopt;
  page.name = "[R].&[\"x\"]";
  page.caption = "East";
  table.mutable_page_fields().push_back(page);

  const std::string xml = write_pivot_table_definition(table);

  EXPECT_NE(xml.find("<pageFields count=\"1\"><pageField fld=\"0\" item=\"3\" name=\"[R].&amp;[&quot;x&quot;]\" "
                     "cap=\"East\"/></pageFields>"),
            std::string::npos)
      << xml;
  EXPECT_EQ(xml.find("<pageFields", xml.find("<pageFields") + 1), std::string::npos) << xml;
}

// Two page fields read from a file: an order set on the model is what is
// saved, and a field moved onto the page axis later joins it last.
TEST(PivotTableWriter, PageOrderSetOnALoadedTableRoundTrips) {
  std::string authored(kXmlDecl);
  authored.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  authored.append("<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  authored.append("<pivotFields count=\"3\"><pivotField axis=\"axisPage\"/><pivotField axis=\"axisPage\"/>");
  authored.append("<pivotField axis=\"axisRow\"/></pivotFields>");
  authored.append("<rowFields count=\"1\"><field x=\"2\"/></rowFields>");
  authored.append("<pageFields count=\"2\"><pageField fld=\"0\" item=\"1\" hier=\"-1\"/>");
  authored.append("<pageField fld=\"1\" hier=\"-1\"/></pageFields>");
  authored.append("</pivotTableDefinition>");
  auto loaded_or = read_pivot_table_definition(Bytes(authored));
  ASSERT_TRUE(static_cast<bool>(loaded_or)) << loaded_or.error().message;
  pivot::PivotTable table = std::move(loaded_or.value());

  std::reverse(table.mutable_page_fields().begin(), table.mutable_page_fields().end());
  auto reread_or = read_pivot_table_definition(Bytes(write_pivot_table_definition(table)));
  ASSERT_TRUE(static_cast<bool>(reread_or)) << reread_or.error().message;
  EXPECT_EQ(reread_or.value().page_field_order(), (std::vector<std::uint32_t>{1U, 0U}));
  ASSERT_NE(reread_or.value().find_page_field(0U), nullptr);
  EXPECT_EQ(reread_or.value().find_page_field(0U)->item_index, std::optional<std::uint32_t>(1U));

  table.mutable_row_field_order().clear();
  table.set_field_axis(2U, pivot::PivotAxis::Page);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<pageFields count=\"3\"><pageField fld=\"1\" hier=\"-1\"/><pageField fld=\"0\" item=\"1\" "
                     "hier=\"-1\"/><pageField fld=\"2\" hier=\"-1\"/></pageFields>"),
            std::string::npos)
      << xml;
  auto moved_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(moved_or)) << moved_or.error().message;
  EXPECT_EQ(moved_or.value().page_field_order(), (std::vector<std::uint32_t>{1U, 0U, 2U}));

  // Leaving the axis drops the entry; joining again appends it.
  table.set_field_axis(1U, pivot::PivotAxis::None);
  EXPECT_EQ(table.page_field_order(), (std::vector<std::uint32_t>{0U, 2U}));
  table.set_field_axis(1U, pivot::PivotAxis::Page);
  EXPECT_EQ(table.page_field_order(), (std::vector<std::uint32_t>{0U, 2U, 1U}));
}

TEST(PivotTableWriter, APageFieldWhoseAxisWasSetDirectlyJoinsTheOrderLast) {
  pivot::PivotTable table = BuildPageAxisTable();
  table.mutable_page_fields().push_back(pivot::PivotPageField{0U, std::nullopt});
  pivot::PivotField quarter;
  quarter.source_name = "Quarter";
  quarter.axis = pivot::PivotAxis::Page;
  table.mutable_fields().push_back(std::move(quarter));
  EXPECT_EQ(table.page_field_order(), (std::vector<std::uint32_t>{0U, 2U}));
  table.mutable_fields()[0].axis = pivot::PivotAxis::Row;
  EXPECT_EQ(table.page_field_order(), (std::vector<std::uint32_t>{2U}));
  EXPECT_EQ(table.find_page_field(0U), nullptr);
}

// ---------------------------------------------------------------------------
// Items synthesized from a cache write plan
// ---------------------------------------------------------------------------

// A one-field cache whose records read West, East, blank, North; the shared
// items are therefore West(0) East(1) blank(2) North(3), ascending 1 3 0 2.
struct DerivedRegion {
  pivot::PivotCache cache;
  PivotCacheWritePlan plan;

  DerivedRegion() {
    const eval::ScopedEvalProfile scope(default_excel_profile());
    pivot::PivotCacheField field;
    field.name = "Region";
    cache.mutable_fields().push_back(std::move(field));
    for (const char* text : {"West", "East", "", "North"}) {
      pivot::PivotCacheRecord record;
      if (text[0] == '\0') {
        record.cells.push_back(Value::blank());
      } else {
        cache.mutable_text_storage().emplace_back(text);
        record.cells.push_back(Value::text(cache.mutable_text_storage().back()));
      }
      cache.mutable_records().push_back(std::move(record));
    }
    plan = build_pivot_cache_write_plan(cache, {});
  }
};

pivot::PivotTable RegionTable(pivot::PivotAxis axis) {
  pivot::PivotTable table;
  table.set_anchor(0U, 0U, 1U, 1U);
  pivot::PivotField region;
  region.source_name = "Region";
  region.axis = axis;
  table.mutable_fields().push_back(std::move(region));
  return table;
}

bool Has(const std::string& xml, std::string_view needle) {
  return xml.find(needle) != std::string::npos;
}

TEST(PivotTableWriter, SynthesizedItemsFollowAscendingOrderWithBlankLast) {
  const DerivedRegion derived;
  pivot::PivotTable table = RegionTable(pivot::PivotAxis::Row);
  const std::string xml = write_pivot_table_definition(table, std::nullopt, &derived.plan);
  EXPECT_TRUE(Has(xml,
                  "<items count=\"5\"><item x=\"1\"/><item x=\"3\"/><item x=\"0\"/><item x=\"2\"/>"
                  "<item t=\"default\"/></items>"))
      << xml;
  // The model is not mutated.
  EXPECT_TRUE(table.fields()[0].items.empty());
}

TEST(PivotTableWriter, SynthesizedItemsReverseForDescending) {
  const DerivedRegion derived;
  pivot::PivotTable table = RegionTable(pivot::PivotAxis::Row);
  table.mutable_fields()[0].sort.ascending = false;
  const std::string xml = write_pivot_table_definition(table, std::nullopt, &derived.plan);
  EXPECT_TRUE(Has(xml, "<item x=\"2\"/><item x=\"0\"/><item x=\"3\"/><item x=\"1\"/><item t=\"default\"/>")) << xml;
}

TEST(PivotTableWriter, SynthesizedItemsKeepSharedOrderForManualSort) {
  const DerivedRegion derived;
  pivot::PivotTable table = RegionTable(pivot::PivotAxis::Col);
  table.mutable_fields()[0].sort.manual = true;
  const std::string xml = write_pivot_table_definition(table, std::nullopt, &derived.plan);
  EXPECT_TRUE(Has(xml, "<item x=\"0\"/><item x=\"1\"/><item x=\"2\"/><item x=\"3\"/><item t=\"default\"/>")) << xml;
  EXPECT_FALSE(Has(xml, "sortType=")) << xml;
}

TEST(PivotTableWriter, SynthesizedItemsCoverAPageField) {
  const DerivedRegion derived;
  pivot::PivotTable table = RegionTable(pivot::PivotAxis::Page);
  const std::string xml = write_pivot_table_definition(table, std::nullopt, &derived.plan);
  EXPECT_TRUE(Has(xml, "axis=\"axisPage\"")) << xml;
  EXPECT_TRUE(Has(xml, "<items count=\"5\">")) << xml;
}

TEST(PivotTableWriter, OffAxisDerivedFieldGetsNoSynthesizedItems) {
  const DerivedRegion derived;
  pivot::PivotTable table = RegionTable(pivot::PivotAxis::Value);
  const std::string xml = write_pivot_table_definition(table, std::nullopt, &derived.plan);
  EXPECT_FALSE(Has(xml, "<items")) << xml;
}

TEST(PivotTableWriter, AuthoredItemsBindByLabelAndMissingSharedItemsFollow) {
  const DerivedRegion derived;
  pivot::PivotTable table = RegionTable(pivot::PivotAxis::Row);
  pivot::PivotItem north;
  north.name = "North";
  pivot::PivotItem east;
  east.name = "East";
  east.visible = false;
  table.mutable_fields()[0].items.push_back(north);
  table.mutable_fields()[0].items.push_back(east);
  const std::string xml = write_pivot_table_definition(table, std::nullopt, &derived.plan);
  EXPECT_TRUE(Has(xml,
                  "<items count=\"5\"><item x=\"3\"/><item x=\"1\" h=\"1\"/><item x=\"0\"/><item x=\"2\"/>"
                  "<item t=\"default\"/></items>"))
      << xml;
}

TEST(PivotTableWriter, DefaultSubtotalOffDropsTheDefaultMarkerOfSynthesizedItems) {
  const DerivedRegion derived;
  pivot::PivotTable table = RegionTable(pivot::PivotAxis::Row);
  table.mutable_fields()[0].default_subtotal = false;
  const std::string xml = write_pivot_table_definition(table, std::nullopt, &derived.plan);
  EXPECT_TRUE(Has(xml, "defaultSubtotal=\"0\"")) << xml;
  EXPECT_TRUE(Has(xml, "<items count=\"4\">")) << xml;
  EXPECT_FALSE(Has(xml, "t=\"default\"")) << xml;
}

// ---------------------------------------------------------------------------
// Report options, item captions / collapse state and the x14 field extension
// ---------------------------------------------------------------------------

constexpr std::string_view kX14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

// One row field (Region) with three items closed by `<item t="default"/>`,
// plus the rendered-report bins a file carries.
std::string LoadedDefinition(std::string_view default_marker) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs);
  xml.append(" name=\"P\" cacheId=\"1\" dataCaption=\"Values\" updatedVersion=\"8\" useAutoFormatting=\"1\">");
  xml.append("<location ref=\"A3:B7\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  xml.append("<pivotFields count=\"3\">");
  xml.append("<pivotField axis=\"axisRow\" showAll=\"0\"><items count=\"4\"><item x=\"2\"/><item x=\"0\"/>");
  xml.append("<item x=\"1\"/>").append(default_marker).append("</items></pivotField>");
  xml.append("<pivotField axis=\"axisPage\" showAll=\"0\"/>");
  xml.append("<pivotField dataField=\"1\" showAll=\"0\"/></pivotFields>");
  xml.append("<rowFields count=\"1\"><field x=\"0\"/></rowFields>");
  xml.append("<rowItems count=\"1\"><i><x/></i></rowItems>");
  xml.append("<colItems count=\"1\"><i/></colItems>");
  xml.append("<pageFields count=\"1\"><pageField fld=\"1\" hier=\"-1\"/></pageFields>");
  xml.append("<dataFields count=\"1\"><dataField name=\"Sum of Amt\" fld=\"2\" baseField=\"0\" baseItem=\"0\"/>");
  xml.append("</dataFields></pivotTableDefinition>");
  return xml;
}

pivot::PivotTable ReadLoaded(std::string_view default_marker = "<item t=\"default\"/>") {
  auto table_or = read_pivot_table_definition(Bytes(LoadedDefinition(default_marker)));
  EXPECT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  return std::move(table_or.value());
}

TEST(PivotTableWriter, ReportOptionsAreWrittenOnlyWhenSet) {
  pivot::PivotTable table;
  std::string xml = write_pivot_table_definition(table);
  // A table created in memory autofits, as every table Excel creates does.
  EXPECT_NE(xml.find(" useAutoFormatting=\"1\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("grandTotalCaption"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("enableDrill"), std::string::npos) << xml;

  table.set_grand_total_caption(std::string("TOTAL & more"));
  table.set_auto_formatting(false);
  table.set_enable_drill(false);
  xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find(" grandTotalCaption=\"TOTAL &amp; more\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("useAutoFormatting"), std::string::npos) << xml;
  EXPECT_NE(xml.find(" enableDrill=\"0\""), std::string::npos) << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << parsed_or.error().message;
  EXPECT_EQ(parsed_or.value().grand_total_caption(), std::optional<std::string>("TOTAL & more"));
  EXPECT_FALSE(parsed_or.value().auto_formatting());
  EXPECT_FALSE(parsed_or.value().enable_drill());
}

TEST(PivotTableWriter, BlankRowAndFillDownLabelsAreWrittenAsExcelWritesThem) {
  pivot::PivotTable table = TableWithTwoItemRowField();
  table.mutable_fields()[0].insert_blank_row = true;
  table.mutable_fields()[0].fill_down_labels = true;
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find(" insertBlankRow=\"1\""), std::string::npos) << xml;
  std::string ext =
      "<item t=\"default\"/></items><extLst><ext uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\" xmlns:x14=\"";
  ext.append(kX14Ns).append("\"><x14:pivotField fillDownLabels=\"1\"/></ext></extLst></pivotField>");
  EXPECT_NE(xml.find(ext), std::string::npos) << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << parsed_or.error().message;
  EXPECT_TRUE(parsed_or.value().fields()[0].insert_blank_row);
  EXPECT_TRUE(parsed_or.value().fields()[0].fill_down_labels);
}

TEST(PivotTableWriter, FillDownLabelsOnAFieldWithoutItemsOpensTheElement) {
  pivot::PivotTable table;
  pivot::PivotField field;
  field.axis = pivot::PivotAxis::None;
  field.fill_down_labels = true;
  table.mutable_fields().push_back(std::move(field));
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<x14:pivotField fillDownLabels=\"1\"/></ext></extLst></pivotField>"), std::string::npos) << xml;
}

TEST(PivotTableWriter, X14FieldExtensionKeepsItsOtherContentAndTheOtherExtensions) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  xml.append("<pivotFields count=\"1\"><pivotField axis=\"axisRow\" showAll=\"0\"><extLst>");
  xml.append("<ext uri=\"{2946ed86-a175-432a-8ac1-64e0c546d7de}\" xmlns:x14=\"").append(kX14Ns);
  xml.append("\"><x14:pivotField fillDownLabels=\"1\" ignore=\"1\"><x14:child/></x14:pivotField></ext>");
  xml.append("<ext uri=\"{AAAAAAAA-0000-0000-0000-000000000000}\"><other a=\"1\"/></ext>");
  xml.append("</extLst></pivotField></pivotFields></pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable table = std::move(table_or.value());
  ASSERT_TRUE(table.fields()[0].fill_down_labels);

  std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<ext uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\""), std::string::npos) << written;
  EXPECT_NE(written.find("<x14:pivotField fillDownLabels=\"1\" ignore=\"1\"><x14:child/></x14:pivotField></ext>"
                         "<ext uri=\"{AAAAAAAA-0000-0000-0000-000000000000}\"><other a=\"1\"/></ext></extLst>"),
            std::string::npos)
      << written;

  // Turning the option off keeps the extension's other content.
  table.mutable_fields()[0].fill_down_labels = false;
  written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<x14:pivotField ignore=\"1\"><x14:child/></x14:pivotField>"), std::string::npos) << written;
  EXPECT_EQ(written.find("fillDownLabels"), std::string::npos) << written;
}

// An x14 field extension the reader keeps verbatim (here: a second element
// beside `<pivotField>`) still follows the model's fillDownLabels.
TEST(PivotTableWriter, FillDownLabelsIsPatchedIntoAVerbatimX14FieldExtension) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  xml.append("<pivotFields count=\"1\"><pivotField axis=\"axisRow\" showAll=\"0\"><extLst>");
  xml.append("<ext uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\" xmlns:y=\"").append(kX14Ns);
  xml.append("\"><y:pivotField ignore=\"1\"/><y:other/></ext>");
  xml.append("</extLst></pivotField></pivotFields></pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  pivot::PivotTable table = std::move(table_or.value());
  ASSERT_FALSE(table.fields()[0].fill_down_labels);

  table.mutable_fields()[0].fill_down_labels = true;
  std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<y:pivotField ignore=\"1\" fillDownLabels=\"1\"/><y:other/></ext>"), std::string::npos)
      << written;
  auto again_or = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(again_or)) << again_or.error().message;
  EXPECT_TRUE(again_or.value().fields()[0].fill_down_labels);

  table.mutable_fields()[0].fill_down_labels = false;
  written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<y:pivotField ignore=\"1\"/><y:other/></ext>"), std::string::npos) << written;
}

TEST(PivotTableWriter, ItemCaptionAndCollapseAreWrittenAsNAndSd) {
  pivot::PivotTable table = TableWithTwoItemRowField();
  table.mutable_fields()[0].items[0].caption = "zz";
  table.mutable_fields()[0].items[1].show_detail = false;
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<item x=\"0\" n=\"zz\"/><item x=\"1\" sd=\"0\"/>"), std::string::npos) << xml;

  auto parsed_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << parsed_or.error().message;
  EXPECT_EQ(parsed_or.value().fields()[0].items[0].caption, "zz");
  EXPECT_FALSE(parsed_or.value().fields()[0].items[1].show_detail);
  EXPECT_TRUE(parsed_or.value().fields()[0].items[0].passthrough_attrs.empty());
}

TEST(PivotTableWriter, CollapsingAFieldRegeneratesItsDefaultMarker) {
  pivot::PivotTable table = ReadLoaded("<item t=\"default\" c=\"1\"/>");
  // Unchanged, the captured marker is kept with its other attributes.
  EXPECT_NE(write_pivot_table_definition(table).find("<item t=\"default\" c=\"1\"/></items>"), std::string::npos);

  pivot::PivotField& field = table.mutable_fields()[0];
  for (pivot::PivotItem& item : field.items) {
    item.show_detail = false;
  }
  field.marker_items[0].show_detail = false;
  const std::string xml = write_pivot_table_definition(table);
  // `sd` is patched into the captured marker, which keeps its other attributes.
  EXPECT_NE(xml.find("<items count=\"4\"><item x=\"2\" sd=\"0\"/><item x=\"0\" sd=\"0\"/><item x=\"1\" sd=\"0\"/>"
                     "<item t=\"default\" c=\"1\" sd=\"0\"/></items>"),
            std::string::npos)
      << xml;
}

TEST(PivotTableWriter, ExpandingACollapsedFieldDropsSdFromItsDefaultMarker) {
  pivot::PivotTable table = ReadLoaded("<item t=\"default\" sd=\"0\"/>");
  ASSERT_FALSE(table.fields()[0].marker_items[0].show_detail);
  EXPECT_NE(write_pivot_table_definition(table).find("<item t=\"default\" sd=\"0\"/></items>"), std::string::npos);

  table.mutable_fields()[0].marker_items[0].show_detail = true;
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<item x=\"1\"/><item t=\"default\"/></items>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("sd="), std::string::npos) << xml;
}

TEST(PivotTableWriter, ExpandingPatchesSdOutOfTheCapturedMarkerEitherQuoted) {
  pivot::PivotTable table = ReadLoaded("<item c=\"1\" sd=\"0\" t=\"default\"/>");
  table.mutable_fields()[0].marker_items[0].show_detail = true;
  EXPECT_NE(write_pivot_table_definition(table).find("<item c=\"1\" t=\"default\"/></items>"), std::string::npos);

  // A marker captured with single quotes is read and patched the same way.
  pivot::PivotMarkerItem& marker = table.mutable_fields()[0].marker_items[0];
  marker.xml = "<item t='default' sd='0' c='1'/>";
  marker.show_detail = false;
  EXPECT_NE(write_pivot_table_definition(table).find("<item t='default' sd='0' c='1'/></items>"), std::string::npos);
  marker.show_detail = true;
  EXPECT_NE(write_pivot_table_definition(table).find("<item t='default' c='1'/></items>"), std::string::npos);
}

TEST(PivotTableWriter, RegeneratedSubtotalsKeepTheCollapsedDefaultMarker) {
  // The stray `sum` marker no longer matches the selection, so the subtotal
  // markers are regenerated; the collapse state of `default` carries over.
  const pivot::PivotTable table = ReadLoaded("<item t=\"default\" sd=\"0\"/><item t=\"sum\"/>");
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<item x=\"1\"/><item t=\"default\" sd=\"0\"/></items>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("t=\"sum\""), std::string::npos) << xml;
}

TEST(PivotTableWriter, UntouchedLoadedTableKeepsItsRenderedItems) {
  const std::string xml = write_pivot_table_definition(ReadLoaded());
  EXPECT_NE(xml.find("<rowItems count=\"1\"><i><x/></i></rowItems>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<colItems count=\"1\"><i/></colItems><pageFields"), std::string::npos) << xml;
}

TEST(PivotTableWriter, DirtyTableDropsRowAndColItemsButKeepsPageFields) {
  pivot::PivotTable table = ReadLoaded();
  ASSERT_FALSE(table.layout_dirty());
  table.mark_layout_dirty();
  PivotCacheWritePlan plan;
  plan.refresh_on_load = true;
  const std::string xml = write_pivot_table_definition(table, std::nullopt, &plan);
  EXPECT_EQ(xml.find("<rowItems"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<colItems"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<pageFields count=\"1\"><pageField fld=\"1\" hier=\"-1\"/></pageFields>"), std::string::npos)
      << xml;
  EXPECT_NE(xml.find("</rowFields><pageFields"), std::string::npos) << xml;
}

TEST(PivotTableWriter, DirtyTableKeepsItsItemsWhenTheCacheIsNotRefreshedOnLoad) {
  // Nothing would redraw a report whose items were dropped, so they stay.
  pivot::PivotTable table = ReadLoaded();
  table.mark_layout_dirty();
  PivotCacheWritePlan plan;
  plan.refresh_on_load = false;
  for (const std::string& xml :
       {write_pivot_table_definition(table, std::nullopt, &plan), write_pivot_table_definition(table)}) {
    EXPECT_NE(xml.find("<rowItems count=\"1\"><i><x/></i></rowItems>"), std::string::npos) << xml;
    EXPECT_NE(xml.find("<colItems count=\"1\"><i/></colItems>"), std::string::npos) << xml;
  }
}

TEST(PivotTableWriter, ACloneKeepsTheDirtyFlag) {
  pivot::PivotTable table = ReadLoaded();
  table.mark_layout_dirty();
  EXPECT_TRUE(table.clone()->layout_dirty());
}

TEST(PivotTableWriter, GeneratedItemsReplaceTheReadOnesAtTheirSchemaSlots) {
  pivot::PivotTable table = ReadLoaded();
  table.mark_layout_dirty();
  PivotCacheWritePlan plan;
  PivotItemsXml items;
  items.row_items = "<rowItems count=\"2\"><i><x v=\"1\"/></i><i t=\"grand\"><x/></i></rowItems>";
  items.col_items = "<colItems count=\"1\"><i/></colItems>";
  // Whatever the plan says, the generated lines stand in for the read ones.
  for (const bool refresh : {true, false}) {
    plan.refresh_on_load = refresh;
    const std::string xml = write_pivot_table_definition(table, std::nullopt, &plan, &items);
    EXPECT_NE(xml.find("</rowFields>" + items.row_items), std::string::npos) << xml;
    EXPECT_NE(xml.find(items.col_items + "<pageFields"), std::string::npos) << xml;
    EXPECT_EQ(xml.find("<rowItems count=\"1\">"), std::string::npos) << xml;
    EXPECT_EQ(xml.find("<rowItems"), xml.rfind("<rowItems")) << xml;
    EXPECT_EQ(xml.find("<colItems"), xml.rfind("<colItems")) << xml;
  }
}

// The x14 field extension under an unusual shape is kept whole and never
// duplicated by a generated one.
pivot::PivotTable ReadFieldExtLst(std::string_view ext_xml) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
  xml.append("<pivotFields count=\"1\"><pivotField axis=\"axisRow\" showAll=\"0\"><extLst>").append(ext_xml);
  xml.append("</extLst></pivotField></pivotFields></pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  EXPECT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  return std::move(table_or.value());
}

std::size_t Occurrences(const std::string& haystack, std::string_view needle) {
  std::size_t count = 0;
  for (std::size_t at = haystack.find(needle); at != std::string::npos; at = haystack.find(needle, at + 1)) {
    ++count;
  }
  return count;
}

TEST(PivotTableWriter, X14FieldUnderAnotherPrefixIsModelled) {
  std::string ext = "<ext uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\" xmlns:y=\"";
  ext.append(kX14Ns).append("\"><y:pivotField fillDownLabels=\"1\"/></ext>");
  const pivot::PivotTable table = ReadFieldExtLst(ext);
  const pivot::PivotField& field = table.fields()[0];
  EXPECT_TRUE(field.fill_down_labels);
  EXPECT_TRUE(field.raw_ext_xml.empty());
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find("<x14:pivotField fillDownLabels=\"1\"/></ext></extLst>"), std::string::npos) << xml;
}

TEST(PivotTableWriter, X14FieldExtensionWithASiblingIsKeptWhole) {
  std::string ext = "<ext uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\" xmlns:x14=\"";
  ext.append(kX14Ns).append("\"><x14:pivotField fillDownLabels=\"1\"/><x14:other/></ext>");
  const pivot::PivotTable table = ReadFieldExtLst(ext);
  const pivot::PivotField& field = table.fields()[0];
  EXPECT_TRUE(field.fill_down_labels);
  EXPECT_TRUE(field.x14_attrs.empty());
  EXPECT_NE(field.raw_ext_xml.find("<x14:other/>"), std::string::npos);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_EQ(Occurrences(xml, "2946ED86"), 1U) << xml;
  EXPECT_NE(xml.find("<x14:pivotField fillDownLabels=\"1\"/><x14:other/></ext></extLst>"), std::string::npos) << xml;
}

TEST(PivotTableWriter, X14FieldExtensionWithoutItsElementIsKeptWhole) {
  const pivot::PivotTable table =
      ReadFieldExtLst("<ext uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\"><unknown a=\"1\"/></ext>");
  const pivot::PivotField& field = table.fields()[0];
  EXPECT_FALSE(field.fill_down_labels);
  const std::string xml = write_pivot_table_definition(table);
  EXPECT_EQ(Occurrences(xml, "2946ED86"), 1U) << xml;
  EXPECT_NE(xml.find("<unknown a=\"1\"/></ext></extLst>"), std::string::npos) << xml;
}

TEST(PivotTableWriter, DataOnRowsFollowsTheValuesMarkerAndDropsAStaleAttribute) {
  pivot::PivotTable table;
  table.mutable_passthrough_attrs().emplace_back("dataOnRows", "1");
  table.mutable_passthrough_attrs().emplace_back("dataPosition", "0");
  table.mutable_data_fields().resize(2);
  table.mutable_data_fields()[1].field_index = 1;

  table.set_row_values_position(0U);
  std::string xml = write_pivot_table_definition(table);
  EXPECT_NE(xml.find(" dataOnRows=\"1\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("dataPosition"), std::string::npos) << xml;

  // Moved to the columns, or fewer than two data fields: the attribute goes.
  table.set_row_values_position(std::nullopt);
  table.set_col_values_position(0U);
  EXPECT_EQ(write_pivot_table_definition(table).find("dataOnRows"), std::string::npos);
  table.set_col_values_position(std::nullopt);
  table.set_row_values_position(0U);
  table.mutable_data_fields().pop_back();
  EXPECT_EQ(write_pivot_table_definition(table).find("dataOnRows"), std::string::npos);
}

TEST(PivotTableReader, DataOnRowsAndDataPositionAreNotKeptAsPassthrough) {
  const std::string xml = std::string(kXmlDecl) + "<pivotTableDefinition" + std::string(kPivotNs) +
                          " name=\"P\" cacheId=\"1\" dataOnRows=\"1\" dataPosition=\"1\">"
                          "<location ref=\"A1:B2\"/></pivotTableDefinition>";
  auto parsed = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed)) << parsed.error().message;
  for (const auto& attr : parsed.value().passthrough_attrs()) {
    EXPECT_NE(attr.first, "dataOnRows");
    EXPECT_NE(attr.first, "dataPosition");
  }
}

// ---------------------------------------------------------------------------
// Table options
// ---------------------------------------------------------------------------

std::string OptionsXml(std::string_view root_attrs, std::string_view fields = "<pivotField axis=\"axisRow\"/>",
                       std::string_view tail = "") {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\"").append(root_attrs);
  xml.append("><location ref=\"A1:B2\"/><pivotFields count=\"1\">").append(fields);
  xml.append("</pivotFields><rowFields count=\"1\"><field x=\"0\"/></rowFields>").append(tail);
  xml.append("</pivotTableDefinition>");
  return xml;
}

struct OptionCase {
  const char* label;
  const char* attrs;
  bool (*check)(const pivot::PivotTable&);
};

TEST(PivotTableOptionsRoundTrip, EveryOptionReadsWritesAndReadsBack) {
  const OptionCase cases[] = {
      {"showError", " showError=\"1\" errorCaption=\"ERR\"",
       [](const pivot::PivotTable& t) { return t.show_error() && t.error_caption() == "ERR"; }},
      {"showMissing", " showMissing=\"0\" missingCaption=\"-\"",
       [](const pivot::PivotTable& t) { return !t.show_missing() && t.missing_caption() == "-"; }},
      {"showHeaders", " showHeaders=\"0\"", [](const pivot::PivotTable& t) { return !t.show_headers(); }},
      {"showDrill", " showDrill=\"0\"", [](const pivot::PivotTable& t) { return !t.show_drill(); }},
      {"printDrill", " printDrill=\"1\"", [](const pivot::PivotTable& t) { return t.print_drill(); }},
      {"showDataTips", " showDataTips=\"0\"", [](const pivot::PivotTable& t) { return !t.show_data_tips(); }},
      {"gridDropZones", " gridDropZones=\"1\"", [](const pivot::PivotTable& t) { return t.grid_drop_zones(); }},
      {"showEmptyRow", " showEmptyRow=\"1\"", [](const pivot::PivotTable& t) { return t.show_empty_row(); }},
      {"showEmptyCol", " showEmptyCol=\"1\"", [](const pivot::PivotTable& t) { return t.show_empty_col(); }},
      {"fieldPrintTitles", " fieldPrintTitles=\"1\"",
       [](const pivot::PivotTable& t) { return t.field_print_titles(); }},
      {"itemPrintTitles", " itemPrintTitles=\"1\"", [](const pivot::PivotTable& t) { return t.item_print_titles(); }},
      {"showDataDropDown", " showDataDropDown=\"0\"",
       [](const pivot::PivotTable& t) { return !t.show_data_drop_down(); }},
      {"multipleFieldFilters", " multipleFieldFilters=\"0\"",
       [](const pivot::PivotTable& t) { return !t.multiple_field_filters(); }},
      {"preserveFormatting", " preserveFormatting=\"0\"",
       [](const pivot::PivotTable& t) { return !t.preserve_formatting(); }},
      {"mergeItem", " mergeItem=\"1\"", [](const pivot::PivotTable& t) { return t.merge_item(); }},
      {"indent", " indent=\"2\"", [](const pivot::PivotTable& t) { return t.indent() == 2U; }},
      {"indentZero", " indent=\"0\"", [](const pivot::PivotTable& t) { return t.indent() == 0U; }},
      {"pageWrap", " pageWrap=\"2\" pageOverThenDown=\"1\"",
       [](const pivot::PivotTable& t) { return t.page_wrap() == 2U && t.page_over_then_down(); }},
      {"customListSort", " customListSort=\"0\"", [](const pivot::PivotTable& t) { return !t.custom_list_sort(); }},
  };
  for (const OptionCase& c : cases) {
    SCOPED_TRACE(c.label);
    auto first = read_pivot_table_definition(Bytes(OptionsXml(c.attrs)));
    ASSERT_TRUE(static_cast<bool>(first)) << first.error().message;
    EXPECT_TRUE(c.check(first.value()));
    for (const auto& attr : first.value().passthrough_attrs()) {
      EXPECT_EQ(std::string(c.attrs).find(" " + attr.first + "="), std::string::npos) << attr.first;
    }
    const std::string written = write_pivot_table_definition(first.value());
    // Each attribute is written verbatim, spelled as authored.
    std::string_view rest = c.attrs;
    while (!rest.empty()) {
      const std::size_t end = rest.find("\" ", 1);
      const std::string_view one = rest.substr(0, end == std::string_view::npos ? rest.size() : end + 1);
      EXPECT_NE(written.find(one), std::string::npos) << one << " in " << written;
      rest = end == std::string_view::npos ? std::string_view() : rest.substr(end + 1);
    }
    auto second = read_pivot_table_definition(Bytes(written));
    ASSERT_TRUE(static_cast<bool>(second)) << second.error().message;
    EXPECT_TRUE(c.check(second.value()));
    EXPECT_EQ(write_pivot_table_definition(second.value()), written);
  }
}

TEST(PivotTableOptionsRoundTrip, SchemaDefaultsAreReadAndLeftUnwritten) {
  auto table = read_pivot_table_definition(Bytes(OptionsXml("")));
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
  const pivot::PivotTable& t = table.value();
  EXPECT_FALSE(t.show_error());
  EXPECT_FALSE(t.error_caption().has_value());
  EXPECT_TRUE(t.show_missing());
  EXPECT_FALSE(t.missing_caption().has_value());
  EXPECT_TRUE(t.show_headers());
  EXPECT_TRUE(t.show_drill());
  EXPECT_FALSE(t.print_drill());
  EXPECT_TRUE(t.show_data_tips());
  EXPECT_FALSE(t.grid_drop_zones());
  EXPECT_FALSE(t.show_empty_row());
  EXPECT_FALSE(t.show_empty_col());
  EXPECT_FALSE(t.field_print_titles());
  EXPECT_FALSE(t.item_print_titles());
  EXPECT_TRUE(t.show_data_drop_down());
  EXPECT_TRUE(t.multiple_field_filters());
  EXPECT_TRUE(t.preserve_formatting());
  EXPECT_FALSE(t.merge_item());
  EXPECT_EQ(t.indent(), 1U);
  EXPECT_FALSE(t.page_over_then_down());
  EXPECT_EQ(t.page_wrap(), 0U);
  EXPECT_TRUE(t.custom_list_sort());
  EXPECT_TRUE(t.alt_text().empty());
  EXPECT_TRUE(t.alt_text_summary().empty());
  const std::string written = write_pivot_table_definition(t);
  for (const char* name :
       {"showError",          "errorCaption",     "showMissing",     "missingCaption",   "showHeaders",
        "showDrill",          "printDrill",       "showDataTips",    "gridDropZones",    "showEmptyRow",
        "showEmptyCol",       "fieldPrintTitles", "itemPrintTitles", "showDataDropDown", "multipleFieldFilters",
        "preserveFormatting", "mergeItem",        "indent",          "pageWrap",         "pageOverThenDown",
        "customListSort",     "asteriskTotals",   "altText"}) {
    EXPECT_EQ(written.find(std::string(" ") + name + "="), std::string::npos) << name << " in " << written;
  }
}

TEST(PivotTableOptionsRoundTrip, NewPivotDefaultsWriteExcelsCreationAttributes) {
  pivot::PivotTable table;
  table.apply_new_pivot_defaults();
  EXPECT_TRUE(table.item_print_titles());
  EXPECT_EQ(table.indent(), 0U);
  EXPECT_FALSE(table.multiple_field_filters());
  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find(" itemPrintTitles=\"1\""), std::string::npos) << written;
  EXPECT_NE(written.find(" indent=\"0\""), std::string::npos) << written;
  EXPECT_NE(written.find(" multipleFieldFilters=\"0\""), std::string::npos) << written;
}

TEST(PivotTableOptionsRoundTrip, AltTextLivesOnTheX14Extension) {
  const std::string ext =
      "<extLst><ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotTableDefinition "
      "altText=\"Alt title\" altTextSummary=\"Alt summary text\" hideValuesRow=\"1\" "
      "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/></ext></extLst>";
  auto table = read_pivot_table_definition(Bytes(OptionsXml("", "<pivotField axis=\"axisRow\"/>", ext)));
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
  EXPECT_EQ(table.value().alt_text(), "Alt title");
  EXPECT_EQ(table.value().alt_text_summary(), "Alt summary text");
  EXPECT_EQ(table.value().raw_passthrough_xml().find("altText"), std::string::npos);
  EXPECT_NE(table.value().raw_passthrough_xml().find("hideValuesRow=\"1\""), std::string::npos);

  const std::string written = write_pivot_table_definition(table.value());
  EXPECT_NE(written.find("altText=\"Alt title\" altTextSummary=\"Alt summary text\""), std::string::npos) << written;
  EXPECT_EQ(written.find("altText"), written.rfind("altText=\"Alt title\"")) << written;
  EXPECT_NE(written.find("hideValuesRow=\"1\""), std::string::npos) << written;
  const std::string root_tag = written.substr(0, written.find("<location"));
  EXPECT_EQ(root_tag.find("altText"), std::string::npos) << root_tag;

  auto again = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(again)) << again.error().message;
  EXPECT_EQ(again.value().alt_text(), "Alt title");
  EXPECT_EQ(again.value().alt_text_summary(), "Alt summary text");
}

TEST(PivotTableOptionsRoundTrip, AltTextOnATableWithoutAnExtensionGeneratesIt) {
  pivot::PivotTable table;
  table.set_alt_text("Title");
  table.set_alt_text_summary("Summary");
  const std::string written = write_pivot_table_definition(table);
  EXPECT_NE(written.find("<x14:pivotTableDefinition altText=\"Title\" altTextSummary=\"Summary\" hideValuesRow=\"1\""),
            std::string::npos)
      << written;
  auto parsed = read_pivot_table_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(parsed)) << parsed.error().message;
  EXPECT_EQ(parsed.value().alt_text(), "Title");
  EXPECT_EQ(parsed.value().alt_text_summary(), "Summary");
}

TEST(PivotTableOptionsRoundTrip, AltTextIsAddedToAnExtLstWithoutTheTableExtension) {
  const std::string other =
      "<extLst><ext uri=\"{AAAAAAAA-0000-0000-0000-000000000000}\" "
      "xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:foo/></ext></extLst>";
  auto table =
      read_pivot_table_definition(Bytes(OptionsXml(" updatedVersion=\"8\"", "<pivotField axis=\"axisRow\"/>", other)));
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
  table.value().set_alt_text("T");
  const std::string written = write_pivot_table_definition(table.value());
  EXPECT_NE(written.find("<x15:foo/>"), std::string::npos) << written;
  EXPECT_NE(written.find("<x14:pivotTableDefinition altText=\"T\""), std::string::npos) << written;
  EXPECT_EQ(written.find("<extLst"), written.rfind("<extLst")) << written;
}

TEST(PivotTableOptionsRoundTrip, RootAttributesFollowExcelsOrder) {
  pivot::PivotTable table;
  table.set_auto_formatting(true);
  table.set_item_print_titles(true);
  table.set_merge_item(true);
  table.set_indent(0U);
  table.set_multiple_field_filters(false);
  table.mutable_passthrough_attrs().emplace_back("createdVersion", "8");
  const std::string written = write_pivot_table_definition(table);
  const std::size_t drill = written.find(" useAutoFormatting=");
  const std::size_t print = written.find(" itemPrintTitles=");
  const std::size_t merge = written.find(" mergeItem=");
  const std::size_t created = written.find(" createdVersion=");
  const std::size_t indent = written.find(" indent=");
  const std::size_t multi = written.find(" multipleFieldFilters=");
  ASSERT_NE(drill, std::string::npos) << written;
  EXPECT_LT(drill, print);
  EXPECT_LT(print, merge);
  EXPECT_LT(merge, created);
  EXPECT_LT(created, indent);
  EXPECT_LT(indent, multi);
}

// ---------------------------------------------------------------------------
// pivotField@showAll
// ---------------------------------------------------------------------------

TEST(PivotFieldShowAll, AttributeStatesReadAndWrite) {
  struct Case {
    const char* attr;
    bool show_all;
    bool written;  // whether the writer spells showAll="0"
  };
  const Case cases[] = {{"", true, false}, {" showAll=\"1\"", true, false}, {" showAll=\"0\"", false, true}};
  for (const Case& c : cases) {
    SCOPED_TRACE(c.attr);
    const std::string field = std::string("<pivotField axis=\"axisRow\"") + c.attr + "/>";
    auto table = read_pivot_table_definition(Bytes(OptionsXml("", field)));
    ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
    ASSERT_EQ(table.value().fields().size(), 1U);
    EXPECT_EQ(table.value().fields()[0].show_all, c.show_all);
    for (const auto& attr : table.value().fields()[0].passthrough_attrs) {
      EXPECT_NE(attr.first, "showAll");
    }
    const std::string out = write_pivot_table_definition(table.value());
    EXPECT_EQ(out.find("showAll=\"0\"") != std::string::npos, c.written) << out;
    EXPECT_EQ(out.find("showAll=\"1\""), std::string::npos) << out;
    auto again = read_pivot_table_definition(Bytes(out));
    ASSERT_TRUE(static_cast<bool>(again)) << again.error().message;
    EXPECT_EQ(again.value().fields()[0].show_all, c.show_all);
  }
}

TEST(PivotFieldShowAll, FieldCreatedInMemoryStartsOffAndWritesZero) {
  pivot::PivotField field;
  field.axis = pivot::PivotAxis::None;
  EXPECT_FALSE(field.show_all);
  pivot::PivotTable table;
  table.mutable_fields().push_back(field);
  EXPECT_NE(write_pivot_table_definition(table).find("<pivotField showAll=\"0\"/>"), std::string::npos);
  table.mutable_fields()[0].show_all = true;
  EXPECT_EQ(write_pivot_table_definition(table).find("showAll"), std::string::npos);
}

// A calculated field as Excel saves it (pivot_calc/cf1_price_amt_per_qty.xlsx): the captured drag*
// attributes sit between the form attributes and showAll, ahead of defaultSubtotal.
TEST(PivotFieldShowAll, CapturedAttributesKeepSchemaPosition) {
  const std::string field =
      "<pivotField dataField=\"1\" dragToRow=\"0\" dragToCol=\"0\" dragToPage=\"0\" showAll=\"0\" "
      "defaultSubtotal=\"0\"/>";
  const std::string xml =
      OptionsXml("", field, "<dataFields count=\"1\"><dataField name=\"X\" fld=\"0\"/></dataFields>");
  auto table = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
  const std::string written = write_pivot_table_definition(table.value());
  EXPECT_NE(written.find(field), std::string::npos) << written;
}

// ---------------------------------------------------------------------------
// item@f
// ---------------------------------------------------------------------------

TEST(PivotItemCalculatedFlag, ReadsWritesAndRoundTrips) {
  const std::string field =
      "<pivotField axis=\"axisRow\"><items count=\"3\"><item x=\"0\"/><item f=\"1\" x=\"1\"/><item t=\"default\"/>"
      "</items></pivotField>";
  auto table = read_pivot_table_definition(Bytes(OptionsXml("", field)));
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
  const auto& items = table.value().fields()[0].items;
  ASSERT_EQ(items.size(), 2U);
  EXPECT_FALSE(items[0].calculated);
  EXPECT_TRUE(items[1].calculated);
  for (const auto& attr : items[1].passthrough_attrs) {
    EXPECT_NE(attr.first, "f");
  }
  const std::string out = write_pivot_table_definition(table.value());
  EXPECT_NE(out.find("<item x=\"0\"/>"), std::string::npos) << out;
  EXPECT_NE(out.find("<item f=\"1\" x=\"1\"/>"), std::string::npos) << out;
  auto again = read_pivot_table_definition(Bytes(out));
  ASSERT_TRUE(static_cast<bool>(again)) << again.error().message;
  EXPECT_TRUE(again.value().fields()[0].items[1].calculated);
  EXPECT_FALSE(again.value().fields()[0].items[0].calculated);
}

// ---------------------------------------------------------------------------
// clone()
// ---------------------------------------------------------------------------

TEST(PivotTableClone, CopiesEveryOptionAndProjectionMember) {
  pivot::PivotTable t;
  t.set_show_error(true);
  t.set_error_caption(std::string("E"));
  t.set_show_missing(false);
  t.set_missing_caption(std::string("M"));
  t.set_show_headers(false);
  t.set_show_drill(false);
  t.set_print_drill(true);
  t.set_show_data_tips(false);
  t.set_grid_drop_zones(true);
  t.set_show_empty_row(true);
  t.set_show_empty_col(true);
  t.set_field_print_titles(true);
  t.set_item_print_titles(true);
  t.set_show_data_drop_down(false);
  t.set_multiple_field_filters(false);
  t.set_preserve_formatting(false);
  t.set_merge_item(true);
  t.set_indent(3U);
  t.set_page_over_then_down(true);
  t.set_page_wrap(4U);
  t.set_custom_list_sort(false);
  t.set_alt_text("A");
  t.set_alt_text_summary("S");
  t.set_projected_extent(pivot::PivotRect{1, 2, 3, 4});
  t.mutable_projected_merges().push_back(pivot::PivotRect{5, 6, 7, 8});
  t.set_projection_pending(true);

  const std::unique_ptr<pivot::PivotTable> c = t.clone();
  EXPECT_TRUE(c->show_error());
  EXPECT_EQ(c->error_caption(), "E");
  EXPECT_FALSE(c->show_missing());
  EXPECT_EQ(c->missing_caption(), "M");
  EXPECT_FALSE(c->show_headers());
  EXPECT_FALSE(c->show_drill());
  EXPECT_TRUE(c->print_drill());
  EXPECT_FALSE(c->show_data_tips());
  EXPECT_TRUE(c->grid_drop_zones());
  EXPECT_TRUE(c->show_empty_row());
  EXPECT_TRUE(c->show_empty_col());
  EXPECT_TRUE(c->field_print_titles());
  EXPECT_TRUE(c->item_print_titles());
  EXPECT_FALSE(c->show_data_drop_down());
  EXPECT_FALSE(c->multiple_field_filters());
  EXPECT_FALSE(c->preserve_formatting());
  EXPECT_TRUE(c->merge_item());
  EXPECT_EQ(c->indent(), 3U);
  EXPECT_TRUE(c->page_over_then_down());
  EXPECT_EQ(c->page_wrap(), 4U);
  EXPECT_FALSE(c->custom_list_sort());
  EXPECT_EQ(c->alt_text(), "A");
  EXPECT_EQ(c->alt_text_summary(), "S");
  ASSERT_TRUE(c->projected_extent().has_value());
  EXPECT_TRUE(*c->projected_extent() == (pivot::PivotRect{1, 2, 3, 4}));
  ASSERT_EQ(c->projected_merges().size(), 1U);
  EXPECT_TRUE(c->projected_merges()[0] == (pivot::PivotRect{5, 6, 7, 8}));
  EXPECT_TRUE(c->projection_pending());

  t.set_indent(9U);
  t.mutable_projected_merges().clear();
  EXPECT_EQ(c->indent(), 3U);
  EXPECT_EQ(c->projected_merges().size(), 1U);
}

// ---------------------------------------------------------------------------
// `<location>` offsets derived from the report options
//
// A table whose offsets were not read with it gets them from its shape. Each
// case drops the offsets Excel wrote and requires the derived ones to equal
// them.
// ---------------------------------------------------------------------------

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

std::vector<std::uint8_t> ReadFixtureBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

void AppendLocationAttr(std::string& out, const char* name, std::optional<std::uint32_t> value) {
  if (value.has_value()) {
    out.append(" ").append(name).append("=\"").append(std::to_string(*value)).append("\"");
  }
}

/// The offset attributes of the written `<location>`, after its `ref`.
std::string WrittenLocationOffsets(const std::string& xml) {
  const std::size_t at = xml.find("<location ref=\"");
  if (at == std::string::npos) {
    return "(no location)";
  }
  const std::size_t ref_end = xml.find('"', at + 15U);
  const std::size_t end = xml.find("/>", at);
  return ref_end == std::string::npos || end == std::string::npos ? "(malformed)"
                                                                  : xml.substr(ref_end + 1U, end - ref_end - 1U);
}

class PivotTableWriterLocation : public ::testing::TestWithParam<const char*> {};

TEST_P(PivotTableWriterLocation, DerivedOffsetsMatchTheOnesExcelSaved) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_options/" + GetParam();
  const std::vector<std::uint8_t> bytes = ReadFixtureBytes(path);
  ASSERT_FALSE(bytes.empty()) << "could not read " << path;
  auto read_or = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  Workbook& wb = read_or.value().workbook;
  std::size_t tables = 0;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const std::unique_ptr<pivot::PivotTable>& table : wb.sheet(s).mutable_pivot_tables()) {
      std::string excel;
      AppendLocationAttr(excel, "firstHeaderRow", table->location_first_header_row());
      AppendLocationAttr(excel, "firstDataRow", table->location_first_data_row());
      AppendLocationAttr(excel, "firstDataCol", table->location_first_data_col());
      AppendLocationAttr(excel, "rowPageCount", table->location_row_page_count());
      AppendLocationAttr(excel, "colPageCount", table->location_col_page_count());
      table->set_location_attributes(std::nullopt, std::nullopt, std::nullopt, std::nullopt, std::nullopt);
      EXPECT_EQ(WrittenLocationOffsets(write_pivot_table_definition(*table)), excel) << table->name();
      ++tables;
    }
  }
  EXPECT_GT(tables, 0U);
}

std::string LocationFixtureName(const ::testing::TestParamInfo<const char*>& info) {
  std::string name;
  for (const char* p = info.param; *p != '\0' && *p != '.'; ++p) {
    name.push_back(*p);
  }
  return name;
}

INSTANTIATE_TEST_SUITE_P(ExcelSaved, PivotTableWriterLocation,
                         ::testing::Values("op_headers_off.xlsx", "op_griddz.xlsx", "op_pagewrap2_down_over.xlsx",
                                           "op_pagewrap2_over_down.xlsx", "op_merge_tabular.xlsx",
                                           "op_merge_outline.xlsx", "op_emptyrow_show_all.xlsx",
                                           "op_emptycol_show_all.xlsx", "op_indent3.xlsx", "op_missing_caption.xlsx"),
                         LocationFixtureName);

TEST(PivotTableWriter, HiddenFieldCaptionsDropTheColumnLabelsRow) {
  // Two data fields beside a column field: without captions the
  // "Column Labels" row goes, and both offsets move up one row.
  pivot::PivotTable table;
  table.mutable_fields().resize(3);
  table.mutable_fields()[0].axis = pivot::PivotAxis::Row;
  table.mutable_fields()[1].axis = pivot::PivotAxis::Col;
  table.mutable_fields()[2].axis = pivot::PivotAxis::Value;
  table.mutable_row_field_order() = {0U};
  table.mutable_col_field_order() = {1U};
  table.mutable_data_fields().resize(2);
  table.mutable_data_fields()[0].field_index = 2U;
  table.mutable_data_fields()[1].field_index = 2U;
  table.set_col_values_position(std::optional<std::size_t>(1U));
  EXPECT_EQ(WrittenLocationOffsets(write_pivot_table_definition(table)),
            " firstHeaderRow=\"1\" firstDataRow=\"3\" firstDataCol=\"1\"");
  table.set_show_headers(false);
  EXPECT_EQ(WrittenLocationOffsets(write_pivot_table_definition(table)),
            " firstHeaderRow=\"0\" firstDataRow=\"2\" firstDataCol=\"1\"");
}

}  // namespace
}  // namespace formulon::io
